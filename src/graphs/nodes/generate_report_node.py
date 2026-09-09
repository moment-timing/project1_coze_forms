import os
import time
import math
from typing import Dict, List, Any, Optional

from langchain_core.runnables import RunnableConfig
from langgraph.runtime import Runtime
from coze_coding_utils.runtime_ctx.context import Context
from coze_coding_dev_sdk.s3 import S3SyncStorage

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.marker import Marker

from graphs.state import GenerateReportInput, GenerateReportOutput, CategoryData

HEADER_FILL = "4472C4"
TITLE_FILL = "1F4E79"
SUBTOTAL_FILL = "DDEBF7"
TOTAL_FILL = "FFF2CC"
THIN_SIDE = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_CENTER = Alignment(horizontal="left", vertical="center", wrap_text=True)

# 表格行高 / 列宽(自适应，避免数据被遮盖)
HEADER_ROW_H = 24
BODY_ROW_H = 22
TITLE_ROW_H = 32
COL_WIDTHS = {"A": 18, "B": 12, "C": 58}
YEAR_COL_W = 16
GROWTH_COL_W = 13


def _gmv_for(agg_year: List[Dict[str, Any]], year: str) -> float:
    """返回某类目在指定年份的销售额(万元)，不存在返回 0。"""
    for ay in agg_year:
        if not isinstance(ay, dict):
            continue
        if str(ay.get("年")) == str(year):
            v = ay.get("销售额(元)")
            return (float(v) / 10000.0) if v is not None else 0.0
    return 0.0


def _calc_growth(last: float, prev: float) -> Optional[float]:
    """计算较上年增幅(百分数)，上年为0或缺失返回 None。"""
    if prev and abs(prev) > 1e-9:
        return round((last - prev) / prev * 100.0, 2)
    return None


def _build_monthly_by_year_wan(monthly: List[Dict[str, Any]], years: List[str]) -> Dict[str, List[float]]:
    """将全局月度汇总按年份组织为 [1..12] 每月的销售额列表，单位换算为万元。"""
    by_year: Dict[str, List[float]] = {y: [0.0] * 12 for y in years}
    for m in monthly:
        if not isinstance(m, dict):
            continue
        ym = str(m.get("月", ""))
        if len(ym) < 6:
            continue
        y = ym[:4]
        try:
            month = int(ym[4:6])
        except ValueError:
            continue
        if y in by_year and 1 <= month <= 12:
            v = m.get("销售额(元)")
            by_year[y][month - 1] = (float(v) / 10000.0) if v is not None else 0.0
    return by_year


def _apply_border(ws, r1: int, r2: int, c1: int, c2: int) -> None:
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = BORDER


def _nice_step(maxv: float) -> float:
    """根据最大值计算一个清晰的 Y 轴刻度步长。"""
    if maxv <= 0:
        return 1.0
    raw = maxv / 5.0
    if raw <= 0:
        return 1.0
    mag = 10 ** int(math.floor(math.log10(raw)))
    for m in (1, 2, 2.5, 5, 10):
        if raw <= m * mag:
            return m * mag
    return 10 * mag


def _ceil_to(value: float, step: float) -> float:
    """将 value 向上取整到 step 的整数倍。"""
    if step <= 0:
        return value
    return math.ceil(value / step) * step


def _build_source_block(ws, years: List[str], monthly: List[Dict[str, Any]], src_col: int) -> None:
    """在指定起始列写入折线图的数据源(单位为万元)，月份为01-12。"""
    ws.cell(1, src_col, "月份")
    for i, y in enumerate(years):
        ws.cell(1, src_col + 1 + i, y)
    by_year = _build_monthly_by_year_wan(monthly, years)
    for m in range(1, 13):
        ws.cell(1 + m, src_col, f"{m:02d}")
        for i, y in enumerate(years):
            ws.cell(1 + m, src_col + 1 + i, round(by_year[y][m - 1], 2))


def _build_table(ws, shop_name: str, stat_time: str, years: List[str], categories: List[CategoryData]) -> int:
    """构建“市场体量”数据表，返回表体最后一行行号。"""
    n_cols = 3 + len(years) + 1
    last_col = get_column_letter(n_cols)

    # 标题行
    ws.cell(1, 1, f"{shop_name}电商市场分析报表")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    for c in range(1, n_cols + 1):
        ws.cell(1, c).fill = PatternFill("solid", fgColor=TITLE_FILL)
    title_cell = ws.cell(1, 1)
    title_cell.font = Font(size=16, bold=True, color="FFFFFF")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = TITLE_ROW_H

    # 统计信息行
    ws.cell(2, 1, f"统计时间：{stat_time}    |    数据来源：各平台品类数据")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    info_cell = ws.cell(2, 1)
    info_cell.font = Font(size=9, color="808080")
    info_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 18

    # 表头行
    headers: List[str] = ["目标产品", "平台", "类目"] + [f"{y}年GMV(万元)" for y in years] + [f"{years[-1]}年增幅"]
    header_row = 3
    for j, h in enumerate(headers, start=1):
        cell = ws.cell(header_row, j, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
        cell.alignment = CENTER
    ws.row_dimensions[header_row].height = HEADER_ROW_H

    first_body_row = header_row + 1
    row = first_body_row

    # 按平台分组(保持数据顺序)
    platform_order: List[str] = []
    plat_map: Dict[str, List[CategoryData]] = {}
    for c in categories:
        if c.platform not in plat_map:
            plat_map[c.platform] = []
            platform_order.append(c.platform)
        plat_map[c.platform].append(c)

    total_gmv: List[float] = [0.0] * len(years)

    for plat in platform_order:
        cats = plat_map[plat]
        block_start = row
        plat_gmv: List[float] = [0.0] * len(years)
        for cat in cats:
            ws.cell(row, 1)
            ws.cell(row, 2)
            cat_cell = ws.cell(row, 3, cat.category_name)
            cat_cell.alignment = LEFT_CENTER
            for i, y in enumerate(years):
                v = _gmv_for(cat.agg_year, y)
                cell = ws.cell(row, 4 + i, round(v, 1))
                cell.number_format = "#,##0.0"
                plat_gmv[i] += v
                total_gmv[i] += v
            last_col_idx = 3 + len(years) + 1
            g = _calc_growth(_gmv_for(cat.agg_year, years[-1]), _gmv_for(cat.agg_year, years[-2]))
            cell = ws.cell(row, last_col_idx, g)
            cell.number_format = '0.00"%"'
            if g is not None and g < 0:
                cell.font = Font(color="C00000")
            ws.row_dimensions[row].height = BODY_ROW_H
            row += 1

        # 平台小计行
        ws.cell(row, 1)
        ws.cell(row, 2)
        sub_cell = ws.cell(row, 3, f"{plat}小计")
        sub_cell.font = Font(bold=True)
        sub_cell.alignment = CENTER
        for i in range(len(years)):
            cell = ws.cell(row, 4 + i, round(plat_gmv[i], 1))
            cell.number_format = "#,##0.0"
            cell.font = Font(bold=True)
        last_col_idx = 3 + len(years) + 1
        g = _calc_growth(plat_gmv[-1], plat_gmv[-2])
        cell = ws.cell(row, last_col_idx, g)
        cell.number_format = '0.00"%"'
        cell.font = Font(bold=True)
        if g is not None and g < 0:
            cell.font = Font(bold=True, color="C00000")
        for c in range(1, n_cols + 1):
            ws.cell(row, c).fill = PatternFill("solid", fgColor=SUBTOTAL_FILL)
        ws.row_dimensions[row].height = BODY_ROW_H
        row += 1

        # 合并平台列(类目行 + 小计行)
        ws.merge_cells(start_row=block_start, start_column=2, end_row=row - 1, end_column=2)
        plat_cell = ws.cell(block_start, 2, plat)
        plat_cell.alignment = Alignment(horizontal="center", vertical="center")

    # 总计行
    tc = ws.cell(row, 3, "/")
    tc.font = Font(bold=True)
    for c in (1, 2):
        ws.cell(row, c).font = Font(bold=True)
    for i in range(len(years)):
        cell = ws.cell(row, 4 + i, round(total_gmv[i], 1))
        cell.number_format = "#,##0.0"
        cell.font = Font(bold=True)
    last_col_idx = 3 + len(years) + 1
    g = _calc_growth(total_gmv[-1], total_gmv[-2])
    cell = ws.cell(row, last_col_idx, g)
    cell.number_format = '0.00"%"'
    cell.font = Font(bold=True)
    if g is not None and g < 0:
        cell.font = Font(bold=True, color="C00000")
    for c in range(1, n_cols + 1):
        ws.cell(row, c).fill = PatternFill("solid", fgColor=TOTAL_FILL)
    ws.row_dimensions[row].height = BODY_ROW_H
    total_row = row
    row += 1

    # 合并目标产品列(覆盖全部数据行含总计)
    ws.merge_cells(start_row=first_body_row, start_column=1, end_row=total_row, end_column=1)
    shop_cell = ws.cell(first_body_row, 1, shop_name)
    shop_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 边框
    if total_row >= header_row:
        _apply_border(ws, header_row, total_row, 1, n_cols)

    # 列宽(自适应，避免数据被遮盖)
    ws.column_dimensions["A"].width = COL_WIDTHS["A"]
    ws.column_dimensions["B"].width = COL_WIDTHS["B"]
    ws.column_dimensions["C"].width = COL_WIDTHS["C"]
    for i in range(len(years)):
        ws.column_dimensions[get_column_letter(4 + i)].width = YEAR_COL_W
    ws.column_dimensions[last_col].width = GROWTH_COL_W

    # 底部说明
    note_row = total_row + 2
    notes = [
        f"数据来源：{shop_name} 相关类目各平台品类数据",
        "单位：GMV 数值为万元；增幅为较上年度的增长率。",
    ]
    for n in notes:
        cell = ws.cell(note_row, 1, n)
        cell.font = Font(size=9, color="808080")
        ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=n_cols)
        note_row += 1

    return note_row - 1


def _build_chart(ws, shop_name: str, years: List[str], monthly: List[Dict[str, Any]],
                 table_last_row: int, n_cols: int, src_col: int) -> None:
    """在表格下方构建折线图：x轴为01-12月份，每条折线代表一个年份(单位万元)，刻度清晰标注。"""
    # 写入图表数据源(万元)，并隐藏
    _build_source_block(ws, years, monthly, src_col)
    for i in range(len(years) + 1):
        ws.column_dimensions[get_column_letter(src_col + i)].hidden = True

    chart = LineChart()
    chart.title = f"{shop_name} 月度销售额趋势(万元)"
    chart.style = 13
    # 图表与表格同宽：按表格各列宽估算像素再换算为厘米
    col_chars = COL_WIDTHS["A"] + COL_WIDTHS["B"] + COL_WIDTHS["C"] + YEAR_COL_W * len(years) + GROWTH_COL_W
    chart.width = int(col_chars * 7.0 / 96.0 * 2.54)
    chart.height = 12

    data = Reference(ws, min_col=src_col + 1, min_row=1, max_col=src_col + len(years), max_row=13)
    cats = Reference(ws, min_col=src_col, min_row=2, max_row=13)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    # 关键修复：数据源位于隐藏列，若保持“仅绘制可见单元格”会导致图表空白
    chart.visible_cells_only = False

    # 线条样式：每条折线一种颜色 + 圆形数据点标记(蓝/橙/绿/红)
    _line_colors = ("4472C4", "ED7D31", "70AD47", "C00000")
    for idx, ser in enumerate(chart.ser):
        color = _line_colors[idx % len(_line_colors)]
        ser.graphicalProperties.line.solidFill = color
        ser.graphicalProperties.line.width = 20000  # 约2pt
        mk = Marker(symbol="circle", size=7)
        mk.graphicalProperties.solidFill = color
        mk.graphicalProperties.line.solidFill = color
        ser.marker = mk

    # 坐标轴刻度清晰标注
    chart.x_axis.majorTickMark = "out"
    chart.y_axis.majorTickMark = "out"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.title = "销售额(万元)"
    chart.x_axis.title = "月份"
    chart.y_axis.numFmt = '#,##0"万"'
    # 浅色网格线(横向+纵向)，方便读数
    chart.x_axis.majorGridlines = ChartLines()
    chart.y_axis.majorGridlines = ChartLines()

    # 统计最大值以获得清晰Y轴刻度(从0开始，向上取整)
    maxv = 1.0
    by_year = _build_monthly_by_year_wan(monthly, years)
    for y in years:
        m = max(by_year[y]) if by_year[y] else 0.0
        if m > maxv:
            maxv = m
    _step = _nice_step(maxv)
    chart.y_axis.scaling.min = 0
    chart.y_axis.majorUnit = _step
    chart.y_axis.scaling.max = _ceil_to(maxv, _step)

    # 数据点数值标注(万元)，刻度标清楚
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    chart.dataLabels.numFmt = '#,##0"万"'
    chart.dataLabels.dLblPos = "t"  # 数值显示在数据点上方

    chart.legend.position = "b"
    anchor_row = table_last_row + 3
    ws.add_chart(chart, f"A{anchor_row}")


def generate_report_node(
    state: GenerateReportInput,
    config: RunnableConfig,
    runtime: Runtime[Context],
) -> GenerateReportOutput:
    """
    title: 生成Excel分析报表
    desc: 构建“市场体量”表(含合并单元格、万元单位)与下方同宽“市场趋势”折线图(mysql网格刻度清晰)，生成Excel并上传对象存储
    integrations: 对象存储
    """
    ctx = runtime.context
    shop_name = state.shop_name or "分析报表"
    years = list(state.years)
    if not years:
        raise ValueError("缺少可用的年份数据")

    wb = Workbook()
    ws = wb.active
    ws.title = "市场体量"

    n_cols = 3 + len(years) + 1
    table_last_row = _build_table(ws, shop_name, state.stat_time, years, state.categories)

    src_col = n_cols + 3  # 数据源放在表格右侧空列，随后隐藏
    _build_chart(ws, shop_name, years, state.monthly_summary, table_last_row, n_cols, src_col)

    local_dir = "/tmp"
    fname = f"sales_analysis_report_{int(time.time())}.xlsx"
    local_path = os.path.join(local_dir, fname)
    wb.save(local_path)

    with open(local_path, "rb") as f:
        content = f.read()

    storage = S3SyncStorage(
        endpoint_url=os.getenv("COZE_BUCKET_ENDPOINT_URL"),
        access_key="",
        secret_key="",
        bucket_name=os.getenv("COZE_BUCKET_NAME"),
        region="cn-beijing",
    )
    key = storage.upload_file(
        file_content=content,
        file_name=f"report/{fname}",
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    url = storage.generate_presigned_url(key=key, expire_time=2592000)

    return GenerateReportOutput(report_key=key, report_url=url)