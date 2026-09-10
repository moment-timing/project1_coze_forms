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
from openpyxl.chart import LineChart, PieChart, Reference
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

# -------------------------------- 折线关闭圆滑 ---------------------------------------------------
  
def _build_raw_and_chart_data(monthly: List[Dict[str, Any]], years: List[str]):
    # 原始数据，全部是数字，用于Y轴max计算
    by_year_raw: Dict[str, List[float]] = {y: [0.0] * 12 for y in years}
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
        if y in by_year_raw and 1 <= month <= 12:
            v = m.get("销售额(元)")
            by_year_raw[y][month - 1] = (float(v) / 10000.0) if v is not None else 0.0

    # 生成图表数据：最后>0点之后全部置None，用于截断折线
    by_year_chart: Dict[str, List[Optional[float]]] = {}
    for year, data_list in by_year_raw.items():
        chart_list: List[Optional[float]] = data_list.copy()
        last_valid_idx = -1
        for idx, num in enumerate(chart_list):
            if num > 0:
                last_valid_idx = idx
        if last_valid_idx != -1:
            for idx in range(last_valid_idx + 1, 12):
                chart_list[idx] = None
        by_year_chart[year] = chart_list
    return by_year_raw, by_year_chart



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

# ------------------------- 折线关闭圆滑 ------------------------------------------
# def _build_source_block(ws, years: List[str], monthly: List[Dict[str, Any]], src_col: int) -> None:
#     """在指定起始列写入折线图的数据源(单位为万元)，月份为01-12。"""
#     ws.cell(1, src_col, "月份")
#     for i, y in enumerate(years):
#         ws.cell(1, src_col + 1 + i, y)
#     by_year = _build_monthly_by_year_wan(monthly, years)
#     for m in range(1, 13):
#         ws.cell(1 + m, src_col, f"{m:02d}")
#         for i, y in enumerate(years):
#             ws.cell(1 + m, src_col + 1 + i, round(by_year[y][m - 1], 2))

def _build_source_block(ws, years: List[str], monthly: List[Dict[str, Any]], src_col: int) -> Dict[str, List[float]]:
    """写入折线图数据源，返回原始by_year_raw（用于坐标轴计算）"""
    ws.cell(1, src_col, "月份")
    for i, y in enumerate(years):
        ws.cell(1, src_col + 1 + i, y)
    by_year_raw, by_year_chart = _build_raw_and_chart_data(monthly, years)

    for m in range(1, 13):
        ws.cell(1 + m, src_col, f"{m:02d}")
        for i, y in enumerate(years):
            val = by_year_chart[y][m - 1]
            if val is not None:
                ws.cell(1 + m, src_col + 1 + i, round(val, 2))
            else:
                ws.cell(1 + m, src_col + 1 + i, None)
    return by_year_raw


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
    # _build_source_block(ws, years, monthly, src_col)
    # ------------------------------ 关闭折线圆滑  ----------------------------------------------------------------
    by_year_raw = _build_source_block(ws, years, monthly, src_col)

    for i in range(len(years) + 1):
        ws.column_dimensions[get_column_letter(src_col + i)].hidden = True

    chart = LineChart()
    chart.title = f"{shop_name} 月度销售额趋势(万元)"
    chart.style = 13
    # 图表与表格同宽：按表格各列宽估算像素再换算为厘米
    col_chars = COL_WIDTHS["A"] + COL_WIDTHS["B"] + COL_WIDTHS["C"] + YEAR_COL_W * len(years) + GROWTH_COL_W
    chart.width = int(col_chars * 7.0 / 96.0 * 2.54)
    chart.height = 14  # 加高图表

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
      # -------------------------------------  关闭折线圆滑  -----------------------------------
        ser.smooth = False  # 新增：直线折线，取消弯曲平滑
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
    # maxv = 1.0
    # by_year = _build_monthly_by_year_wan(monthly, years)
    # for y in years:
    #     m = max(by_year[y]) if by_year[y] else 0.0
    #     if m > maxv:
    #         maxv = m
    #     _step = _nice_step(maxv)
    # -------------------------- 关闭折线圆滑 ----------------------------
    maxv = 1.0
    for y in years:
        valid_nums = [x for x in by_year_raw[y] if x > 0]
        m = max(valid_nums) if valid_nums else 0.0
        if m > maxv:
            maxv = m
    _step = _nice_step(maxv)

    chart.y_axis.scaling.min = 0
    chart.y_axis.majorUnit = _step
    chart.y_axis.scaling.max = _ceil_to(maxv, _step)
    # ===== 全部标签关闭：数值、类别、系列名 =====
        # ===== 全部标签关闭：数值、类别、系列名 =====
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = False
    chart.dataLabels.showCatName = False
    chart.dataLabels.showSerName = False

    # 图例放在图表底部，不遮挡折线
    chart.legend.position = "b"
    chart.legend.overlay = False

    # 图表加高一点，给图例留出空间
    chart.height = 14

    anchor_row = table_last_row + 3
    ws.add_chart(chart, f"A{anchor_row}")
    return anchor_row


# =============================== 品牌分析：表格 + 饼图 + 小结 ==================================
BRAND_COLS = 8
BRAND_HEADERS = [
    "目标产品", "平台", "类目", "品牌情况",
    "集中度说明", "竞争情况", "目标产品类目占比", "切入难度",
]
BRAND_HEADER_FILL = "548235"
_PIE_COLORS = ("4472C4", "ED7D31", "70AD47", "C00000", "7030A0", "00B0F0", "FFC000", "A6A6A6")


def _write_brand_table(ws, shop_name: str, rows: List[Dict[str, Any]], start_row: int) -> int:
    """写入品牌分析表格(8列)，返回表格末行行号。"""
    # 标题行
    ws.cell(start_row, 1, f"{shop_name} 品牌分析")
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=BRAND_COLS)
    for c in range(1, BRAND_COLS + 1):
        ws.cell(start_row, c).fill = PatternFill("solid", fgColor=TITLE_FILL)
    tc = ws.cell(start_row, 1)
    tc.font = Font(size=14, bold=True, color="FFFFFF")
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[start_row].height = 28

    header_row = start_row + 1
    for j, h in enumerate(BRAND_HEADERS, start=1):
        cell = ws.cell(header_row, j, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=BRAND_HEADER_FILL)
        cell.alignment = CENTER
    ws.row_dimensions[header_row].height = 24

    first_body = header_row + 1
    row = first_body
    if not rows:
        cell = ws.cell(row, 2, "暂无数据")
        cell.alignment = CENTER
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=BRAND_COLS)
        ws.row_dimensions[row].height = 22
        row += 1
    else:
        n = len(rows)
        for i, rdata in enumerate(rows):
            values = [
                str(rdata.get("目标产品", "")),
                str(rdata.get("平台", "")),
                str(rdata.get("类目", "")),
                str(rdata.get("品牌情况", "")),
                str(rdata.get("集中度说明", "")),
                str(rdata.get("竞争情况", "")),
                str(rdata.get("目标产品类目占比", "")),
                str(rdata.get("切入难度", "")),
            ]
            for j, v in enumerate(values, start=1):
                cell = ws.cell(row, j, v)
                cell.alignment = CENTER if j in (1, 2, 7, 8) else LEFT_CENTER
            ws.row_dimensions[row].height = 34
            row += 1
        # 合并目标产品列(覆盖全部行)
        ws.merge_cells(start_row=first_body, start_column=1, end_row=row - 1, end_column=1)
        shop_cell = ws.cell(first_body, 1, shop_name)
        shop_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    if row - 1 >= header_row:
        _apply_border(ws, header_row, row - 1, 1, BRAND_COLS)

    # 列宽(品牌情况/竞争情况较宽)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 38
    ws.column_dimensions["E"].width = 26
    ws.column_dimensions["F"].width = 38
    ws.column_dimensions["G"].width = 16
    ws.column_dimensions["H"].width = 12

    return row - 1


def _write_pie_charts(ws, pie_data: List[Dict[str, Any]], shop_name: str,
                      anchor_row: int, data_start_col: int, data_start_row: int) -> int:
    """为每个平台生成品牌份额饼图(数据源写入隐藏列)，返回饼图区占用末行。"""
    if not pie_data:
        return anchor_row - 1

    # 标题行
    trow = anchor_row
    ws.cell(trow, 1, f"{shop_name} 各平台品牌份额饼图")
    ws.merge_cells(start_row=trow, start_column=1, end_row=trow, end_column=BRAND_COLS)
    for c in range(1, BRAND_COLS + 1):
        ws.cell(trow, c).fill = PatternFill("solid", fgColor=SUBTOTAL_FILL)
    tt = ws.cell(trow, 1)
    tt.font = Font(size=12, bold=True, color="1F4E79")
    tt.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[trow].height = 22

    chart_row = trow + 2
    max_rows = 0
    for k, item in enumerate(pie_data):
        platform = str(item.get("平台", "平台"))
        slices = item.get("slices", [])
        if not isinstance(slices, list) or not slices:
            continue
        dcol = data_start_col + k * 2
        n = len(slices)
        # 数据源(隐藏列)：标题 + 品牌名/占比
        ws.cell(data_start_row, dcol, "品牌")
        ws.cell(data_start_row, dcol + 1, f"{platform}-占比")
        for j, s in enumerate(slices):
            if not isinstance(s, dict):
                continue
            ws.cell(data_start_row + 1 + j, dcol, str(s.get("品牌", "")))
            pct = s.get("占比")
            ws.cell(data_start_row + 1 + j, dcol + 1, round(float(pct) / 100.0, 6) if pct is not None else 0.0)
        ws.cell(data_start_row + 1 + n, dcol, "合计")
        ws.cell(data_start_row + 1 + n, dcol + 1, 1.0)

        data = Reference(ws, min_col=dcol + 1, min_row=data_start_row + 1,
                         max_col=dcol + 1, max_row=data_start_row + 1 + n)
        cats = Reference(ws, min_col=dcol, min_row=data_start_row + 1,
                         max_row=data_start_row + 1 + n)

        pie = PieChart()
        pie.title = f"{platform}品牌份额"
        pie.style = 13
        pie.add_data(data, titles_from_data=True)
        pie.set_categories(cats)
        pie.visible_cells_only = False
        pie.dataLabels = DataLabelList()
        pie.dataLabels.showPercent = True
        pie.dataLabels.showVal = False
        # 品牌名 + 占比标签
        pie.height = 8
        pie.width = 9
        anchor_c = get_column_letter(1 + k * 6)
        ws.add_chart(pie, f"{anchor_c}{chart_row}")
        if n > max_rows:
            max_rows = n

    return chart_row + max_rows


def _write_analysis_text(ws, analysis_result: Dict[str, Any], start_row: int) -> int:
    """在品牌区域后写入市场分析报告文本(合并单元格换行)。"""
    text = analysis_result.get("analysis_text") or "无数据"
    if not isinstance(text, str):
        text = str(text)

    row = start_row
    ws.cell(row, 1, "市场分析报告")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=BRAND_COLS)
    for c in range(1, BRAND_COLS + 1):
        ws.cell(row, c).fill = PatternFill("solid", fgColor=SUBTOTAL_FILL)
    tc = ws.cell(row, 1)
    tc.font = Font(size=12, bold=True, color="1F4E79")
    ws.row_dimensions[row].height = 22
    row += 1

    # 将文本按行写入
    lines = text.splitlines()
    if not lines:
        lines = [text]
    for ln in lines:
        cell = ws.cell(row, 1, ln)
        cell.alignment = LEFT_CENTER
        cell.font = Font(size=10)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=BRAND_COLS)
        ws.row_dimensions[row].height = 24 if ln else 6
        row += 1
    return row


def _write_brand_section(ws, shop_name: str, analysis_result: Dict[str, Any],
                         pie_data: List[Dict[str, Any]], start_row: int) -> int:
    """在折线图下方依次写入：品牌分析表格 + 饼图 + 饼图说明/数据来源/小结 + 分析报告文本。"""
    rows = analysis_result.get("brand_table") or []
    if not isinstance(rows, list):
        rows = []

    cur = start_row
    table_last = _write_brand_table(ws, shop_name, rows, cur)
    cur = table_last + 2

    # 饼图数据源使用隐藏列(位于品牌表8列之后)
    data_start_col = BRAND_COLS + 3  # 从 K 列开始
    data_start_row = cur - 1
    pie_bottom = _write_pie_charts(ws, pie_data, shop_name, cur, data_start_col, data_start_row)

    # 饼图说明/数据来源/小结
    note_row = max(cur, pie_bottom) + 2
    summary = analysis_result.get("brand_summary") or "无数据"
    if not isinstance(summary, str):
        summary = str(summary)
    for ln in summary.splitlines():
        cell = ws.cell(note_row, 1, ln)
        cell.font = Font(size=9, color="808080")
        cell.alignment = LEFT_CENTER
        ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=BRAND_COLS)
        ws.row_dimensions[note_row].height = 18 if ln else 6
        note_row += 1

    # 分析报告文本
    text_last = _write_analysis_text(ws, analysis_result, note_row + 1)
    return text_last




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

    # 折线图下方：品牌分析表格 + 饼图 + 小结 + 分析报告文本
    # 折线图 anchor = table_last_row + 3，高14cm 约占26行，因此品牌区从其后约30行开始
    chart_anchor = table_last_row + 3
    brand_start = chart_anchor + 32
    _write_brand_section(
        ws,
        shop_name,
        state.analysis_result if isinstance(state.analysis_result, dict) else {},
        state.pie_data if isinstance(state.pie_data, list) else [],
        brand_start,
    )

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