import os
import time
from typing import Dict, List, Any, Optional

from langchain_core.runnables import RunnableConfig
from langgraph.runtime import Runtime
from coze_coding_utils.runtime_ctx.context import Context
from coze_coding_dev_sdk.s3 import S3SyncStorage

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference

from graphs.state import GenerateReportInput, GenerateReportOutput, CategoryData

HEADER_FILL = "4472C4"
TITLE_FILL = "1F4E79"
SUBTOTAL_FILL = "DDEBF7"
TOTAL_FILL = "FFF2CC"
THIN_SIDE = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_CENTER = Alignment(horizontal="left", vertical="center", wrap_text=True)


def _gmv_for(agg_year: List[Dict[str, Any]], year: str) -> float:
    """返回某类目在指定年份的销售额(元)，不存在返回 0。"""
    for ay in agg_year:
        if not isinstance(ay, dict):
            continue
        if str(ay.get("年")) == str(year):
            v = ay.get("销售额(元)")
            return float(v) if v is not None else 0.0
    return 0.0


def _calc_growth(last: float, prev: float) -> Optional[float]:
    """计算较上年增幅(百分数)，上年为0或缺失返回 None。"""
    if prev and abs(prev) > 1e-9:
        return round((last - prev) / prev * 100.0, 2)
    return None


def _build_monthly_by_year(monthly: List[Dict[str, Any]], years: List[str]) -> Dict[str, List[float]]:
    """将全局月度汇总按年份组织为 [1..12] 每月的销售额列表。"""
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
            by_year[y][month - 1] = float(v) if v is not None else 0.0
    return by_year


def _apply_border(ws, r1: int, r2: int, c1: int, c2: int) -> None:
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = BORDER


def _build_table(ws, shop_name: str, stat_time: str, years: List[str], categories: List[CategoryData]) -> int:
    """构建“市场体量”数据表，返回最后一行的行号。"""
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
    ws.row_dimensions[1].height = 30

    # 统计信息行
    ws.cell(2, 1, f"统计时间：{stat_time}    |    数据来源：各平台品类数据")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    info_cell = ws.cell(2, 1)
    info_cell.font = Font(size=9, color="808080")
    info_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 18

    # 表头行
    headers: List[str] = ["目标产品", "平台", "类目"] + [f"{y}年预估GMV" for y in years] + [f"{years[-1]}年增幅"]
    header_row = 3
    for j, h in enumerate(headers, start=1):
        cell = ws.cell(header_row, j, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
        cell.alignment = CENTER
    ws.row_dimensions[header_row].height = 22

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
            ws.cell(row, 3, cat.category_name).alignment = LEFT_CENTER
            for i, y in enumerate(years):
                v = _gmv_for(cat.agg_year, y)
                cell = ws.cell(row, 4 + i, round(v / 10000.0, 1))
                cell.number_format = "#,##0.0"
                plat_gmv[i] += v
                total_gmv[i] += v
            last_col_idx = 3 + len(years) + 1
            g = _calc_growth(_gmv_for(cat.agg_year, years[-1]), _gmv_for(cat.agg_year, years[-2]))
            cell = ws.cell(row, last_col_idx, g)
            cell.number_format = '0.00"%"'
            if g is not None and g < 0:
                cell.font = Font(color="C00000")
            row += 1

        # 平台小计行
        ws.cell(row, 1)
        ws.cell(row, 2)
        sub_cell = ws.cell(row, 3, f"{plat}小计")
        sub_cell.font = Font(bold=True)
        sub_cell.alignment = CENTER
        for i in range(len(years)):
            cell = ws.cell(row, 4 + i, round(plat_gmv[i] / 10000.0, 1))
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
        cell = ws.cell(row, 4 + i, round(total_gmv[i] / 10000.0, 1))
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
    total_row = row
    row += 1

    # 合并目标产品列(覆盖全部数据行含总计)
    ws.merge_cells(start_row=first_body_row, start_column=1, end_row=total_row, end_column=1)
    shop_cell = ws.cell(first_body_row, 1, shop_name)
    shop_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 边框
    if total_row >= header_row:
        _apply_border(ws, header_row, total_row, 1, n_cols)

    # 列宽
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 52
    for i in range(len(years)):
        ws.column_dimensions[get_column_letter(4 + i)].width = 14
    ws.column_dimensions[last_col].width = 12

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

    return total_row


def _build_chart(wb, shop_name: str, years: List[str], monthly: List[Dict[str, Any]]) -> None:
    """构建折线图：x轴为01-12月份，每条折线代表一个年份的销售额。"""
    ws2 = wb.create_sheet("趋势数据")
    ws2.cell(1, 1, "月份")
    for i, y in enumerate(years):
        ws2.cell(1, 2 + i, str(y))
    by_year = _build_monthly_by_year(monthly, years)
    for m in range(1, 13):
        ws2.cell(1 + m, 1, f"{m:02d}")
        for i, y in enumerate(years):
            ws2.cell(1 + m, 2 + i, round(by_year[y][m - 1], 2))

    ws3 = wb.create_sheet("市场趋势")
    chart = LineChart()
    chart.title = f"{shop_name}月度销售额趋势(元)"
    chart.style = 2
    chart.y_axis.title = "销售额(元)"
    chart.x_axis.title = "月份"
    chart.height = 12
    chart.width = 30

    data = Reference(ws2, min_col=2, min_row=1, max_col=1 + len(years), max_row=13)
    cats = Reference(ws2, min_col=1, min_row=2, max_row=13)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws3.add_chart(chart, "A1")


def generate_report_node(
    state: GenerateReportInput,
    config: RunnableConfig,
    runtime: Runtime[Context],
) -> GenerateReportOutput:
    """
    title: 生成Excel分析报表
    desc: 基于解析后的数据构建“市场体量”表(含合并单元格)与“市场趋势”折线图，生成Excel并上传对象存储
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
    _build_table(ws, shop_name, state.stat_time, years, state.categories)
    _build_chart(wb, shop_name, years, state.monthly_summary)

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