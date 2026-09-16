import os
import time
import math
import re
from openpyxl.chart.text import RichText
from openpyxl.chart.title import Title
from openpyxl.drawing.text import (
    Paragraph, ParagraphProperties, CharacterProperties, RegularTextRun,
)
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
from openpyxl.chart.marker import Marker, DataPoint
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.text import RichText as ChartRichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import cm_to_EMU

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
BODY_ROW_H = 30
TITLE_ROW_H = 32
HEAD_FONT = Font(size=13, bold=True, color="1F4E79")  # 分析章节标题样式(加粗+着色)

# 已知的分析章节标题标签(用于识别内联标题，如 核心机会点：正文)
_HEAD_LABELS = (
    "核心机会点", "核心风险点", "机会点", "风险点", "品牌格局", "市场分析报告",
    "综合结论与洞察", "综合结论", "总结", "结论", "摘要", "洞察", "开篇",
    "年度维度分析", "季度维度分析", "月度维度分析", "平台表现与品牌格局分析",
    "大盘趋势分析", "趋势分析", "建议", "总体概述", "现状",
)
COL_WIDTHS = {"A": 20, "B": 18, "C": 70}
YEAR_COL_W = 16
GROWTH_COL_W = 13

# ============ 可手动调节的行高/列宽系数(推荐先改这里即可调整体排版) ============
ROW_H_SCALE = 1.0      # 行高整体缩放系数(如文字仍被遮挡可调至 1.1~1.3)
COL_W_SCALE = 1.0      # 列宽整体缩放系数(如拼音太挤可调至 1.1~1.2)
LINE_H = 15.0          # 单行文字基线高度(磅)，与字号联动
MIN_ROW_H = 20.0       # 最小行高(磅)
CH_PER_UNIT = 2.0      # 一列宽单位可容纳的显示宽度(中文按2计)
# 品牌分析区列宽(可手动微调；最终会再乘 COL_W_SCALE)
BRAND_COL_WIDTHS = {"A": 20, "B": 18, "C": 38, "D": 18, "E": 18, "F": 18, "G": 18, "H": 13}
# 饼图区参数
PIE_SPACING = 2        # 相邻饼图间隔列数(过大会导致三个饼图相距过远)
PIE_WIDTH = 8        # 饼图宽度(cm)
PIE_HEIGHT = 8       # 饼图高度(cm)
PIE_H_ROWS = 16        # 饼图占用行数(由高度估算，用于文字避让)
PIE_GAP_CM = 3     # 相邻两张饼图之间的留白(厘米)，间距即由它控制
PIE_SLOT_CM = 9.5   # 每张饼图占用的横向槽宽(厘米)，含标签+间距

# 行高/列宽可整体调节系数(内容换行后按需放大，手动微调请改这里)
ROW_H_SCALE = 1.0      # 行高整体缩放(>1 放大，<1 缩小)
COL_W_SCALE = 1.0      # 列宽整体缩放
LINE_H = 15.0          # 单行文字基准高度(磅)



def _disp_width(s: Any) -> float:
    """估算文本的显示宽度(中文/全角按2单位，其余按1)。"""
    if s is None:
        return 0.0
    return float(sum(2.0 if ord(ch) > 0x2E7F else 1.0 for ch in str(s)))


def _calc_lines(s: Any, merged_chars: float) -> int:
    """估算一段文本在指定列宽(字符单位)下换行后占用的行数(含显式换行)。"""
    text = "" if s is None else str(s)
    if not text:
        return 1
    per_line = max(4.0, float(merged_chars) - 2.0)
    total = 0
    for seg in str(text).split("\n"):
        total += max(1, math.ceil(_disp_width(seg) / per_line))
    return max(1, total)


def _auto_height(ws, row: int, values: List[Any], merged_chars: float) -> None:
    """根据单元格内容自动计算行高，解决文字被遮挡问题。"""
    lines = 1
    for v in values:
        lines = max(lines, _calc_lines(v, merged_chars))
    ws.row_dimensions[row].height = max(20.0, lines * LINE_H * ROW_H_SCALE + 6.0)




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
    chart.width = int(col_chars * 6.28 / 96.0 * 2.54)
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





def _row_h(*text_col: Any) -> float:
    """根据多个 (文本,列宽) 计算该行需要的高度(磅)。"""
    need = 1
    for text, w in text_col:
        need = max(need, _calc_lines(text, w))
    h = need * LINE_H * ROW_H_SCALE + 4.0
    return max(MIN_ROW_H, h)


def _set_auto_row(ws, row: int, col_widths: Dict[str, float],
                  *cell_texts: Any) -> None:
    """按单元格内容自动设置第 row 行行高。cell_texts 顺序对应 1..N 列。"""
    pairs = []
    for idx, text in enumerate(cell_texts, start=1):
        w = float(col_widths.get(get_column_letter(idx), 12.0))
        pairs.append((text, w))
    ws.row_dimensions[row].height = _row_h(*pairs)


def _write_brand_table(ws, shop_name: str, rows: List[Dict[str, Any]], start_row: int) -> int:
    """写入品牌分析表格(8列)，返回表格末行行号。"""
    # 应用品牌列宽度(含整体缩放系数)
    for col, w in BRAND_COL_WIDTHS.items():
        ws.column_dimensions[col].width = round(w * COL_W_SCALE, 2)

    # 标题行
    ws.cell(start_row, 1, f"{shop_name} 品牌分析")
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=BRAND_COLS)
    for c in range(1, BRAND_COLS + 1):
        ws.cell(start_row, c).fill = PatternFill("solid", fgColor=TITLE_FILL)
    tc = ws.cell(start_row, 1)
    tc.font = Font(size=14, bold=True, color="FFFFFF")
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[start_row].height = _row_h((f"{shop_name} 品牌分析", sum(BRAND_COL_WIDTHS.values())))

    header_row = start_row + 1
    for j, h in enumerate(BRAND_HEADERS, start=1):
        cell = ws.cell(header_row, j, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=BRAND_HEADER_FILL)
        cell.alignment = CENTER
    ws.row_dimensions[header_row].height = max(24, _row_h(*[(h, BRAND_COL_WIDTHS[get_column_letter(i)]) for i, h in enumerate(BRAND_HEADERS, start=1)]))

    first_body = header_row + 1
    row = first_body
    if not rows:
        cell = ws.cell(row, 2, "暂无数据")
        cell.alignment = CENTER
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=BRAND_COLS)
        ws.row_dimensions[row].height = 22
        row += 1
    else:
        # 记录每一行的平台值，用于同平台合并
        plat_values: List[str] = []
        for rdata in rows:
            values = [
                str(rdata.get("目标产品", "")),
                str(rdata.get("平台", "")),
                str(rdata.get("类目", "")),
                str(rdata.get("品牌情况", "")).replace("发布图", "份额图"),
                str(rdata.get("集中度说明", "")),
                str(rdata.get("竞争情况", "")),
                str(rdata.get("目标产品类目占比", "")),
                str(rdata.get("切入难度", "")),
            ]
            plat_values.append(values[1])
            for j, v in enumerate(values, start=1):
                cell = ws.cell(row, j, v)
                cell.alignment = CENTER if j in (1, 2, 7, 8) else LEFT_CENTER
            # 行高根据各列内容(含换行)自动计算
            ws.row_dimensions[row].height = _row_h(*[
                (values[i], BRAND_COL_WIDTHS[get_column_letter(i + 1)]) for i in range(BRAND_COLS)
            ])
            row += 1
        # 合并目标产品列(覆盖全部行)
        ws.merge_cells(start_row=first_body, start_column=1, end_row=row - 1, end_column=1)
        shop_cell = ws.cell(first_body, 1, shop_name)
        shop_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # 同平台平台列合并显示(类目列隔开，保留每行)
        i = 0
        body_end = row - 1
        while i < len(plat_values):
            pv = plat_values[i]
            start_idx = i
            while i + 1 < len(plat_values) and plat_values[i + 1] == pv:
                i += 1
            end_idx = i
            if end_idx > start_idx:
                ws.merge_cells(start_row=first_body + start_idx, start_column=2,
                               end_row=first_body + end_idx, end_column=2)
                pc = ws.cell(first_body + start_idx, 2)
                pc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            i += 1

    if row - 1 >= header_row:
        _apply_border(ws, header_row, row - 1, 1, BRAND_COLS)

    return row - 1


def _col_width_cm(ws, col_idx: int) -> float:
    """某列的宽度换算成厘米(1 列宽单位≈7px, 96px=1in, 1in=2.54cm)。"""
    dim = ws.column_dimensions[get_column_letter(col_idx)]
    w = dim.width if dim.width else 8.43   # Excel 默认列宽
    return w * 7.0 / 96.0 * 2.54

def _col_at_cm(ws, cm_pos: float, max_col: int = 400) -> int:
    """给定一个累计厘米位置(从第1列左边缘算起)，返回它落在哪一列(列号)。
    用于按物理厘米精确排布饼图，保证多张图间距严格一致。"""
    acc = 0.0
    for c in range(1, max_col):
        acc += _col_width_cm(ws, c)
        if acc > cm_pos:
            return c
    return max_col


def _find_col_by_cm(ws, start_col: int, cm: float, max_col: int = 200) -> int:
    """从 start_col 开始向右累加列宽，返回「累计宽度刚超过 cm」时的下一列列号。
    用于按物理距离(厘米)推进饼图锚点，避免列宽不均导致间距不一。"""
    acc = 0.0
    col = start_col
    while col < max_col:
        acc += _col_width_cm(ws, col)
        if acc >= cm:
            return col + 1
        col += 1
    return col

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
    drawn = 0
    cm_cursor = 0.0          # 累计厘米游标：从 A 列左边缘开始
    for item in pie_data:
        platform = str(item.get("平台", "平台"))
        slices = item.get("slices", [])
        if not isinstance(slices, list):
            slices = []
        valid = [s for s in slices if isinstance(s, dict)]
        if not valid:
            continue
        # 仅使用第 6 个及之后的长尾品牌不展示，防止标签混乱
        n = len(valid)
        dcol = data_start_col + drawn * 2
        ws.cell(data_start_row, dcol, "品牌")
        ws.cell(data_start_row, dcol + 1, f"{platform}-销售额(元)")
        total_val = sum(
            float(s.get("销售额(元)") or 0)
            for s in valid
            if isinstance(s.get("销售额(元)"), (int, float))
            and not isinstance(s.get("销售额(元)"), bool)
        )
        for j, s in enumerate(valid):
            if j >= 40:
                break
            brand = str(s.get("品牌", ""))
            val = s.get("销售额(元)")
            if isinstance(val, bool) or not isinstance(val, (int, float)):
                val = 0.0
            pct = (float(val) / total_val * 100) if total_val else 0.0
            label = f"{brand} {pct:.1f}%"
            ws.cell(data_start_row + 1 + j, dcol, label)      # 类别名 = "品牌 XX%"
            ws.cell(data_start_row + 1 + j, dcol + 1, float(val))
        total_data_cols = drawn * 2
        for i in range(26):  # AD 到 AZ 共 26 列
          col_letter = get_column_letter(data_start_col + i)
          ws.column_dimensions[col_letter].hidden = True
        data = Reference(ws, min_col=dcol + 1, min_row=data_start_row + 1,
                         max_col=dcol + 1, max_row=data_start_row + n)
        cats = Reference(ws, min_col=dcol, min_row=data_start_row + 1,
                         max_row=data_start_row + n)

        pie = PieChart()
        pie.title = f"{platform}品牌份额"
        # 统一标题字体：加粗，避免三个饼图标题粗细不一致
        try:
            from openpyxl.chart.text import RichTextProperties
            pie.title.tx.rich.p[0].r[0].rPr = CharacterProperties(
                sz=1200, b=True, solidFill="000000",
                latin=None,            # 不指定拉丁字体
                ea=None,               # 不指定东亚字体
                cs=None,
            )
            # 单独给东亚字体（中文用）
            from openpyxl.drawing.text import Font as DrawFont
            pie.title.tx.rich.p[0].r[0].rPr.ea = DrawFont(typeface="微软雅黑")
            pie.title.tx.rich.p[0].r[0].rPr.latin = DrawFont(typeface="微软雅黑")
        except Exception:
            pass
        
        pie.style = 13
        pie.add_data(data, titles_from_data=False)  # 不按首行当标题，避免错位
        pie.set_categories(cats)
        pie.visible_cells_only = False
        pie.dataLabels = DataLabelList()
        pie.dataLabels.showPercent = False   # 保留百分比
        pie.dataLabels.showVal = False      # 不显示数值
        pie.dataLabels.showSerName = False  # 关掉“系列1”这种冗余标签
        pie.dataLabels.showCatName = False   # 显示品牌名，帮助识别
        pie.legend.position = "r"          # 系列设在图右侧，"b"底部，"t"头部，"l"左侧
        # pie.legend.overlay = False       #设置系列文字不覆盖饼状图
        # 标签置于扇区外侧并带连接线，避免“字体太挤”/相互重叠
        try:
            pie.dataLabels.dLblPos = "outEnd"
        except Exception:
            pass
        # 数据标签字号加大，避免“字体太挤”；仅让品牌占比清晰可读
        try:
            _lbl_p = Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=1150)))
            pie.dataLabels.txPr = ChartRichText(p=[_lbl_p])
        except Exception:
            pass
        # 显式设置系列名(平台名)，彻底消除默认“系列1”标签与图例错乱
        try:
            pie.series[0].tx = SeriesLabel(v=f"{platform}品牌份额")
        except Exception:
            pass
        pie.height = max(1, int(round(PIE_HEIGHT)))
        pie.width = max(1, int(round(PIE_WIDTH)))
        # pie.legend = None                   # 关闭图例，给饼图留空间
        # 每个扇区设置易区分的颜色(避免各平台颜色相近/相同)
        _pts = []
        for _j in range(n):
            _dp = DataPoint(idx=_j)
            _dp.spPr.solidFill = _PIE_COLORS[_j % len(_PIE_COLORS)]
            _pts.append(_dp)
        pie.series[0].data_points = _pts
        slot_cm = drawn * PIE_SLOT_CM          # 该饼图槽的起点（从第1列左边缘算起）
        anchor_col = _col_at_cm(ws, slot_cm)   # 落点列
        # 计算落点列左边缘相对第1列左边缘的累计厘米，得出 colOff 补偿
        prev_cm = 0.0
        for c in range(1, anchor_col):
            prev_cm += _col_width_cm(ws, c)
        col_off_cm = max(0.0, slot_cm - prev_cm)
        marker = AnchorMarker(
            col=anchor_col - 1,               # 0-based
            colOff=cm_to_EMU(col_off_cm),
            row=chart_row - 1,                # 0-based
            rowOff=0,
        )
        size = XDRPositiveSize2D(
            cx=cm_to_EMU(PIE_WIDTH),
            cy=cm_to_EMU(PIE_HEIGHT),
        )
        pie.anchor = OneCellAnchor(_from=marker, ext=size)
        ws.add_chart(pie)                     # 注意：不传坐标参数
        drawn += 1
        if n > max_rows:
            max_rows = n

    # 饼图底部预留 PIE_H_ROWS 行，避免后续“说明/小结”文字被图表遮挡
    return chart_row + PIE_H_ROWS


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

    # 将文本按行写入(标题加粗调大、去除 markdown 符号；正文正常展示，不整体处理)
    merged_w = sum(BRAND_COL_WIDTHS.values()) * COL_W_SCALE
    lines = text.splitlines()
    if not lines:
        lines = [text]
    for ln in lines:
        clean = _strip_md(ln)
        cell = ws.cell(row, 1)
        cell.alignment = LEFT_CENTER
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=BRAND_COLS)
        need = _calc_lines(clean, merged_w)
        body_h = need * LINE_H * ROW_H_SCALE
        if not clean:
            cell.value = ""
            cell.font = Font(size=4)
            ws.row_dimensions[row].height = 6
            row += 1
            continue
        label, tail = _heading_split(clean)
        if label and tail is None:
            # 独立章节标题：整行加粗 + 着色
            cell.value = label
            cell.font = HEAD_FONT
            ws.row_dimensions[row].height = max(24, body_h + 6)
        elif label and isinstance(tail, str) and tail.strip():
            # 内联标题：标签加粗着色，正文保持默认样式
            cell.value = CellRichText([
                TextBlock(InlineFont(sz=13, b=True, color="1F4E79"), label),
                TextBlock(InlineFont(sz=10), tail),
            ])
            ws.row_dimensions[row].height = max(18, body_h + 4)
        else:
            # 正文：正常展示，不整体处理；长内容按需换行，行高充足避免被遮挡
            cell.value = clean
            cell.font = Font(size=10)
            ws.row_dimensions[row].height = max(18, body_h + 4)
        row += 1
    return row


def _strip_md(s: str) -> str:
    """去除常见的 markdown 符号(# * _ ` 及首尾空白)，避免在 Excel 中显示乱码。"""
    if s is None:
        return ""
    return re.sub(r"[#*_`>~]", "", s).strip()


def _heading_split(ln: str):
    """拆分内联标题：若行内有‘标签：正文’，返回 (label, tail)；
    若为独立标题或非标题，返回 (None, None) 以便上层区分。"""
    s = (ln or "").strip()
    if not s:
        return None, None
    # 独立章节标题：整行即标题
    if re.match(r"^[一二三四五六七八九十]+[、.．][^：:]{0,24}$", s):
        return s, None
    # 已知标签开头
    for lab in _HEAD_LABELS:
        if s.startswith(lab):
            tail = s[len(lab):]
            if tail.strip() in ("", "：", ":", "、", "。"):
                return lab, None
            return lab, tail
    return None, None


def _write_brand_section(ws, shop_name: str, analysis_result: Dict[str, Any],
                         pie_data: List[Dict[str, Any]], start_row: int) -> int:
    """在折线图下方依次写入：品牌分析表格 + 饼图 + 饼图说明/数据来源/小结 + 分析报告文本。"""
    rows = analysis_result.get("brand_table") or []
    if not isinstance(rows, list):
        rows = []

    cur = start_row
    table_last = _write_brand_table(ws, shop_name, rows, cur)
    cur = table_last + 2

    # 饼图数据源使用隐藏列(放得很靠右，避免与饼图/表格/说明文字互相遮挡)
    data_start_col = 30  # AD 列起，远离饼图与表格区域
    data_start_row = cur - 1
    pie_bottom = _write_pie_charts(ws, pie_data, shop_name, cur, data_start_col, data_start_row)

    # 饼图说明/数据来源/小结(自动避让饼图占区)
    merged_w = sum(BRAND_COL_WIDTHS.values()) * COL_W_SCALE
    note_row = max(cur, pie_bottom) + 2
    summary = analysis_result.get("brand_summary") or "无数据"
    if not isinstance(summary, str):
        summary = str(summary)
    for ln in summary.splitlines():
        clean = _strip_md(ln)
        cell = ws.cell(note_row, 1, clean)
        cell.font = Font(size=9, color="595959")
        cell.alignment = LEFT_CENTER
        ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=BRAND_COLS)
        need = _calc_lines(clean, merged_w)
        ws.row_dimensions[note_row].height = 6 if not clean else max(18, need * LINE_H * 0.9 * ROW_H_SCALE)
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
    brand_start = chart_anchor + 30
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



