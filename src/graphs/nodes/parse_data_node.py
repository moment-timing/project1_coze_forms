import ast
import json
from typing import List, Dict, Any, Optional

from langchain_core.runnables import RunnableConfig
from langgraph.runtime import Runtime
from coze_coding_utils.runtime_ctx.context import Context

from graphs.state import ParseDataInput, ParseDataOutput, CategoryData, QuarterData


def _parse_json_like(text: str) -> Any:
    """尽量宽容地将文本解析为 Python/JSON 对象。

    钉钉平台输出的内容可能使用单引号(非标准 JSON)，这里依次尝试：
    1. json.loads（标准 JSON）
    2. ast.literal_eval（兼容 Python 字面量，如单引号键/字符串）
    3. 将单引号键/值转换为双引号后 json.loads（兜底）
    """
    s = text.strip()
    if not s:
        raise ValueError("输入数据为空")
    try:
        return json.loads(s)
    except Exception:
        pass
    # 尝试 Python 字面量（单引号等）
    try:
        return ast.literal_eval(s)
    except Exception as exc:
        raise ValueError(f"无法解析输入数据: {exc}") from exc


def _extract_inner_data(raw: str) -> Dict[str, Any]:
    """将输入数据解析为最内层的数据对象。

    支持两种输入：
    1. 外层包裹 {"output_json_string": "<json字符串>"}
    2. 直接为数据对象本身
    对非标准 JSON（如单引号）做容错解析。
    """
    parsed = _parse_json_like(raw)
    data = parsed if isinstance(parsed, dict) else {}

    inner = data.get("output_json_string")
    if inner is not None:
        if isinstance(inner, str):
            data = _parse_json_like(inner)
        elif isinstance(inner, dict):
            data = inner

    if not isinstance(data, dict):
        raise ValueError("数据解析后不是合法的JSON对象")

    if not data:
        raise ValueError("数据内容为空，请提供有效数据")

    return data


def _parse_categories(cats_raw: Any) -> List[CategoryData]:
    categories: List[CategoryData] = []
    if not isinstance(cats_raw, list):
        return categories
    for c in cats_raw:
        if not isinstance(c, dict):
            continue
        agg_year: List[Dict[str, Any]] = []
        raw_agg = c.get("agg_year")
        if isinstance(raw_agg, list):
            for ay in raw_agg:
                if isinstance(ay, dict):
                    agg_year.append(dict(ay))
        agg_quarter: List[Dict[str, Any]] = []
        raw_q = c.get("agg_quarter")
        if isinstance(raw_q, list):
            for aq in raw_q:
                if isinstance(aq, dict):
                    agg_quarter.append(dict(aq))
        raw_detail: List[Dict[str, Any]] = []
        raw_d = c.get("latest_month_raw_data")
        if isinstance(raw_d, list):
            for r in raw_d:
                if isinstance(r, dict):
                    raw_detail.append(dict(r))
        sales_yoy_pct: Optional[float] = None
        raw_yoy = c.get("sales_yoy_pct")
        if isinstance(raw_yoy, (int, float)) and not isinstance(raw_yoy, bool):
            sales_yoy_pct = float(raw_yoy)
        categories.append(
            CategoryData(
                platform=str(c.get("platform", "")),
                category_name=str(c.get("category_name", "")),
                agg_year=agg_year,
                agg_quarter=agg_quarter,
                latest_month_raw_data=raw_detail,
                sales_yoy_pct=sales_yoy_pct,
            )
        )
    return categories


def _collect_quarter_data(categories: List[CategoryData]) -> List[QuarterData]:
    """从类目中提取按季度聚合数据(平台/类目维度)。"""
    result: List[QuarterData] = []
    for c in categories:
        if not c.agg_quarter:
            continue
        result.append(
            QuarterData(
                platform=c.platform,
                category_name=c.category_name,
                agg_quarter=[dict(q) for q in c.agg_quarter],
            )
        )
    return result


def _collect_raw_product_data(categories: List[CategoryData]) -> List[Dict[str, Any]]:
    """按平台合并各品类的最新一月商品原始明细数据(如实收集，不做任何改动)。"""
    result: List[Dict[str, Any]] = []
    for c in categories:
        for item in c.latest_month_raw_data:
            result.append(dict(item))
    return result


def _collect_years(categories: List[CategoryData], monthly: List[Dict[str, Any]]) -> List[str]:
    years: set = set()
    for c in categories:
        for ay in c.agg_year:
            y = ay.get("年")
            if y is not None and str(y):
                years.add(str(y))
    for m in monthly:
        ym = str(m.get("月", ""))
        if len(ym) >= 4:
            years.add(ym[:4])
    if not years:
        raise ValueError("未从数据中识别到任何年份")
    return sorted(years)


def parse_data_node(
    state: ParseDataInput,
    config: RunnableConfig,
    runtime: Runtime[Context],
) -> ParseDataOutput:
    """
    title: 数据解析
    desc: 解析钉钉平台输出的JSON数据，提取目标产品、类目按年聚合数据与全局月度汇总数据
    """
    ctx = runtime.context
    data = _extract_inner_data(state.data_json)

    meta = data.get("meta")
    if not isinstance(meta, dict):
        meta = {}
    shop_name = str(meta.get("shop_name", ""))
    stat_time = str(meta.get("stat_time", ""))

    categories = _parse_categories(data.get("category_list"))
    quarter_data = _collect_quarter_data(categories)

    monthly: List[Dict[str, Any]] = []
    raw_monthly = data.get("global_monthly_summary")
    if isinstance(raw_monthly, list):
        for m in raw_monthly:
            if isinstance(m, dict):
                monthly.append(dict(m))

    # 平台品牌汇总数据(platform_brand_summary)
    brand_summary: List[Dict[str, Any]] = []
    raw_brand = data.get("platform_brand_summary")
    if isinstance(raw_brand, list):
        for b in raw_brand:
            if isinstance(b, dict):
                brand_summary.append(dict(b))

    years = _collect_years(categories, monthly)

    raw_product_data = _collect_raw_product_data(categories)

    global_yoy: Optional[float] = None
    raw_global_yoy = data.get("global_sales_yoy_pct")
    if isinstance(raw_global_yoy, (int, float)) and not isinstance(raw_global_yoy, bool):
        global_yoy = float(raw_global_yoy)

    # 分平台销售额同比增幅：platform -> sales_yoy_pct
    platform_yoy: Dict[str, Optional[float]] = {}
    raw_plat_yoy = data.get("platform_sales_yoy_pct")
    if isinstance(raw_plat_yoy, list):
        for p in raw_plat_yoy:
            if not isinstance(p, dict):
                continue
            pname = str(p.get("platform", "") or "")
            if not pname:
                continue
            pval = p.get("sales_yoy_pct")
            platform_yoy[pname] = (
                float(pval)
                if isinstance(pval, (int, float)) and not isinstance(pval, bool)
                else None
            )

    return ParseDataOutput(
        shop_name=shop_name,
        stat_time=stat_time,
        categories=categories,
        quarter_data=quarter_data,
        monthly_summary=monthly,
        years=years,
        platform_brand_summary=brand_summary,
        raw_product_data=raw_product_data,
        global_sales_yoy_pct=global_yoy,
        platform_sales_yoy_pct=platform_yoy,
    )