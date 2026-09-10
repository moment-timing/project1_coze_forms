import json
from typing import List, Dict, Any

from langchain_core.runnables import RunnableConfig
from langgraph.runtime import Runtime
from coze_coding_utils.runtime_ctx.context import Context

from graphs.state import ParseDataInput, ParseDataOutput, CategoryData


def _extract_inner_data(raw: str) -> Dict[str, Any]:
    """将输入数据解析为最内层的数据对象。

    支持两种输入：
    1. 外层包裹 {"output_json_string": "<json字符串>"}
    2. 直接为数据对象本身
    """
    if not raw or not raw.strip():
        raise ValueError("输入数据为空，请提供钉钉平台的输出数据")

    data = json.loads(raw)
    if not isinstance(data, dict):
        raise ValueError("输入数据不是合法的JSON对象")

    inner = data.get("output_json_string")
    if inner is not None:
        if isinstance(inner, str):
            data = json.loads(inner)
        elif isinstance(inner, dict):
            data = inner

    if not isinstance(data, dict):
        raise ValueError("数据解析后不是合法的JSON对象")

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
        categories.append(
            CategoryData(
                platform=str(c.get("platform", "")),
                category_name=str(c.get("category_name", "")),
                agg_year=agg_year,
            )
        )
    return categories


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

    return ParseDataOutput(
        shop_name=shop_name,
        stat_time=stat_time,
        categories=categories,
        monthly_summary=monthly,
        years=years,
        platform_brand_summary=brand_summary,
    )