import json
import os
from typing import Dict, List, Any, Optional

from jinja2 import Template
from langchain_core.runnables import RunnableConfig
from langgraph.runtime import Runtime
from coze_coding_utils.runtime_ctx.context import Context
from coze_coding_dev_sdk import LLMClient
from langchain_core.messages import SystemMessage, HumanMessage

from graphs.state import AnalysisInput, AnalysisOutput, CategoryData, QuarterData

_EMPTY = "暂无数据"


def _fmt(v: Optional[float]) -> str:
    """金额格式为万元保留2位小数。"""
    if v is None:
        return _EMPTY
    return f"{v:,.2f}"


# ------------------------- 聚合计算（确定性，严格基于源数据） -----------------------------------
def _build_year_ctx(categories: List[CategoryData]) -> Dict[str, List[Dict[str, Any]]]:
    """按类目聚合年度销售额(万元)与同比。返回 {platform: [dict]}。"""
    by_plat: Dict[str, List[Dict[str, Any]]] = {}
    for c in categories:
        items = []
        years = {}
        for ay in c.agg_year:
            if not isinstance(ay, dict):
                continue
            y = str(ay.get("年", ""))
            v = ay.get("销售额(元)")
            years[y] = float(v) / 10000.0 if v is not None else 0.0
        ylist = sorted(years.keys())
        prev = None
        for y in ylist:
            item = {"年": y, "销售额(万元)": years[y]}
            if prev is not None and prev > 1e-9:
                item["同比%"] = round((years[y] - prev) / prev * 100.0, 2)
            else:
                item["同比%"] = None
            items.append(item)
            prev = years[y]
        by_plat.setdefault(c.platform, []).append(
            {"类目": c.category_name, "年度": items}
        )
    return by_plat


def _build_quarter_ctx(quarter_data: List[QuarterData], raw_data: str) -> str:
    """基于解析出的季度聚合数据(quarter_data)汇总为文本上下文。

    优先使用解析节点已结构化输出的 quarter_data(严格来自源数据 agg_quarter)；
    缺失时回退从原始 raw_data 中二次解析，避免"季度数据被解析丢失"。
    """
    by_ym: Dict[str, float] = {}
    lines: List[str] = []

    for qd in quarter_data:
        if not isinstance(qd, QuarterData):
            continue
        plat = qd.platform
        cat = qd.category_name
        for q in qd.agg_quarter:
            if not isinstance(q, dict):
                continue
            qname = str(q.get("季度", ""))
            v = q.get("销售额(元)")
            key = f"{plat}|{cat}|{qname}"
            by_ym[key] = float(v) / 10000.0 if v is not None else 0.0

    # 回退：从 raw_data 中二次解析 agg_quarter
    if not by_ym:
        by_ym.update(_extract_quarter_from_raw(raw_data))

    if not by_ym:
        return ""
    year_agg: Dict[str, Dict[str, float]] = {}
    for key, val in by_ym.items():
        qname = key.split("|")[-1]
        y = qname[:4]
        if y not in year_agg:
            year_agg[y] = {}
        entry: Dict[str, float] = year_agg[y]
        if qname not in entry:
            entry[qname] = 0.0
        entry[qname] = entry[qname] + val
    for y in sorted(year_agg):
        qs = year_agg[y]
        parts = [f"{q}={_fmt(qs.get(q))}" for q in sorted(qs)]
        lines.append(f"{y}年: {'、'.join(parts)}")
    return "年度季度趋势(全类目各平台合计，单位万元)：\n" + "\n".join(lines)


def _extract_quarter_from_raw(raw_data: str) -> Dict[str, float]:
    """从原始数据 raw_data 中二次解析 agg_quarter，返回 {平台|类目|季度: 万元}。"""
    result: Dict[str, float] = {}
    try:
        data = json.loads(raw_data)
        if isinstance(data, dict):
            inner = data.get("output_json_string")
            if inner is not None:
                data = inner if isinstance(inner, dict) else json.loads(inner)
        if not isinstance(data, dict):
            return result
    except Exception:
        return result
    cats_raw = data.get("category_list")
    if not isinstance(cats_raw, list):
        return result
    for c in cats_raw:
        if not isinstance(c, dict):
            continue
        plat = str(c.get("platform", ""))
        cat = str(c.get("category_name", ""))
        qs = c.get("agg_quarter")
        if not isinstance(qs, list):
            continue
        for q in qs:
            if not isinstance(q, dict):
                continue
            qname = str(q.get("季度", ""))
            v = q.get("销售额(元)")
            key = f"{plat}|{cat}|{qname}"
            result[key] = float(v) / 10000.0 if v is not None else 0.0
    return result


def _build_trend_ctx(monthly: List[Dict[str, Any]], years: List[str]) -> str:
    """按年+月度汇总为趋势上下文(万元)。"""
    by_year: Dict[str, List[Optional[float]]] = {y: [0.0] * 12 for y in years}
    for m in monthly:
        if not isinstance(m, dict):
            continue
        ym = str(m.get("月", ""))
        if len(ym) < 6:
            continue
        y = ym[:4]
        try:
            mon = int(ym[4:6])
        except ValueError:
            continue
        if y in by_year and 1 <= mon <= 12:
            v = m.get("销售额(元)")
            by_year[y][mon - 1] = (float(v) / 10000.0) if v is not None else 0.0
    lines: List[str] = []
    for y in years:
        vals = [f"{m+1:02d}月={_fmt(v)}" for m, v in enumerate(by_year[y])]
        lines.append(f"{y}年: " + "、".join(vals))
    return "月度大盘趋势(单位万元)：\n" + "\n".join(lines)


def _is_other(name: str) -> bool:
    """判断品牌名是否属于『其他/Other』聚合分类(不参与竞争情况TOP分析)。"""
    n = (name or "").lower()
    return "其他" in n or ("other" in n and "cother" not in n) or n.strip() in ("", "未知品牌")


def _conc_level(cr3: float, hhi: float) -> str:
    if cr3 >= 70 or hhi >= 2500:
        return "高集中度"
    if cr3 >= 40 or hhi >= 1000:
        return "中集中度"
    return "低集中度"


def _farm_level(top1: float, hhi: float) -> str:
    if top1 >= 50 or hhi >= 2500:
        return "高"
    if top1 >= 30 or hhi >= 1500:
        return "中"
    return "低"


def _build_brand_analysis(
    platform_brand_summary: List[Dict[str, Any]],
    shop_name: str,
    categories: List[CategoryData],
) -> Dict[str, Any]:
    """确定性计算品牌分析表格 + 饼图数据 + 小结。

    严格基于 platform_brand_summary 的真实数字；无品牌数据返回“暂无数据”。
    """
    if not platform_brand_summary:
        return {
            "has_data": False,
            "brand_table": [],
            "pie_data": [],
            "brand_summary": "由于数据未包含分平台品牌聚合信息（by_platform_brand=false），品牌分析暂无数据。",
        }

    table_rows: List[Dict[str, str]] = []
    pie_data: List[Dict[str, Any]] = []
    pie_notes: List[str] = []

    for p in platform_brand_summary:
        if not isinstance(p, dict):
            continue
        plat = str(p.get("platform", ""))
        brands = p.get("brand_list")
        if not isinstance(brands, list) or not brands:
            continue
        total = 0.0
        for b in brands:
            if isinstance(b, dict):
                v = b.get("销售额(元)")
                total += float(v) if v is not None else 0.0
        if total <= 1e-9:
            continue

        # 计算每个品牌占比(降序排列)
        shares: List[Dict[str, Any]] = []
        for b in brands:
            if not isinstance(b, dict):
                continue
            name = str(b.get("品牌", "")) or "未知品牌"
            v = b.get("销售额(元)")
            sale = float(v) if v is not None else 0.0
            if sale <= 0:
                continue
            shares.append({"品牌": name, "销售额(元)": sale, "占比": sale / total * 100.0})
        shares.sort(key=lambda x: x["销售额(元)"], reverse=True)
        if not shares:
            continue

        top3 = shares[:3]
        cr3 = sum(s["占比"] for s in top3)
        hhi = sum(s["占比"] ** 2 for s in shares)

        # 品牌情况：按指令写“如下图（XX品牌发布图）”，指向下方对应平台的饼图
        brand_desc = f"如下图（{plat}品牌发布图）"

        # 竞争情况：取除“其他Other”外销售额前三，逐行展示 top1/top2/top3
        comp_top = [s for s in shares if not _is_other(s["品牌"])][:3]
        if comp_top:
            comp = "\n".join(
                f"TOP{i+1} {s['品牌']} {s['占比']:.2f}%" for i, s in enumerate(comp_top)
            )
        else:
            comp = "无有效上榜品牌"

        # 集中度说明：CR3 / HHI / 市场集中度结论 逐行展示
        concentration = "\n".join([
            f"CR3={cr3:.2f}%",
            f"HHI={hhi:.0f}",
            f"市场{_conc_level(cr3, hhi)}",
        ])

        # 类目取该平台对应类目名；若存在多个类目则分开成多行显示(不在同一格合并)
        cat_names = [c.category_name for c in categories if c.platform == plat]
        if not cat_names:
            cat_names = ["全品类"]

        # 饼图数据：每个平台按品牌销售额做饼图，brand_list 里有几个品牌就画几个(含其他Other)，不做长尾合并
        slices_sorted = sorted(shares, key=lambda x: x["销售额(元)"], reverse=True)
        pie_slices = [{"品牌": s["品牌"], "销售额(元)": s["销售额(元)"], "占比": s["占比"]} for s in slices_sorted]
        pie_data.append({"平台": plat, "slices": pie_slices})
        pie_notes.append(
            f"{plat}：头部品牌{shares[0]['品牌']}占比{shares[0]['占比']:.2f}%，CR3={cr3:.2f}%，市场{_conc_level(cr3, hhi)}。"
        )

        # 品牌分析表：每个类目单独一行显示(类目列不合并多值)
        for cat_name in cat_names:
            table_rows.append({
                "目标产品": shop_name,
                "平台": plat,
                "类目": cat_name,
                "品牌情况": brand_desc,
                "集中度说明": concentration,
                "竞争情况": comp,
                "目标产品类目占比": "100.00%",
                "切入难度": _farm_level(shares[0]["占比"], hhi),
            })

    if not table_rows:
        return {
            "has_data": False,
            "brand_table": [],
            "pie_data": [],
            "brand_summary": "数据中未解析到有效的平台品牌份额，品牌分析暂无数据。",
        }

    summary = (
        "饼图说明：每个平台对应一张品牌份额饼图，brand_list 中的品牌按销售额全部画出，头部品牌用独立色块表示。\n"
        "数据来源：各平台品类分品牌聚合数据（platform_brand_summary）。\n"
        "小结：" + "；".join(pie_notes) +
        " 整体品牌集中度较高，切入时建议聚焦细分差异与价格带运营。"
    )
    return {
        "has_data": True,
        "brand_table": table_rows,
        "pie_data": pie_data,
        "brand_summary": summary,
    }


def _build_context(state: AnalysisInput, categories: List[CategoryData]) -> str:
    year_ctx = _build_year_ctx(categories)
    q_ctx = _build_quarter_ctx(
        list(state.quarter_data) if isinstance(state.quarter_data, list) else [],
        state.raw_data,
    )
    t_ctx = _build_trend_ctx(state.monthly_summary, state.years)

    parts: List[str] = []
    parts.append(f"目标产品(shop_name)：{state.shop_name}")
    parts.append(f"统计时间：{state.stat_time}")

    yr_lines: List[str] = []
    for plat, cats in year_ctx.items():
        for row in cats:
            items = row["年度"]
            cells = []
            for it in items:
                yoy = f"，同比{it['同比%']}%" if it.get("同比%") is not None else ""
                cells.append(f"{it['年']}年={_fmt(it['销售额(万元)'])}万元{yoy}")
            yr_lines.append(f"{plat} - {row['类目']}：" + "；".join(cells))
    if yr_lines:
        parts.append("年度维度(各类目各年销售额，单位万元)：\n" + "\n".join(yr_lines))
    else:
        parts.append("年度维度：无数据")

    parts.append(q_ctx if q_ctx else "季度维度：无数据")
    parts.append(t_ctx if t_ctx else "趋势维度：无数据")

    # 品牌维度
    brand_lines: List[str] = []
    if isinstance(state.platform_brand_summary, list):
        for p in state.platform_brand_summary:
            if not isinstance(p, dict):
                continue
            plat = str(p.get("platform", ""))
            bl = p.get("brand_list")
            if not isinstance(bl, list):
                continue
            items = []
            for b in bl:
                if not isinstance(b, dict):
                    continue
                items.append(f"{b.get('品牌','')}={b.get('销售额(元)','')}元")
            brand_lines.append(f"{plat}：品牌TOP 「" + "；".join(items) + "」")
    if brand_lines:
        parts.append("平台品牌维度(分平台品牌TOP聚合)：\n" + "\n".join(brand_lines))
    else:
        parts.append("平台品牌维度：无数据(by_platform_brand=false)")

    return "\n\n".join(parts)


def _call_llm(ctx, config: RunnableConfig, context_text: str) -> str:
    cfg_path = os.path.join(
        os.getenv("COZE_WORKSPACE_PATH", ""), config["metadata"]["llm_cfg"]
    )
    with open(cfg_path, "r", encoding="utf-8") as fd:
        cfg = json.load(fd)
    llm_args = cfg.get("config", {})
    model = llm_args.get("model", "doubao-seed-2-0-pro-260215")
    sp = cfg.get("sp", "")
    up_tpl = Template(cfg.get("up", ""))
    up = up_tpl.render({"analysis_context": context_text})

    client = LLMClient(ctx=ctx)
    messages = [
        SystemMessage(content=sp),
        HumanMessage(content=up),
    ]
    resp = client.invoke(
        messages=messages,
        model=model,
        thinking=llm_args.get("thinking", "disabled"),
        temperature=float(llm_args.get("temperature", 0.0)),
        top_p=float(llm_args.get("top_p", 0.0)),
        max_completion_tokens=int(llm_args.get("max_completion_tokens", 3000)),
    )
    content = resp.content
    if isinstance(content, str):
        return content.strip()
    if isinstance(content, list):
        texts = []
        for item in content:
            if isinstance(item, str):
                texts.append(item)
            elif isinstance(item, dict):
                t = item.get("text")
                if t:
                    texts.append(t)
        return " ".join(texts).strip()
    return str(content)


def analysis_node(
    state: AnalysisInput,
    config: RunnableConfig,
    runtime: Runtime[Context],
) -> AnalysisOutput:
    """
    title: 市场数据分析
    desc: 基于源数据确定性计算品牌指标与饼图数据，并结合年度/季度/趋势聚合调用大模型撰写市场分析报告文本；无品牌数据如实输出“无数据”
    integrations: 大语言模型
    """
    ctx = runtime.context

    # 1) 确定性品牌分析（不依赖大模型，严格基于源数据）
    brand = _build_brand_analysis(
        list(state.platform_brand_summary) if isinstance(state.platform_brand_summary, list) else [],
        state.shop_name,
        list(state.categories) if isinstance(state.categories, list) else [],
    )

    # 2) 构建上下文并调用大模型撰写叙述性分析
    ctx_text = _build_context(state, list(state.categories))
    analysis_text = _EMPTY
    try:
        analysis_text = _call_llm(ctx, config, ctx_text)
    except Exception:
        analysis_text = "（大模型分析暂不可用，以上为基于源数据的结构性汇总。）"

    analysis_result: Dict[str, Any] = {
        "analysis_text": analysis_text,
        "brand_table": brand["brand_table"],
        "brand_has_data": brand["has_data"],
        "brand_summary": brand["brand_summary"],
    }
    return AnalysisOutput(
        analysis_result=analysis_result,
        pie_data=brand["pie_data"],
    )