from typing import List, Dict, Any, Optional
from pydantic import BaseModel, Field


class CategoryData(BaseModel):
    """单个类目的聚合信息"""
    platform: str = Field(default="", description="所属平台（如天猫/抖音/京东）")
    category_name: str = Field(default="", description="类目名称")
    agg_year: List[Dict[str, Any]] = Field(default=[], description="按年聚合数据，每项含 年/销售额(元)/销量(件)/均价(元)")
    agg_quarter: List[Dict[str, Any]] = Field(default=[], description="按季度聚合数据，每项含 季度/销售额(元)")


class QuarterData(BaseModel):
    """季度维度聚合信息（按平台/类目）"""
    platform: str = Field(default="", description="所属平台")
    category_name: str = Field(default="", description="类目名称")
    agg_quarter: List[Dict[str, Any]] = Field(default=[], description="按季度聚合数据，每项含 季度/销售额(元)")


class GlobalState(BaseModel):
    """全局状态定义"""
    raw_data: str = Field(default="", description="钉钉平台输出的原始数据JSON字符串")
    shop_name: str = Field(default="", description="目标产品名称(大类)")
    stat_time: str = Field(default="", description="统计时间")
    categories: List[CategoryData] = Field(default=[], description="平台类目聚合数据(含年度与季度)")
    quarter_data: List[QuarterData] = Field(default=[], description="平台类目按季度聚合数据")
    monthly_summary: List[Dict[str, Any]] = Field(default=[], description="全局月度汇总数据")
    years: List[str] = Field(default=[], description="出现的年份列表(升序)")
    platform_brand_summary: List[Dict[str, Any]] = Field(default=[], description="平台品牌汇总数据(platform_brand_summary)")
    analysis_result: Dict[str, Any] = Field(default={}, description="市场分析结果(年度/季度/平台品牌/趋势/结论 + 品牌分析表)")
    pie_data: List[Dict[str, Any]] = Field(default=[], description="各平台品牌份额饼图数据")
    report_key: str = Field(default="", description="对象存储上报的key")
    report_url: str = Field(default="", description="Excel报告下载URL")


class GraphInput(BaseModel):
    """工作流输入"""
    data_json: str = Field(..., description="钉钉平台输出的原始数据JSON字符串(含 output_json_string 或直接为数据对象)")


class GraphOutput(BaseModel):
    """工作流输出"""
    report_key: str = Field(..., description="对象存储上报的key")
    report_url: str = Field(..., description="Excel报告下载URL")
    shop_name: str = Field(..., description="目标产品名称")


class ParseDataInput(BaseModel):
    """数据解析节点输入"""
    data_json: str = Field(..., description="钉钉平台输出的原始数据JSON字符串")


class ParseDataOutput(BaseModel):
    """数据解析节点输出"""
    shop_name: str = Field(..., description="目标产品名称(大类)")
    stat_time: str = Field(..., description="统计时间")
    categories: List[CategoryData] = Field(default=[], description="平台类目聚合数据(含年度与季度)")
    quarter_data: List[QuarterData] = Field(default=[], description="平台类目按季度聚合数据")
    monthly_summary: List[Dict[str, Any]] = Field(default=[], description="全局月度汇总数据")
    years: List[str] = Field(default=[], description="出现的年份列表(升序)")
    platform_brand_summary: List[Dict[str, Any]] = Field(default=[], description="平台品牌汇总数据")


class AnalysisInput(BaseModel):
    """市场分析节点输入"""
    shop_name: str = Field(..., description="目标产品名称(大类)")
    stat_time: str = Field(..., description="统计时间")
    raw_data: str = Field(default="", description="原始数据JSON字符串(用于季度等维度分析)")
    categories: List[CategoryData] = Field(default=[], description="平台类目聚合数据(含年度与季度)")
    quarter_data: List[QuarterData] = Field(default=[], description="平台类目按季度聚合数据")
    monthly_summary: List[Dict[str, Any]] = Field(default=[], description="全局月度汇总数据")
    years: List[str] = Field(default=[], description="出现的年份列表(升序)")
    platform_brand_summary: List[Dict[str, Any]] = Field(default=[], description="平台品牌汇总数据")


class AnalysisOutput(BaseModel):
    """市场分析节点输出"""
    analysis_result: Dict[str, Any] = Field(default={}, description="市场分析结果(含分析文本与品牌分析表)")
    pie_data: List[Dict[str, Any]] = Field(default=[], description="各平台品牌份额饼图数据")


class GenerateReportInput(BaseModel):
    """报告生成节点输入"""
    shop_name: str = Field(..., description="目标产品名称(大类)")
    stat_time: str = Field(..., description="统计时间")
    categories: List[CategoryData] = Field(default=[], description="平台类目按年聚合数据")
    monthly_summary: List[Dict[str, Any]] = Field(default=[], description="全局月度汇总数据")
    years: List[str] = Field(default=[], description="出现的年份列表(升序)")
    platform_brand_summary: List[Dict[str, Any]] = Field(default=[], description="平台品牌汇总数据")
    analysis_result: Dict[str, Any] = Field(default={}, description="市场分析结果(含分析文本与品牌分析表)")
    pie_data: List[Dict[str, Any]] = Field(default=[], description="各平台品牌份额饼图数据")


class GenerateReportOutput(BaseModel):
    """报告生成节点输出"""
    report_key: str = Field(..., description="对象存储上报的key")
    report_url: str = Field(..., description="Excel报告下载URL")


class FileReportOutput(BaseModel):
    """文件型报告输出"""
    report_url: str = Field(..., description="Excel报告下载URL")