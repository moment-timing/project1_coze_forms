## 项目概述
- **名称**: 电商销售数据分析报表工作流
- **功能**: 接收钉钉平台输出的销售数据JSON，解析数据并生成包含"市场体量"数据表(合并单元格+各年份GMV)与"市场趋势"折线图(按月、按年)、品牌分析区、市场分析报告的Excel分析报表；若存在商品原始明细(latest_month_raw_data)则另开"商品源数据"sheet如实填充，上传对象存储返回下载URL

### 节点清单
| 节点名 | 文件位置 | 类型 | 功能描述 | 分支逻辑 | 配置文件 |
|-------|---------|------|---------|---------|---------|
| parse_data | `src/graphs/nodes/parse_data_node.py` | task | 解析钉钉JSON数据，提取目标产品、类目按年聚合、按季度聚合(agg_quarter→QuarterData)、月度汇总、年份列表、分平台品牌汇总、最近一月商品原始明细(按平台合并 raw_product_data)、类目同比 sales_yoy_pct 与全局同比 global_sales_yoy_pct | - | - |
| analysis | `src/graphs/nodes/analysis_node.py` | agent | 用源数据计算年度/季度/趋势/品牌指标（CR3、HHI、集中度、切入难度）与饼图数据，组装含季度维度(quarter_data)的分析上下文并调用大模型生成叙述性市场分析；无数据如实输出"无数据"/"无季度(deterministic fallback)"，禁止编造 | - | `config/analysis_llm_cfg.json` |
| generate_report | `src/graphs/nodes/generate_report_node.py` | task | 构建Excel报表(市场体量表+折线图+品牌分析表与饼图+市场分析报告)，增幅列读取 sales_yoy_pct/global_sales_yoy_pct(不自行计算)，折线图下方"数据范围"注释按图表高度动态定位不遮挡，存在 raw_product_data 时另开"商品源数据"sheet如实填充(含动态"四级品类"列)，上传对象存储 | - | - |

**类型说明**: task(task节点) / agent(大模型) / condition(条件分支) / looparray(列表循环) / loopcond(条件循环)

## 子图清单
无

## 技能使用
- 节点 `analysis` 使用大语言模型(llm)技能生成叙述性市场分析报告
- 节点 `generate_report` 使用对象存储(storage)技能上传Excel并生成签名URL
- 依赖库：`openpyxl` 生成带样式、合并单元格与内嵌折线图/饼图的 xlsx
- 饼图数据标签：设置系列标题以避免"系列1"，`showCatName`+`showPercent`+`dLblPos=outEnd`，数据标签字号经 `txPr defRPr` 放大以提高可读性、避免字体拥挤重叠

## 数据说明
- 输入 `data_json`：钉钉平台输出，支持 `{"output_json_string": "<json字符串>"}` 包裹或直接为数据对象
- 数据内字段：`meta.shop_name` / `meta.stat_time`、`category_list[].{platform, category_name, agg_year, agg_quarter}`、`global_monthly_summary[].{月, 销售额(元)}`、`platform_brand_summary[].{platform, brand_list:[{品牌, 销售额(元), 销量(件)}]}`
- agg_year 内为 `{年, 销售额(元), 销量(件), 均价(元)}`
- agg_quarter 内为 `{季度, 销售额(元)}`，由 parse_data 解析为 `QuarterData(platform, category_name, agg_quarter)` 经 GlobalState.quarter_data 流入 analysis 上下文，保证季度维度有据可依
- `category_list[].latest_month_raw_data` 为最近一月商品原始明细(含年/月/平台/品类/品牌/店铺/商品名/商品ID/URL/销售额/销量/均价，可能含"四级品类")。parse_data 按平台合并为 `raw_product_data`，仅经 GlobalState→GenerateReportInput 流入 generate_report，**不进入 AnalysisInput/大模型**；generate_report 在其中存在数据时另开"商品源数据"sheet 如实填充，若任一明细含"四级品类"则动态在"三级品类"后插入该列，否则沿用基础列；商品ID按文本写避免科学计数法，无数据则不新增该sheet
- 增幅字段(同比)不再内部计算，直接读取源数据 `category_list[].sales_yoy_pct` 与顶层 `global_sales_yoy_pct` 展示规则：数字(如134.13)→"134.13%"并加粗；缺失/0/null→"最新年份-1年同时期为0"；负值标红。`global_sales_yoy_pct` 用于表格"总计"行(bold)。parse_data 解析二者(布尔排除)后经 ParseDataOutput→GlobalState→GenerateReportInput 流入；**不进入 AnalysisInput/大模型**，避免改变 LLM 分析输入
- 品牌指标（CR3、HHI、集中度、切入难度、饼图占比）由 `analysis` 节点基于 `platform_brand_summary` 确定性计算，严格取自源数据，无数据时如实输出"无数据"，禁止编造