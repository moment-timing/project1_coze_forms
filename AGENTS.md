## 项目概述
- **名称**: 电商销售数据分析报表工作流
- **功能**: 接收钉钉平台输出的销售数据JSON，解析数据并生成包含"市场体量"数据表(合并单元格+各年份GMV)与"市场趋势"折线图(按月、按年)的Excel分析报表，上传对象存储返回下载URL

### 节点清单
| 节点名 | 文件位置 | 类型 | 功能描述 | 分支逻辑 | 配置文件 |
|-------|---------|------|---------|---------|---------|
| parse_data | `src/graphs/nodes/parse_data_node.py` | task | 解析钉钉JSON数据，提取目标产品、类目按年聚合数据、月度汇总、年份列表与分平台品牌汇总 | - | - |
| analysis | `src/graphs/nodes/analysis_node.py` | agent | 用源数据计算年度/季度/趋势/品牌指标（CR3、HHI、集中度、切入难度）与饼图数据，并调用大模型生成叙述性市场分析；无品牌数据时输出"无数据" | - | `config/analysis_llm_cfg.json` |
| generate_report | `src/graphs/nodes/generate_report_node.py` | task | 构建Excel报表(市场体量表格+市场趋势折线图+品牌分析表与饼图+市场分析报告)并上传对象存储 | - | - |

**类型说明**: task(task节点) / agent(大模型) / condition(条件分支) / looparray(列表循环) / loopcond(条件循环)

## 子图清单
无

## 技能使用
- 节点 `analysis` 使用大语言模型(llm)技能生成叙述性市场分析报告
- 节点 `generate_report` 使用对象存储(storage)技能上传Excel并生成签名URL
- 依赖库：`openpyxl` 生成带样式、合并单元格与内嵌折线图/饼图的 xlsx

## 数据说明
- 输入 `data_json`：钉钉平台输出，支持 `{"output_json_string": "<json字符串>"}` 包裹或直接为数据对象
- 数据内字段：`meta.shop_name` / `meta.stat_time`、`category_list[].{platform, category_name, agg_year, agg_quarter}`、`global_monthly_summary[].{月, 销售额(元)}`、`platform_brand_summary[].{platform, brand_list:[{品牌, 销售额(元), 销量(件)}]}`
- agg_year 内为 `{年, 销售额(元), 销量(件), 均价(元)}`
- 品牌指标（CR3、HHI、集中度、切入难度、饼图占比）由 `analysis` 节点基于 `platform_brand_summary` 确定性计算，严格取自源数据，无数据时如实输出"无数据"，禁止编造