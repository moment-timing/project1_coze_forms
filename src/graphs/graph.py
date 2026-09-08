from langgraph.graph import StateGraph, END

from graphs.state import GlobalState, GraphInput, GraphOutput
from graphs.nodes.parse_data_node import parse_data_node
from graphs.nodes.generate_report_node import generate_report_node

# 主图编排：解析数据 -> 生成Excel报告
builder = StateGraph(
    GlobalState,
    input_schema=GraphInput,
    output_schema=GraphOutput,
)

builder.add_node("parse_data", parse_data_node)
builder.add_node("generate_report", generate_report_node)

builder.set_entry_point("parse_data")
builder.add_edge("parse_data", "generate_report")
builder.add_edge("generate_report", END)

main_graph = builder.compile()