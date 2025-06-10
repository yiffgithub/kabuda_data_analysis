from pygwalker.api.streamlit import StreamlitRenderer
import pandas as pd
import streamlit as st

st.set_page_config(page_title="kabuda 数据可视化平台", layout="wide")
st.title("kabuda 数据可视化平台")

EXCEL_PATH = r"E:\kabuda_data_analysis\司机业务信息库\司机业务_统计分析.xlsx"

# 获取所有sheet名
sheet_names = pd.ExcelFile(EXCEL_PATH).sheet_names
sheet_choice = st.selectbox("请选择要分析的 sheet", sheet_names)

@st.cache_resource
def get_pyg_renderer(sheet_name) -> "StreamlitRenderer":
    df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name)
    return StreamlitRenderer(df, spec=f"./gw_config_{sheet_name}.json", spec_io_mode="rw")

renderer = get_pyg_renderer(sheet_choice)
renderer.explorer()
