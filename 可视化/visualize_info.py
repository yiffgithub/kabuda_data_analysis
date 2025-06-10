#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
driver_data_viz.py
------------------
使用 pandas 和 pyecharts 对 “司机信息.xlsx” 中各个 sheet 数据进行预处理，
并按维度生成以下图表：
 1. 司机等级分布 (Bar)
 2. 区域覆盖 TreeMap
 3. 职业 vs 活跃度 联合分布 (Clustered Bar)
 4. 车型分布 (Pie)
 5. 联系方式完整率 (Pie)
 6. 提交时间趋势 (Line)
 7. 接单时间偏好分布 (Bar)
 8. 补充信息词云 (WordCloud)
 9. 地区排名柱状图 (Bar)
"""

import pandas as pd
import os
from pyecharts import options as opts
from pyecharts.charts import Bar, Pie, Line, TreeMap, WordCloud

# Excel 文件路径
EXCEL_PATH = r"E:\kabuda_data_analysis\司机信息数据库\司机信息.xlsx"
# 图表输出目录
OUTPUT_DIR = r"E:\kabuda_data_analysis\司机信息数据库\visualization_chart"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def load_sheets(path):
    xls = pd.ExcelFile(path)
    return {name: pd.read_excel(xls, name) for name in xls.sheet_names}

# 拆分多区域字段
def explode_regions(df, col="活动地区"):
    s = df[col].fillna("").astype(str).str.split(r"[，,；;]")
    return df.assign(**{col: s}).explode(col)

# 保存图表
def save_chart(chart, name):
    out = os.path.join(OUTPUT_DIR, f"{name}.html")
    chart.render(out)
    print(f"Saved {out}")

# 1. 司机等级分布
def chart_driver_level(df_main):
    cnt = df_main['司机等级'].value_counts().to_dict()
    bar = (
        Bar()
        .add_xaxis(list(cnt.keys()))
        .add_yaxis("人数", list(cnt.values()))
        .reversal_axis()
        .set_global_opts(
            title_opts=opts.TitleOpts(title="司机等级分布"),
            xaxis_opts=opts.AxisOpts(name="人数"),
            yaxis_opts=opts.AxisOpts(name="等级")
        )
    )
    save_chart(bar, "driver_level_distribution")

# 2. 区域覆盖 TreeMap
def chart_region_treemap(df_main):
    df = explode_regions(df_main)
    cnt = df['活动地区'].value_counts().reset_index()
    cnt.columns = ["name", "value"]
    treemap = (
        TreeMap()
        .add("", cnt.to_dict(orient="records"))
        .set_global_opts(title_opts=opts.TitleOpts(title="区域覆盖 TreeMap"))
    )
    save_chart(treemap, "region_treemap")

# 3. 职业 vs 活跃度
def chart_profession_activity(df_business):
    df = df_business.copy()
    df['活跃度标签'] = df['活跃度标签'].fillna("未知")
    pivot = df.groupby(['职业', '活跃度标签']).size().unstack(fill_value=0)
    bar = Bar().add_xaxis(list(pivot.index))
    for col in pivot.columns:
        bar.add_yaxis(col, list(pivot[col]))
    bar.set_global_opts(
        title_opts=opts.TitleOpts(title="职业 vs 活跃度"),
        xaxis_opts=opts.AxisOpts(name="职业", axislabel_opts=opts.LabelOpts(rotate=30)),
        yaxis_opts=opts.AxisOpts(name="人数"),
        legend_opts=opts.LegendOpts(pos_top="5%")
    )
    save_chart(bar, "profession_activity")

# 4. 车型分布
def chart_vehicle_pie(df_vehicle):
    cnt = df_vehicle['车型'].value_counts().to_dict()
    data = [list(item) for item in cnt.items()]
    pie = (
        Pie()
        .add("", data)
        .set_global_opts(title_opts=opts.TitleOpts(title="车型分布"))
        .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {d}%"))
    )
    save_chart(pie, "vehicle_pie")

# 5. 联系方式完整率
def chart_contact_complete(df_contact):
    cnt = df_contact['联系方式完整'].fillna("未知").value_counts().to_dict()
    data = [list(item) for item in cnt.items()]
    pie = (
        Pie()
        .add("", data)
        .set_global_opts(title_opts=opts.TitleOpts(title="联系方式完整率"))
        .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {d}%"))
    )
    save_chart(pie, "contact_complete")

# 6. 提交时间趋势
def chart_submission_trend(df_main):
    df = df_main.copy()
    df['提交时间'] = pd.to_datetime(df['提交时间'])
    ts = df.set_index('提交时间').resample('W').size()
    line = (
        Line()
        .add_xaxis([t.strftime("%Y-%m-%d") for t in ts.index])
        .add_yaxis("新增司机数", ts.values, is_smooth=True)
        .set_global_opts(
            title_opts=opts.TitleOpts(title="提交时间趋势"),
            xaxis_opts=opts.AxisOpts(name="周", type_="category", boundary_gap=False),
            yaxis_opts=opts.AxisOpts(name="数量")
        )
    )
    save_chart(line, "submission_trend")

# 7. 接单时间偏好
def chart_order_pref(df_business):
    cnt = df_business['接单时间'].fillna("未知").value_counts().to_dict()
    bar = (
        Bar()
        .add_xaxis(list(cnt.keys()))
        .add_yaxis("人数", list(cnt.values()))
        .set_global_opts(
            title_opts=opts.TitleOpts(title="接单时间偏好分布"),
            xaxis_opts=opts.AxisOpts(name="偏好", axislabel_opts=opts.LabelOpts(rotate=30)),
            yaxis_opts=opts.AxisOpts(name="人数")
        )
    )
    save_chart(bar, "order_preference")

# 8. 补充信息词云
def chart_extra_wordcloud(df_extra):
    text_cols = [c for c in df_extra.columns if df_extra[c].dtype == object]
    texts = df_extra[text_cols].fillna("").agg(" ".join, axis=1)
    all_text = " ".join(texts).split()
    wc = (
        WordCloud()
        .add("", [list(item) for item in pd.Series(all_text).value_counts().head(100).items()], word_size_range=[20, 100])
        .set_global_opts(title_opts=opts.TitleOpts(title="补充信息词云"))
    )
    save_chart(wc, "extra_wordcloud")

# 9. 地区排名柱状图
def chart_region_ranking(df_region):
    df = df_region.sort_values("司机数量", ascending=False)
    bar = (
        Bar()
        .add_xaxis(df['活动地区'].tolist())
        .add_yaxis("司机数量", df['司机数量'].tolist())
        .set_global_opts(
            title_opts=opts.TitleOpts(title="地区排名柱状图"),
            xaxis_opts=opts.AxisOpts(name="区域", axislabel_opts=opts.LabelOpts(rotate=30)),
            yaxis_opts=opts.AxisOpts(name="司机数量")
        )
    )
    save_chart(bar, "region_ranking")

# 主函数
def main():
    sheets = load_sheets(EXCEL_PATH)
    df_main     = sheets.get("司机全量档案", pd.DataFrame())
    df_business = sheets.get("业务信息_司机", pd.DataFrame())
    df_vehicle  = sheets.get("车辆信息_司机", pd.DataFrame())
    df_contact  = sheets.get("联系方式_司机", pd.DataFrame())
    df_extra    = sheets.get("补充信息_司机", pd.DataFrame())
    df_region   = sheets.get("地区统计", pd.DataFrame())

    chart_driver_level(df_main)
    chart_region_treemap(df_main)
    chart_profession_activity(df_business)
    chart_vehicle_pie(df_vehicle)
    chart_contact_complete(df_contact)
    chart_submission_trend(df_main)
    chart_order_pref(df_business)
    chart_extra_wordcloud(df_extra)
    chart_region_ranking(df_region)

if __name__ == "__main__":
    main()
