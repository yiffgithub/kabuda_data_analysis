#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
driver_business_viz.py
-----------------------
使用 pandas 和 pyecharts 对 “司机业务_统计分析.xlsx” 中各个 sheet 数据进行预处理，
并按业务维度生成专业标准化图表：
 1. 业务类型分布 (Bar)
 2. 业务类型细分分布 (TreeMap)
 3. 区域分布 (TreeMap)
 4. 金额区间分布 (Bar)
 5. 订单状态分布 (Pie)
 6. 评分区间分布 (Bar)
 7. 迟到分布 (Pie)
 8. 流向统计 (Sankey)
 9. 业务类型对比 (Sankey)
10. 明细表前10笔高单价订单 (Bar)
"""

import os
import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Bar, Pie, TreeMap, Sankey

# 文件路径配置
EXCEL_PATH = r"E:\kabuda_data_analysis\司机业务信息库\司机业务_统计分析.xlsx"
OUTPUT_DIR = r"E:\kabuda_data_analysis\司机业务信息库\visualization_chart"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 加载所有 sheet

def load_sheets(path):
    xls = pd.ExcelFile(path)
    return {sheet: pd.read_excel(xls, sheet) for sheet in xls.sheet_names}

# 保存图表到 HTML

def save_chart(chart, filename):
    filepath = os.path.join(OUTPUT_DIR, f"{filename}.html")
    chart.render(filepath)
    print(f"Saved: {filepath}")

# 1. 业务类型分布 (Bar)

def chart_type_dist(df):
    names = df.iloc[:, 0].astype(str).tolist()
    values = df.iloc[:, 1].tolist()
    bar = (
        Bar()
        .add_xaxis(names)
        .add_yaxis("数量", values)
        .set_global_opts(
            title_opts=opts.TitleOpts(title="业务类型分布"),
            xaxis_opts=opts.AxisOpts(name="业务类型", axislabel_opts=opts.LabelOpts(rotate=30)),
            yaxis_opts=opts.AxisOpts(name="数量")
        )
    )
    save_chart(bar, "业务类型分布")

# 2. 业务类型细分分布 (TreeMap)

def chart_type_sub_dist(df):
    df2 = df.rename(columns={df.columns[0]:"name", df.columns[1]:"value"})[["name","value"]]
    treemap = (
        TreeMap()
        .add("", df2.to_dict(orient="records"))
        .set_global_opts(title_opts=opts.TitleOpts(title="业务类型细分分布"))
    )
    save_chart(treemap, "业务类型细分分布")

# 3. 区域分布 (TreeMap)

def chart_region(df):
    df2 = df.rename(columns={df.columns[0]:"name", df.columns[1]:"value"})[["name","value"]]
    treemap = (
        TreeMap()
        .add("", df2.to_dict(orient="records"))
        .set_global_opts(title_opts=opts.TitleOpts(title="区域分布"))
    )
    save_chart(treemap, "区域分布")

# 4. 金额区间分布 (Bar)

def chart_amount_range(df):
    names = df.iloc[:, 0].astype(str).tolist()
    values = df.iloc[:, 1].tolist()
    bar = (
        Bar()
        .add_xaxis(names)
        .add_yaxis("数量", values)
        .set_global_opts(
            title_opts=opts.TitleOpts(title="金额区间分布"),
            xaxis_opts=opts.AxisOpts(name="金额区间", axislabel_opts=opts.LabelOpts(rotate=30)),
            yaxis_opts=opts.AxisOpts(name="数量")
        )
    )
    save_chart(bar, "金额区间分布")

# 5. 订单状态分布 (Pie)

def chart_order_status(df):
    data = [[str(row[df.columns[0]]), row[df.columns[1]]] for _, row in df.iterrows()]
    pie = (
        Pie()
        .add("", data)
        .set_global_opts(title_opts=opts.TitleOpts(title="订单状态分布"))
        .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {d}%"))
    )
    save_chart(pie, "订单状态分布")

# 6. 评分区间分布 (Bar)

def chart_rating(df):
    names = df.iloc[:, 0].astype(str).tolist()
    values = df.iloc[:, 1].tolist()
    bar = (
        Bar()
        .add_xaxis(names)
        .add_yaxis("数量", values)
        .set_global_opts(
            title_opts=opts.TitleOpts(title="评分区间分布"),
            xaxis_opts=opts.AxisOpts(name="评分区间"),
            yaxis_opts=opts.AxisOpts(name="数量")
        )
    )
    save_chart(bar, "评分区间分布")

# 7. 迟到分布 (Pie)

def chart_late(df):
    data = [[str(row[df.columns[0]]), row[df.columns[1]]] for _, row in df.iterrows()]
    pie = (
        Pie()
        .add("", data)
        .set_global_opts(title_opts=opts.TitleOpts(title="迟到分布"))
        .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {d}%"))
    )
    save_chart(pie, "迟到分布")

# 8. 流向统计 (Sankey)

def chart_flow(df, min_value: int = 3) -> None:
    """
    绘制 Sankey 流向图
    ----------
    Parameters
    ----------
    df : DataFrame
        必须含列【起点归一, 终点归一, 订单数】
    min_value : int, default=3
        过滤阈值，保留 value ≥ min_value 的流向，避免节点爆炸
    """
    src, tgt, val = '起点归一', '终点归一', '订单数'

    # 1️⃣ 清洗 & 去除自循环
    flow = (df[[src, tgt, val]]
            .dropna()
            .assign(**{
                src: lambda d: d[src].astype(str).str.strip(),
                tgt: lambda d: d[tgt].astype(str).str.strip(),
                val: lambda d: pd.to_numeric(d[val], errors='coerce')
            })
            .dropna()
            .loc[lambda d: d[src] != d[tgt]]         # 去自循环
            .groupby([src, tgt], as_index=False)[val]
            .sum()
            .loc[lambda d: d[val] >= min_value])     # 过滤小流量

    if flow.empty:
        print('>>> 流向统计：无有效数据，跳过绘图')
        return

    # 2️⃣ 构造节点 & 链接
    nodes = [{'name': n} for n in { *flow[src], *flow[tgt] }]
    links = flow.rename(columns={src: 'source', tgt: 'target'})[['source', 'target', val]] \
                .to_dict(orient='records')

    # 3️⃣ 绘制 Sankey
    sankey = (
        Sankey(init_opts=opts.InitOpts(width='1200px', height='600px'))
        .add(
            series_name='流向',
            nodes=nodes,
            links=links,
            orient='horizontal',
            node_align='left',
            node_width=28,
            node_gap=14,
            linestyle_opt=opts.LineStyleOpts(
                color='source',            # 线条颜色继承 source 节点
                opacity=0.8,
                width=2,
                curve=0.45
            ),
            label_opts=opts.LabelOpts(position='right', font_size=12)
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(title='流向统计', pos_left='center'),
            tooltip_opts=opts.TooltipOpts(trigger='item', formatter='{b}: {c}')
        )
    )
    save_chart(sankey, '流向统计')





# 9. 业务类型对比 (Sankey)

def chart_type_compare(df):
    src_col = '结构化业务类型'
    tgt_col = '订单类型'
    cnt = df.groupby([src_col, tgt_col]).size().reset_index(name='value')
    nodes = list(set(cnt[src_col]).union(cnt[tgt_col]))
    node_list = [{"name": n} for n in nodes]
    links = cnt.rename(columns={src_col:'source', tgt_col:'target'})[['source','target','value']].to_dict(orient='records')
    sankey = (
        Sankey()
        .add("", node_list, links, label_opts=opts.LabelOpts(position="right"))
        .set_global_opts(title_opts=opts.TitleOpts(title="业务类型对比"))
    )
    save_chart(sankey, "业务类型对比")

# 10. 明细表前10笔高单价订单 (Bar)

def chart_top_orders(df):
    df_sorted = df.sort_values('订单销售价', ascending=False).head(10)
    names = df_sorted['订单编号'].astype(str).tolist()
    values = df_sorted['订单销售价'].tolist()
    bar = (
        Bar()
        .add_xaxis(names)
        .add_yaxis("销售价", values)
        .set_global_opts(
            title_opts=opts.TitleOpts(title="前10高单价订单"),
            xaxis_opts=opts.AxisOpts(name="订单编号", axislabel_opts=opts.LabelOpts(rotate=30)),
            yaxis_opts=opts.AxisOpts(name="销售价(元)")
        )
    )
    save_chart(bar, "前10高单价订单")

# 主流程

def main():
    sheets = load_sheets(EXCEL_PATH)
    chart_type_dist(sheets['业务类型分布'])
    chart_type_sub_dist(sheets['业务类型分布_细分'])
    chart_region(sheets['区域分布'])
    chart_amount_range(sheets['金额区间分布'])
    chart_order_status(sheets['订单状态分布'])
    chart_rating(sheets['评分区间分布'])
    chart_late(sheets['迟到分布'])
    chart_flow(sheets['流向统计'])
    chart_type_compare(sheets['业务类型对比'])
    chart_top_orders(sheets['明细全表'])

if __name__ == '__main__':
    main()