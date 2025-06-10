#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
draw_sankey.py
--------------
根据 “起点 终点 订单数” 三列数据绘制桑基图
"""

import pandas as pd
import io
from pyecharts import options as opts
from pyecharts.charts import Sankey

# ------------------------------------------------------------------
# 1) 准备数据
# ------------------------------------------------------------------
# 方式 A：手动粘贴（适合快速测试）
data = """
起点	终点	订单数
密西沙加	密西沙加	430
多伦多	多伦多	26
皮尔逊机场	多伦多	24
万锦	多伦多	19
北约克	多伦多	18
多伦多	皮尔逊机场	18
万锦	士嘉堡	17
北约克	皮尔逊机场	14
士嘉堡	多伦多	14
皮尔逊机场	北约克	13
士嘉堡	士嘉堡	13
万锦	万锦	13
北约克	万锦	12
多伦多	北约克	12
士嘉堡	万锦	12
北约克	北约克	12
士嘉堡	北约克	10
多伦多	士嘉堡	10
北约克	士嘉堡	9
士嘉堡	皮尔逊机场	9
密西沙加	万锦	7
士嘉堡	列治文山	7
士嘉堡	密西沙加	7
北约克	密西沙加	6
列治文山	万锦	6
列治文山	多伦多	6
多伦多	万锦	6
万锦	皮尔逊机场	6
密西沙加	多伦多	6
万锦	北约克	6
多伦多	密西沙加	5
万锦	密西沙加	5
多伦多	列治文山	4
汉密尔顿	多伦多	4
密西沙加	皮尔逊机场	4
皮尔逊机场	密西沙加	4
皮尔逊机场	士嘉堡	4
皮尔逊机场	万锦	4
列治文山	北约克	3
万锦	贝尔维尔	3
汉密尔顿	万锦	3
列治文山	士嘉堡	3
士嘉堡	汉密尔顿	2
万锦	汉密尔顿	2
皮尔逊机场	哈密尔顿	2
滑铁卢	万锦	2
列治文山	密西沙加	2
列治文山	皮尔逊机场	2
北约克	列治文山	2
怡陶碧谷	怡陶碧谷	2
密西沙加	士嘉堡	2
万锦	列治文山	2
汉密尔顿	密西沙加	2
哈密尔顿	北约克	2
旺市	皮尔逊机场	1
皮尔逊机场	纽马克特	1
贝尔维尔	万锦	1
旺市	多伦多	1
瓦克沃	多伦多	1
皮尔逊机场	布雷斯布里奇	1
旺市（Vaughan）	多伦多	1
汉密尔顿	士嘉堡	1
皮尔逊机场	列治文山	1
瓦特福德	滑铁卢	1
尼亚加拉	皮尔逊机场	1
汉密尔顿	安卡斯特（Anc	1
瓦恩	旺市	1
Fort Erie	密西沙加	1
密西沙加	哈密尔顿	1
奥克维尔	密西沙加	1
万锦	哈密尔顿	1
万锦	多伦多、多伦	1
万锦	旺市	1
万锦	瑞士嘉	1
东约克	多伦多	1
伦敦	皮尔逊机场	1
伯灵顿	汉密尔顿	1
列治文山	列治文山	1
北约克	布兰普顿	1
北约克	布拉姆普顿	1
北约克	布雷斯布里奇	1
北约克	汉密尔顿	1
北约克	瑞士嘉	1
北约克	约克代尔	1
圣凯瑟	皮尔逊机场	1
基于提供的地址信息，	列治文山	1
基于提供的地址信息，	多伦多	1
士嘉堡	贝尔维尔	1
多伦多	Vaughan	1
多伦多	康科德	1
多伦多	纽马克特	1
贝尔维尔	皮尔逊机场	1
""".strip()

df_manual = pd.read_csv(io.StringIO(data), sep="\t")
df = df_manual          # 如果用 Excel，替换成 pd.read_excel(...)

# 若你直接读 Excel sheet: df = sheets['流向统计']

# ------------------------------------------------------------------
# 2) 过滤 & 处理
# ------------------------------------------------------------------
src, tgt, val = '起点', '终点', '订单数'   # 根据 df 列名修正
if all(col in df.columns for col in ['起点归一','终点归一']):
    src, tgt = '起点归一', '终点归一'      # 若已归一化列存在则使用

df_clean = (
    df[[src, tgt, val]]
    .dropna()
    .assign(**{
        src: lambda d: d[src].astype(str).str.strip(),
        tgt: lambda d: d[tgt].astype(str).str.strip(),
        val: lambda d: pd.to_numeric(d[val], errors='coerce')
    })
    .dropna()
    .loc[lambda d: d[src] != d[tgt]]          # 去自循环
    .groupby([src, tgt], as_index=False)[val]
    .sum()
    .loc[lambda d: d[val] >= 3]               # 阈值，可调
)

nodes = [{'name': n} for n in { *df_clean[src], *df_clean[tgt] }]
links = df_clean.rename(columns={src:'source', tgt:'target'})[['source','target',val]] \
               .to_dict(orient='records')

# ------------------------------------------------------------------
# 3) 绘制桑基图
# ------------------------------------------------------------------
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
        linestyle_opt=opts.LineStyleOpts(color='source', opacity=0.8, width=2, curve=0.45),
        label_opts=opts.LabelOpts(position='right', font_size=12)
    )
    .set_global_opts(
        title_opts=opts.TitleOpts(title='流向统计（桑基图）', pos_left='center'),
        tooltip_opts=opts.TooltipOpts(trigger='item', formatter='{b}: {c}')
    )
)

# 保存到当前目录
sankey.render("流向统计_桑基图.html")
print("✓ 已生成流向统计_桑基图.html")
