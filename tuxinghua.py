# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import load_workbook
import plotly.io as pio

# 设置Plotly默认主题
pio.templates.default = "plotly_white"

# ========= 基础工具 =========
LETTER_IDX = {c: i for i, c in enumerate(list("ABCDEFGHIJKLMNOPQRSTUVWXYZ"))}

def col(df, letter):
    return df.iloc[:, LETTER_IDX[letter]]

def safe_sum(series):
    return float(pd.to_numeric(series, errors="coerce").fillna(0).sum())

def to_pct(x, denom):
    try:
        return (float(x) / float(denom)) * 100 if denom else 0.0
    except Exception:
        return 0.0

def coerce_numeric(series, treat_percent=False):
    s = series.astype(str).str.strip().replace({"None": np.nan, "nan": np.nan})
    s = s.str.replace("%", "", regex=False)
    vals = pd.to_numeric(s, errors="coerce")
    if treat_percent and (vals.dropna() > 1).any():
        vals = vals / 100.0
    return vals.fillna(0)

def group_metric(df, group_letter, value_letter, how="mean"):
    g = df.groupby(col(df, group_letter))
    if how == "sum":
        s = g.apply(lambda x: pd.to_numeric(col(x, value_letter), errors="coerce").fillna(0).sum())
    elif how == "max":
        s = g.apply(lambda x: pd.to_numeric(col(x, value_letter), errors="coerce").fillna(0).max())
    else:
        s = g.apply(lambda x: pd.to_numeric(col(x, value_letter), errors="coerce").dropna().mean())
    s = s.replace([np.inf, -np.inf], np.nan).fillna(0)
    return s.reset_index().rename(columns={0: "value", s.index.name: "主体"})

# ========= 读取 X1 =========
def read_x1_text(excel_path):
    candidates = ["数据表", "主体分析表"]
    try:
        wb = load_workbook(excel_path, data_only=True)
        for name in candidates:
            if name in wb.sheetnames:
                v = wb[name]["X1"].value
                if v is not None:
                    return str(v).strip()
    except Exception:
        pass
    for name in candidates:
        try:
            df_raw = pd.read_excel(excel_path, sheet_name=name, header=None)
            return str(df_raw.iat[0, LETTER_IDX["X"]]).strip()
        except Exception:
            continue
    return "未提供"

# ========= KPI =========
def build_kpis(df_subject, excel_path):
    ts_str = read_x1_text(excel_path)
    total_accounts = safe_sum(col(df_subject, "C"))
    delivered_accounts = safe_sum(col(df_subject, "D"))
    unbound_cards = safe_sum(col(df_subject, "F"))
    dead_count = safe_sum(col(df_subject, "H"))
    avg_dead_rate = to_pct(dead_count, total_accounts)
    return ts_str, total_accounts, delivered_accounts, unbound_cards, avg_dead_rate

# ========= TopN 横向柱状图（增加排序控制参数，默认保持原有降序） =========
def topn_chart(df_subject, group_letter, value_letter, title, how="mean", n=10, is_percent=False, sort_ascending=False):
    tmp = df_subject.copy()
    tmp.iloc[:, LETTER_IDX[value_letter]] = coerce_numeric(col(tmp, value_letter), treat_percent=is_percent)
    agg = group_metric(tmp, group_letter, value_letter, how=how)
    agg = agg.rename(columns={agg.columns[0]: "主体", agg.columns[1]: "值"}).copy()
    agg["值"] = pd.to_numeric(agg["值"], errors="coerce").fillna(0)

    # 核心修改：根据sort_ascending参数控制排序方向（True=升序，False=降序）
    agg = agg.sort_values("值", ascending=sort_ascending).head(n)
    order = agg["主体"].tolist()

    if is_percent:
        text_vals = (agg["值"] * 100).round(2).astype(str) + "%"
        x_vals = agg["值"]
    else:
        text_vals = agg["值"].round(2).astype(str)
        x_vals = agg["值"]

    # 保持原有美观颜色方案
    fig = px.bar(
        agg, x=x_vals, y="主体", orientation="h",
        title=title, text=text_vals,
        color="值",
        color_continuous_scale="Viridis",
        height=450
    )
    fig.update_traces(
        textposition="inside", 
        textangle=0, 
        insidetextanchor="middle",
        marker_line_width=0  # 去除边框
    )
    fig.update_layout(
        yaxis=dict(categoryorder="array", categoryarray=order, autorange="reversed"),
        margin=dict(l=150, r=30, t=70, b=40),
        coloraxis_showscale=False,  # 隐藏颜色条
        title_font=dict(size=18),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        hovermode="y unified"  # 统一行悬停
    )
    return fig

# ========= 饼图 =========
def pie_from_filter(df_subject, filter_value, group_letter, value_letter, title, percent_mode=False, total_denominator=None):
    sub = df_subject[col(df_subject, "A") == filter_value].copy()
    sub.iloc[:, LETTER_IDX[value_letter]] = coerce_numeric(col(sub, value_letter), treat_percent=False)
    grouped = sub.groupby(col(sub, group_letter)).apply(lambda x: x.iloc[:, LETTER_IDX[value_letter]].sum())
    grouped = grouped.replace([np.inf, -np.inf], np.nan).fillna(0)
    dfp = grouped.reset_index()
    dfp.columns = ["主体", "值"]
    
    # 过滤掉0值，使饼图更清晰
    dfp = dfp[dfp["值"] > 0]
    
    if percent_mode:
        denom = float(total_denominator) if total_denominator else 0.0
        dfp["值"] = dfp["值"].astype(float).apply(lambda v: (v / denom * 100.0) if denom > 0 else 0.0)
    
    # 保持原有美观饼图样式
    fig = px.pie(
        dfp, 
        values="值", 
        names="主体", 
        title=title, 
        hole=0.4,
        color_discrete_sequence=px.colors.qualitative.Pastel
    )
    fig.update_traces(
        textinfo="label+percent",  # 显示标签和百分比
        textposition="inside",
        insidetextorientation="radial",
        marker=dict(line=dict(color="#FFFFFF", width=2))  # 增加白色边框
    )
    fig.update_layout(
        title_font=dict(size=18),
        margin=dict(l=20, r=20, t=70, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    return fig

# ========= 排行榜 =========
def account_rank_table(df_account):
    status = col(df_account, "H").astype(str).str.strip().str.lower()
    mask = status.isin(["active", "need to pay"])
    sub = df_account[mask].copy()
    sub["排序值"] = pd.to_numeric(col(sub, "J"), errors="coerce")
    sub = sub.sort_values("排序值", ascending=False).head(10)
    out = pd.DataFrame({
        "排名": range(1, len(sub)+1),  # 增加排名列
        "客户名称": col(sub, "C").astype(str),
        "账户名称": col(sub, "D").astype(str),
        "账户ID": col(sub, "E").astype(str),
        "一周消耗": col(sub, "J")
    })
    return out.set_index("排名")  # 用排名作为索引

# ========= 所有主体状态详情 =========
def all_subjects_status(df_subject):
    sub = df_subject.copy()
    for l in ["B","C","D","E","F","G","H"]:
        if l != "B":
            sub[l] = coerce_numeric(col(sub, l), treat_percent=False)
        else:
            sub[l] = col(sub, "B").astype(str)

    agg = sub.groupby("B").agg({
        "C":"sum","D":"sum","E":"sum","F":"sum","G":"sum","H":"sum"
    }).reset_index().rename(columns={"B":"主体"})

    # 处理溢出
    totals = agg["C"].astype(float)
    total_states = agg["D"]+agg["E"]+agg["F"]+agg["G"]+agg["H"]
    overflow = (total_states - totals).clip(lower=0)
    agg["G"] = (agg["G"] - overflow).clip(lower=0)

    # 计算百分比
    denom = totals.replace(0, np.nan)
    pct_cols = {}
    for l in ["D","E","F","G","H"]:
        pct_cols[l] = (agg[l] / denom * 100).fillna(0)

    # 填充分块
    stacked_sum = pct_cols["D"] + pct_cols["E"] + pct_cols["F"] + pct_cols["G"] + pct_cols["H"]
    filler_pct = (100 - stacked_sum).clip(lower=0)

    cats = ["D","E","F","G","H"]
    names = {"D":"已交付","E":"可交付","F":"未绑卡","G":"问题户","H":"死户"}
    # 更协调的颜色方案
    colors = {
        "D":"#3498db",  # 蓝色
        "E":"#2ecc71",  # 绿色
        "F":"#f39c12",  # 橙色
        "G":"#9b59b6",  # 紫色
        "H":"#e74c3c"   # 红色
    }

    agg = agg.assign(
        D_pct=pct_cols["D"], E_pct=pct_cols["E"], F_pct=pct_cols["F"],
        G_pct=pct_cols["G"], H_pct=pct_cols["H"], FILLER=filler_pct
    )
    agg = agg.sort_values("C", ascending=False).reset_index(drop=True)
    order = agg["主体"].tolist()

    fig = go.Figure()

    # 五个实际分块
    for c in cats:
        fig.add_trace(go.Bar(
            y=agg["主体"],
            x=agg[f"{c}_pct"],
            orientation="h",
            name=names[c],
            marker=dict(color=colors[c]),
            text=agg[c].astype(int),
            textposition="inside",
            insidetextanchor="middle",
            textangle=0,
            cliponaxis=False,
            marker_line_width=0  # 去除边框
        ))

    # 透明填充分块
    fig.add_trace(go.Bar(
        y=agg["主体"],
        x=agg["FILLER"],
        orientation="h",
        name="",
        marker=dict(color="rgba(0,0,0,0)"),
        hoverinfo="skip",
        showlegend=False
    ))

    # 右侧总账户数
    fig.add_trace(go.Scatter(
        y=agg["主体"],
        x=[110]*len(agg),
        mode="text",
        text=[f"<b>{int(v)}</b>" for v in agg["C"]],
        textposition="middle left",
        textfont=dict(color="#333333", size=14),  # 深灰色更易读
        showlegend=False,
        hoverinfo="skip",
        cliponaxis=False
    ))

    # 右上角列头
    fig.add_annotation(
        x=110, y=1.02, xref="x", yref="paper",
        text="<b>总账户数</b>",
        showarrow=False,
        font=dict(color="#333333", size=14),
        xanchor="left", yanchor="bottom"
    )

    fig.update_layout(
        barmode="stack",
        title="所有主体状态详情",
        title_font=dict(size=18),
        xaxis=dict(showticklabels=False, showgrid=False, zeroline=False, range=[0,125]),
        yaxis=dict(title="主体", categoryorder="array", categoryarray=order),
        margin=dict(l=150, r=180, t=70, b=40),
        height=max(550, 30*len(agg)),  # 增加行高
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(
            orientation="h", 
            yanchor="bottom", 
            y=1.04, 
            xanchor="center", 
            x=0.5  # 居中显示图例
        ),
        hovermode="y unified"  # 统一行悬停
    )
    return fig

# ========= Streamlit App =========
st.set_page_config(page_title="主体分析可视化", layout="wide")

# 自定义CSS样式优化（保持不变）
st.markdown(
    """
    <style>
    .big-title { 
        font-size: 32px; 
        font-weight: 700; 
        margin-bottom: 15px; 
        color: #2c3e50;
    }
    .metric-container {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 15px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    .section-header {
        font-size: 20px;
        font-weight: 600;
        margin: 25px 0 15px 0;
        color: #34495e;
        border-left: 4px solid #3498db;
        padding-left: 10px;
    }
    .data-time {
        color: #7f8c8d;
        margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True
)

# 页面标题（保持不变）
st.markdown('<div class="big-title">Lucky主体分析看板</div>', unsafe_allow_html=True)

# 数据加载（保持不变）
EXCEL_FILE = "主体分析.xlsx"
try:
    df_subject = pd.read_excel(EXCEL_FILE, sheet_name="主体分析表", header=0)
    df_account = pd.read_excel(EXCEL_FILE, sheet_name="账户表", header=0)
except Exception as e:
    st.error(f"数据加载失败: {str(e)}")
    st.stop()

# 计算KPI（保持不变）
ts_str, total_accounts, delivered_accounts, unbound_cards, avg_dead_rate = build_kpis(df_subject, EXCEL_FILE)
st.markdown(f'<div class="data-time">数据产生时间：{ts_str}</div>', unsafe_allow_html=True)

# 数据概览区域（保持不变）
st.markdown('<div class="section-header">数据概览</div>', unsafe_allow_html=True)
with st.container():
    cols = st.columns(5)
    metrics = [
        ("总账户数", f"{int(total_accounts):,}"),
        ("已交付账户数", f"{int(delivered_accounts):,}"),
        ("未绑卡账户数", f"{int(unbound_cards):,}"),
        ("平均死户率", f"{avg_dead_rate:.2f}%"),
        ("账户活跃度", f"{to_pct(delivered_accounts, total_accounts):.2f}%")
    ]
    
    for i, (label, value) in enumerate(metrics):
        with cols[i]:
            st.markdown('<div class="metric-container">', unsafe_allow_html=True)
            st.metric(label, value)
            st.markdown('</div>', unsafe_allow_html=True)

# 榜单区域（仅修改死户率榜单调用参数）
st.markdown('<div class="section-header">榜单 Top10</div>', unsafe_allow_html=True)
c1, c2 = st.columns(2)
with c1:
    # 综合评分、平均消耗榜单保持原有降序（默认sort_ascending=False）
    st.plotly_chart(
        topn_chart(df_subject, "B", "Q", "主体综合评分榜单（Q）", how="mean", n=10, is_percent=False), 
        use_container_width=True
    )
    st.plotly_chart(
        topn_chart(df_subject, "B", "L", "主体平均消耗榜单（L）", how="mean", n=10, is_percent=False), 
        use_container_width=True
    )
with c2:
    # 好户榜单保持原有降序
    st.plotly_chart(
        topn_chart(df_subject, "B", "P", "主体好户榜单（P）", how="max", n=10, is_percent=True), 
        use_container_width=True
    )
    # 核心修改：死户率榜单设置sort_ascending=True（升序取最小前10），标题优化表述
    st.plotly_chart(
        topn_chart(
            df_subject, "B", "I", "主体死户率榜单（I）- 最低10名", 
            how="max", n=10, is_percent=True, 
            sort_ascending=True  # 升序排列，取死户率最小的前10
        ), 
        use_container_width=True
    )

# 占比分析区域（保持不变）
st.markdown('<div class="section-header">占比分析</div>', unsafe_allow_html=True)
total_accounts_all = safe_sum(col(df_subject, "C"))
c3, c4 = st.columns(2)
with c3:
    st.plotly_chart(
        pie_from_filter(df_subject, "内部户表", "B", "C", "内部户主体账户数占比"), 
        use_container_width=True
    )
    st.plotly_chart(
        pie_from_filter(df_subject, "内部户表", "B", "O", "内部户主体消耗占比"), 
        use_container_width=True
    )
    st.plotly_chart(
        pie_from_filter(
            df_subject, "内部户表", "B", "H", 
            "内部户主体死户数占比（占总账户数%）", 
            percent_mode=True, 
            total_denominator=total_accounts_all
        ), 
        use_container_width=True
    )
with c4:
    st.plotly_chart(
        pie_from_filter(df_subject, "外部户表", "B", "C", "外部户主体账户数占比"), 
        use_container_width=True
    )
    st.plotly_chart(
        pie_from_filter(df_subject, "外部户表", "B", "O", "外部户主体消耗占比"), 
        use_container_width=True
    )
    st.plotly_chart(
        pie_from_filter(
            df_subject, "外部户表", "B", "H", 
            "外部户主体死户数占比（占总账户数%）", 
            percent_mode=True, 
            total_denominator=total_accounts_all
        ), 
        use_container_width=True
    )

# 账户排行榜（保持不变）
st.markdown('<div class="section-header">账户排行榜（Top 10）</div>', unsafe_allow_html=True)
rank_df = account_rank_table(df_account)
st.dataframe(
    rank_df, 
    use_container_width=True,
    column_config={
        "一周消耗": st.column_config.NumberColumn(format="%.2f")
    }
)

# 主体状态详情（保持不变）
st.markdown('<div class="section-header">所有主体状态详情</div>', unsafe_allow_html=True)
st.plotly_chart(all_subjects_status(df_subject), use_container_width=True)

# 添加页脚（保持不变）
st.markdown(
    """
    <hr>
    <div style="text-align: center; color: #7f8c8d; font-size: 14px;">
        Lucky主体分析可视化系统 © 2025
    </div>
    """, 
    unsafe_allow_html=True
)
