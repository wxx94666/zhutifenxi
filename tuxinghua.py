# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import load_workbook

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

# ========= TopN 横向柱状图（统一水平数值，最大在最上） =========
def topn_chart(df_subject, group_letter, value_letter, title, how="mean", n=10, is_percent=False):
    tmp = df_subject.copy()
    tmp.iloc[:, LETTER_IDX[value_letter]] = coerce_numeric(col(tmp, value_letter), treat_percent=is_percent)
    agg = group_metric(tmp, group_letter, value_letter, how=how)
    agg = agg.rename(columns={agg.columns[0]: "主体", agg.columns[1]: "值"}).copy()
    agg["值"] = pd.to_numeric(agg["值"], errors="coerce").fillna(0)

    agg = agg.sort_values("值", ascending=False).head(n)
    order = agg["主体"].tolist()

    if is_percent:
        text_vals = (agg["值"] * 100).round(2).astype(str) + "%"
        x_vals = agg["值"]
    else:
        text_vals = agg["值"].round(2).astype(str)
        x_vals = agg["值"]

    fig = px.bar(
        agg, x=x_vals, y="主体", orientation="h",
        title=title, text=text_vals
    )
    fig.update_traces(textposition="inside", textangle=0, insidetextanchor="middle")
    fig.update_layout(
        yaxis=dict(categoryorder="array", categoryarray=order, autorange="reversed"),
        margin=dict(l=120, r=20, t=60, b=30),
        height=420,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)"
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
    if percent_mode:
        denom = float(total_denominator) if total_denominator else 0.0
        dfp["值"] = dfp["值"].astype(float).apply(lambda v: (v / denom * 100.0) if denom > 0 else 0.0)
    fig = px.pie(dfp, values="值", names="主体", title=title, hole=0.35)
    fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
    return fig

# ========= 排行榜 =========
def account_rank_table(df_account):
    status = col(df_account, "H").astype(str).str.strip().str.lower()
    mask = status.isin(["active", "need to pay"])
    sub = df_account[mask].copy()
    sub["排序值"] = pd.to_numeric(col(sub, "J"), errors="coerce")
    sub = sub.sort_values("排序值", ascending=False).head(10)
    out = pd.DataFrame({
        "客户名称": col(sub, "C").astype(str),
        "账户名称": col(sub, "D").astype(str),
        "账户ID": col(sub, "E").astype(str),
        "一周消耗": col(sub, "J")
    })
    return out

# ========= 所有主体状态详情（严格 100% 堆叠 + 右侧单独一列显示 C） =========
def all_subjects_status(df_subject):
    sub = df_subject.copy()
    for l in ["B","C","D","E","F","G","H"]:
        if l != "B":
            sub[l] = coerce_numeric(col(sub, l), treat_percent=False)
        else:
            sub[l] = col(sub, l).astype(str)

    agg = sub.groupby("B").agg({
        "C":"sum","D":"sum","E":"sum","F":"sum","G":"sum","H":"sum"
    }).reset_index().rename(columns={"B":"主体"})

    # 先处理溢出：若 D+E+F+G+H > C，扣 G
    totals = agg["C"].astype(float)
    total_states = agg["D"]+agg["E"]+agg["F"]+agg["G"]+agg["H"]
    overflow = (total_states - totals).clip(lower=0)
    agg["G"] = (agg["G"] - overflow).clip(lower=0)

    # 计算百分比（注意 0 总数）
    denom = totals.replace(0, np.nan)
    pct_cols = {}
    for l in ["D","E","F","G","H"]:
        pct_cols[l] = (agg[l] / denom * 100).fillna(0)

    # 为了保证每条横条长度一致，增加一个“填充”分块补齐到 100%
    stacked_sum = pct_cols["D"] + pct_cols["E"] + pct_cols["F"] + pct_cols["G"] + pct_cols["H"]
    filler_pct = (100 - stacked_sum).clip(lower=0)  # 防止负数

    cats = ["D","E","F","G","H"]
    names = {"D":"已交付","E":"可交付","F":"未绑卡","G":"问题户","H":"死户"}
    colors = {"D":"#1f77b4","E":"#2ca02c","F":"#ff7f0e","G":"#9467bd","H":"#d62728"}

    # 按总数降序排列（可改为原始顺序）
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
            text=agg[c].astype(int),           # 实际数量
            textposition="inside",
            insidetextanchor="middle",
            textangle=0,
            cliponaxis=False
        ))

    # 透明“填充”分块，保证横条总长=100
    fig.add_trace(go.Bar(
        y=agg["主体"],
        x=agg["FILLER"],
        orientation="h",
        name="",
        marker=dict(color="rgba(0,0,0,0)"),
        hoverinfo="skip",
        showlegend=False
    ))

    # 右侧单独一列：显示 C（总账户数），白色加粗，所有主体对齐
    fig.add_trace(go.Scatter(
        y=agg["主体"],
        x=[110]*len(agg),                      # 固定在100%右侧
        mode="text",
        text=[f"<b>{int(v)}</b>" for v in agg["C"]],
        textposition="middle left",
        textfont=dict(color="#FFFFFF", size=14),
        showlegend=False,
        hoverinfo="skip",
        cliponaxis=False
    ))

    # 右上角列头
    fig.add_annotation(
        x=110, y=1.02, xref="x", yref="paper",
        text="<b>总账户数</b>",
        showarrow=False,
        font=dict(color="#FFFFFF", size=14),
        xanchor="left", yanchor="bottom"
    )

    fig.update_layout(
        barmode="stack",
        title="所有主体状态详情",
        xaxis=dict(showticklabels=False, showgrid=False, zeroline=False, range=[0,125]),
        yaxis=dict(title="主体", categoryorder="array", categoryarray=order),
        margin=dict(l=130, r=160, t=70, b=40),
        height=max(520, 26*len(agg)),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", yanchor="bottom", y=1.04, xanchor="left", x=0)
    )
    return fig

# ========= Streamlit App =========
st.set_page_config(page_title="主体分析可视化", layout="wide")
st.markdown(
    """
    <style>
    .big-title { font-size: 28px; font-weight: 700; margin-bottom: 6px; }
    .kpi .stMetric { padding: 8px 12px; border-radius: 12px; }
    </style>
    """, unsafe_allow_html=True
)
st.markdown('<div class="big-title">主体分析可视化看板</div>', unsafe_allow_html=True)

EXCEL_FILE = "主体分析.xlsx"
df_subject = pd.read_excel(EXCEL_FILE, sheet_name="主体分析表", header=0)
df_account = pd.read_excel(EXCEL_FILE, sheet_name="账户表", header=0)

ts_str, total_accounts, delivered_accounts, unbound_cards, avg_dead_rate = build_kpis(df_subject, EXCEL_FILE)
st.markdown(f"**数据产生时间：** {ts_str}")

st.subheader("数据概览")
with st.container():
    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.metric("总账户数", f"{int(total_accounts):,}")
    with k2:
        st.metric("已交付账户数", f"{int(delivered_accounts):,}")
    with k3:
        st.metric("未绑卡账户数", f"{int(unbound_cards):,}")
    with k4:
        st.metric("平均死户率", f"{avg_dead_rate:.2f}%")
    with k5:
        st.metric(" ", " ")

st.subheader("榜单 Top10")
c1, c2 = st.columns(2)
with c1:
    st.plotly_chart(topn_chart(df_subject, "B", "Q", "主体综合评分榜单（Q）", how="mean", n=10, is_percent=False), use_container_width=True)
    st.plotly_chart(topn_chart(df_subject, "B", "L", "主体平均消耗榜单（L）", how="mean", n=10, is_percent=False), use_container_width=True)
with c2:
    st.plotly_chart(topn_chart(df_subject, "B", "P", "主体好户榜单（P）", how="max", n=10, is_percent=True), use_container_width=True)
    st.plotly_chart(topn_chart(df_subject, "B", "I", "主体死户率榜单（I）", how="max", n=10, is_percent=True), use_container_width=True)

st.subheader("占比分析")
total_accounts_all = safe_sum(col(df_subject, "C"))
c3, c4 = st.columns(2)
with c3:
    st.plotly_chart(pie_from_filter(df_subject, "内部户表", "B", "C", "内部户主体账户数占比"), use_container_width=True)
    st.plotly_chart(pie_from_filter(df_subject, "内部户表", "B", "O", "内部户主体消耗占比"), use_container_width=True)
    st.plotly_chart(pie_from_filter(df_subject, "内部户表", "B", "H", "内部户主体死户数占比（占总账户数%）", percent_mode=True, total_denominator=total_accounts_all), use_container_width=True)
with c4:
    st.plotly_chart(pie_from_filter(df_subject, "外部户表", "B", "C", "外部户主体账户数占比"), use_container_width=True)
    st.plotly_chart(pie_from_filter(df_subject, "外部户表", "B", "O", "外部户主体消耗占比"), use_container_width=True)
    st.plotly_chart(pie_from_filter(df_subject, "外部户表", "B", "H", "外部户主体死户数占比（占总账户数%）", percent_mode=True, total_denominator=total_accounts_all), use_container_width=True)

st.subheader("账户排行榜（Top 10）")
rank_df = account_rank_table(df_account)
st.dataframe(rank_df, use_container_width=True)

st.subheader("所有主体状态详情")
st.plotly_chart(all_subjects_status(df_subject), use_container_width=True)
