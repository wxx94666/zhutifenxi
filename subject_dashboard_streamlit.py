# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

LETTER_IDX = {c: i for i, c in enumerate(list("ABCDEFGHIJKLMNOPQRSTUVWXYZ"))}

def safe_sum(series):
    return float(pd.to_numeric(series, errors="coerce").fillna(0).sum())

def to_pct(x, denom):
    try:
        return (float(x)/float(denom))*100 if denom else 0.0
    except:
        return 0.0

def col(df, letter):
    idx = LETTER_IDX[letter]
    return df.iloc[:, idx]

def group_metric(df, group_letter, value_letter, how="mean"):
    g = df.groupby(col(df, group_letter))
    if how == "sum":
        s = g.apply(lambda x: pd.to_numeric(col(x, value_letter), errors="coerce").fillna(0).sum())
    else:
        s = g.apply(lambda x: pd.to_numeric(col(x, value_letter), errors="coerce").dropna().mean())
    s = s.replace([np.inf, -np.inf], np.nan).fillna(0)
    return s.reset_index().rename(columns={0:"value", s.index.name:"主体"})

def build_kpis(df_subject):
    try:
        gen_time_raw = col(df_subject, "X").iloc[0]
        ts_str = str(gen_time_raw)
    except Exception:
        ts_str = "未提供"

    total_accounts = safe_sum(col(df_subject, "C"))
    delivered_accounts = safe_sum(col(df_subject, "D"))
    unbound_cards = safe_sum(col(df_subject, "F"))
    dead_count = safe_sum(col(df_subject, "H"))
    avg_dead_rate = to_pct(dead_count, total_accounts)

    return ts_str, total_accounts, delivered_accounts, unbound_cards, avg_dead_rate

def topn_chart(df_subject, group_letter, value_letter, title, how="mean", n=10, is_percent=False):
    agg = group_metric(df_subject, group_letter, value_letter, how=how)
    agg = agg.rename(columns={agg.columns[0]: "主体", agg.columns[1]: "值"})
    agg = agg.sort_values("值", ascending=False).head(n)
    fig = px.bar(
        agg.sort_values("值", ascending=True),
        x="值", y="主体", orientation="h",
        title=title
    )
    if is_percent:
        fig.update_traces(text=(agg["值"]*100).round(2).astype(str)+"%", textposition="outside")
    return fig

def pie_from_filter(df_subject, filter_value, group_letter, value_letter, title, percent_mode=False, total_denominator=None):
    sub = df_subject[col(df_subject,"A") == filter_value].copy()
    grouped = sub.groupby(col(sub, group_letter)).apply(lambda x: pd.to_numeric(col(x, value_letter), errors="coerce").fillna(0).sum())
    grouped = grouped.replace([np.inf, -np.inf], np.nan).fillna(0)
    dfp = grouped.reset_index()
    dfp.columns = ["主体", "值"]
    if percent_mode:
        if total_denominator and total_denominator>0:
            dfp["值"] = (dfp["值"] / total_denominator) * 100.0
        else:
            dfp["值"] = 0.0
    fig = px.pie(dfp, values="值", names="主体", title=title, hole=0.3)
    return fig

def account_rank_table(df_account):
    status = col(df_account, "H").astype(str).str.strip().str.lower()
    mask = status.isin(["active", "need to pay"])
    sub = df_account[mask].copy()
    sub["排序值"] = pd.to_numeric(col(sub, "J"), errors="coerce")
    sub = sub.sort_values("排序值", ascending=False).head(10)
    out = pd.DataFrame({
        "C列": col(sub, "C").astype(str),
        "D列": col(sub, "D").astype(str),
        "E列": col(sub, "E").astype(str),
        "J列": col(sub, "J")
    })
    return out

def stacked_bar_internal(df_subject):
    sub = df_subject[col(df_subject,"A") == "内部户表"].copy()
    for letter in ["B","C","D","E","F","G","H"]:
        if letter != "B":
            sub[letter] = pd.to_numeric(col(sub, letter), errors="coerce").fillna(0)
        else:
            sub[letter] = col(sub, letter).astype(str)
    agg = sub.groupby("B").agg({
        "C":"sum","D":"sum","E":"sum","F":"sum","G":"sum","H":"sum"
    }).reset_index().rename(columns={"B":"主体"})
    total_states = agg["D"]+agg["E"]+agg["F"]+agg["G"]+agg["H"]
    overflow = (total_states - agg["C"]).clip(lower=0)
    agg["G"] = (agg["G"] - overflow).clip(lower=0)
    fig = go.Figure()
    categories = ["D","E","F","G","H"]
    names_map = {"D":"已交付","E":"E列","F":"未绑卡","G":"G列","H":"死户"}
    for cat in categories:
        fig.add_trace(go.Bar(x=agg["主体"], y=agg[cat], name=names_map.get(cat, cat)))
    fig.update_layout(barmode="stack", title="内部户主体状态分布（校正后）",
                      xaxis_title="主体", yaxis_title="数量")
    return fig

# ----------------------------
# Streamlit App
# ----------------------------
st.set_page_config(page_title="主体分析可视化", layout="wide")
st.title("主体分析可视化看板")

# 🔹 固定读取仓库里的 Excel 文件
EXCEL_FILE = "主体分析.xlsx"
df_subject = pd.read_excel(EXCEL_FILE, sheet_name="主体分析表", header=0)
df_account = pd.read_excel(EXCEL_FILE, sheet_name="账户表", header=0)

# KPI
ts_str, total_accounts, delivered_accounts, unbound_cards, avg_dead_rate = build_kpis(df_subject)

st.subheader("数据概览")
kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
kpi1.metric("数据时间", ts_str)
kpi2.metric("总账户数", f"{int(total_accounts):,}")
kpi3.metric("已交付账户数", f"{int(delivered_accounts):,}")
kpi4.metric("未绑卡账户数", f"{int(unbound_cards):,}")
kpi5.metric("平均死户率", f"{avg_dead_rate:.2f}%")

st.subheader("榜单 Top10")
col1, col2 = st.columns(2)
with col1:
    st.plotly_chart(topn_chart(df_subject,"B","Q","主体综合评分榜单（Q）", how="mean", n=10), use_container_width=True)
    st.plotly_chart(topn_chart(df_subject,"B","L","主体平均消耗榜单（L）", how="mean", n=10), use_container_width=True)
with col2:
    st.plotly_chart(topn_chart(df_subject,"B","P","主体好户榜单（P）", how="mean", n=10, is_percent=True), use_container_width=True)
    st.plotly_chart(topn_chart(df_subject,"B","I","主体死户率榜单（I）", how="mean", n=10, is_percent=True), use_container_width=True)

st.subheader("占比分析")
total_accounts_all = safe_sum(col(df_subject, "C"))
col3, col4 = st.columns(2)
with col3:
    st.plotly_chart(pie_from_filter(df_subject,"内部户表","B","C","内部户主体账户数占比"), use_container_width=True)
    st.plotly_chart(pie_from_filter(df_subject,"内部户表","B","O","内部户主体消耗占比"), use_container_width=True)
    st.plotly_chart(pie_from_filter(df_subject,"内部户表","B","H","内部户主体死户数占比", percent_mode=True, total_denominator=total_accounts_all), use_container_width=True)
with col4:
    st.plotly_chart(pie_from_filter(df_subject,"外部户表","B","C","外部户主体账户数占比"), use_container_width=True)
    st.plotly_chart(pie_from_filter(df_subject,"外部户表","B","O","外部户主体消耗占比"), use_container_width=True)
    st.plotly_chart(pie_from_filter(df_subject,"外部户表","B","H","外部户主体死户数占比", percent_mode=True, total_denominator=total_accounts_all), use_container_width=True)

st.subheader("账户排行榜")
rank_df = account_rank_table(df_account)
st.dataframe(rank_df)

st.subheader("内部户主体状态分布")
st.plotly_chart(stacked_bar_internal(df_subject), use_container_width=True)
