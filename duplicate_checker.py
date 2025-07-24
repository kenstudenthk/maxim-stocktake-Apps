"""
重複檢查 + JG Outstanding
只須上傳：
  1) Stocktake2.xlsx      （盤點結果）
  2) All3shops.xlsx       （庫存主檔）
即可下載兩份結果
"""
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Stocktake 檢查工具", layout="wide")
st.title("📦 重複檢查 + JG Outstanding")

# ---------- 側邊欄：固定先上傳 Stocktake2 ----------
with st.sidebar:
    f_stock = st.file_uploader("📤 上傳 Stocktake2.xlsx", type=["xlsx"])

# -------------- 主頁 Tab --------------
tab1, tab2 = st.tabs(["🔍 重複 SerialNo / CMDB", "🚩 JG Outstanding"])

# 通用：只要 Stocktake2 沒傳就停止
if f_stock is None:
    st.warning("⚠️ 請先在側邊欄上傳 Stocktake2.xlsx")
    st.stop()

df_stock = pd.read_excel(f_stock).drop_duplicates()
df_stock.columns = df_stock.columns.str.strip()

# ===== Tab1：重複檢查 =====
with tab1:
    dup_serial = (
        df_stock[df_stock["SerialNo"].astype(str).str.contains(r"\d", na=False)]
        .dropna(subset=["SerialNo"])
        .groupby("SerialNo")
        .filter(lambda g: len(g) > 1)
        .assign(Duplicate_Type="SerialNo")
    )
    dup_cmdb = (
        df_stock[df_stock["CMDB"] != "Device Not Found"]
        .dropna(subset=["CMDB"])
        .groupby("CMDB")
        .filter(lambda g: len(g) > 1)
        .assign(Duplicate_Type="CMDB")
    )
    duplicate_df = pd.concat([dup_serial, dup_cmdb]).drop_duplicates()

    st.subheader("重複結果")
    st.dataframe(duplicate_df)
    if not duplicate_df.empty:
        buf = BytesIO()
        duplicate_df.to_excel(buf, index=False, engine="openpyxl")
        st.download_button("📥 下載 duplicate_item.xlsx", buf.getvalue(), "duplicate_item.xlsx")

# ===== Tab2：JG Outstanding =====
with tab2:
    f_shops = st.file_uploader("📤 上傳 All3shops.xlsx", type=["xlsx"], key="shops_tab2")
    if f_shops is None:
        st.stop()

    shops = pd.read_excel(f_shops)
    shops.columns = shops.columns.str.strip()

    jg = shops[shops["From JG"] == "Y"].copy()
    jg["Serial No"] = jg["Serial No"].astype(str).str.strip()
    df_stock["SerialNo"] = df_stock["SerialNo"].astype(str).str.strip()

    # 1) 找出 JG 未盤點項目
    jg["found_in_stocktake"] = jg["Serial No"].isin(df_stock["SerialNo"])
    outstanding = jg[(~jg["found_in_stocktake"]) & (jg["Stock Take"] == "Y")]

    # 2) 新增：只保留 Verified / New Record 的盤點紀錄
    records_for_jg = df_stock[
        df_stock["Stock.Take.Status"].isin(["Verified", "New Record"])
    ]

    st.subheader("1️⃣ JG Outstanding（在 Stocktake2 找不到序號）")
    st.dataframe(outstanding)

    st.subheader("2️⃣ JG 盤點紀錄（狀態 = Verified 或 New Record）")
    st.dataframe(records_for_jg)

    # 下載
    if not outstanding.empty:
        buf = BytesIO()
        outstanding.to_excel(buf, index=False, engine="openpyxl")
        st.download_button("📥 JG_outstanding.xlsx", buf.getvalue(), "JG_outstanding.xlsx")

    if not records_for_jg.empty:
        buf2 = BytesIO()
        records_for_jg.to_excel(buf2, index=False, engine="openpyxl")
        st.download_button("📥 JG_records_Verified_or_NewRecord.xlsx", buf2.getvalue(), "JG_records.xlsx")