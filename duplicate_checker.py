import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Stocktake 檢查工具", layout="wide")
st.title("📦 Stocktake 檢查工具")

# ---------- 側邊欄容器（先留空，稍後動態加入） ----------
sidebar_holder = st.sidebar.empty()

# ---------- Tab 分頁 ----------
tab1, tab2 = st.tabs(["🔍 重複 SerialNo / CMDB", "🚩 JG Outstanding"])

# ===== Tab 1：重複檢查 =====
with tab1:
    # 只在 Tab1 開啟時把上傳元件掛到側邊欄
    with st.sidebar:
        f_stock = st.file_uploader("📁 上傳 Stocktake2.xlsx", type=["xlsx"])

    if f_stock is None:
        st.warning("請先在側邊欄上傳 Stocktake2.xlsx")
        st.stop()

    df = pd.read_excel(f_stock).drop_duplicates()
    df.columns = df.columns.str.strip()

    dup_serial = (
        df[df["SerialNo"].astype(str).str.contains(r"\d", na=False)]
        .dropna(subset=["SerialNo"])
        .groupby("SerialNo")
        .filter(lambda g: len(g) > 1)
        .assign(Duplicate_Type="SerialNo")
    )
    dup_cmdb = (
        df[df["CMDB"] != "Device Not Found"]
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

    st.info("📤 請上傳到 SharePoint：[🔗 前往資料夾](https://pccw0.sharepoint.com/:f:/r/sites/BonniesTeam/Shared%20Documents/General/Maxim%27s%20stock%20take/Do_not_open?csf=1&web=1&e=arYEyY)")

# ===== Tab 2：JG Outstanding =====
with tab2:
    # 只在 Tab2 開啟時把上傳元件掛到側邊欄
    with st.sidebar:
        f_shops = st.file_uploader("📁 上傳 All3shops.xlsx", type=["xlsx"])

    if f_shops is None:
        st.warning("請先在側邊欄上傳 All3shops.xlsx")
        st.stop()

    shops = pd.read_excel(f_shops)
    shops.columns = shops.columns.str.strip()

    # 讀 Stocktake2.xlsx（兩個 Tab 共用同一個檔案，因此再讀一次）
    if "f_stock" not in locals():
        f_stock = st.file_uploader("📁 重新上傳 Stocktake2.xlsx", type=["xlsx"], key="tab2_stock")
        if f_stock is None:
            st.stop()
        df = pd.read_excel(f_stock).drop_duplicates()
        df.columns = df.columns.str.strip()

    # 比對邏輯
    JG = shops[shops["From JG"] == "Y"]
    JG["Serial No"] = JG["Serial No"].astype(str).str.strip()
    df["SerialNo"] = df["SerialNo"].astype(str).str.strip()

    outstanding = JG[
        (JG["Stock Take"] == "Y") &
        (~JG["Serial No"].isin(df["SerialNo"]))
    ]

    st.subheader("JG 未盤點項目")
    st.dataframe(outstanding)

    if not outstanding.empty:
        buf2 = BytesIO()
        outstanding.to_excel(buf2, index=False, engine="openpyxl")
        st.download_button("📥 下載 JG_outstanding.xlsx", buf2.getvalue(), "JG_outstanding.xlsx")

    st.info("📤 請上傳到 SharePoint：[🔗 前往資料夾](https://pccw0.sharepoint.com/:f:/r/sites/BonniesTeam/Shared%20Documents/General/Maxim%27s%20stock%20take/Do_not_open?csf=1&web=1&e=arYEyY)")