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
        st.warning("請在此處上傳 All3shops.xlsx")
        st.stop()

    shops = pd.read_excel(f_shops)
    shops.columns = shops.columns.str.strip()

    # ---- 這裡就是你要的「找出 JG 未盤點項目」 ----
    jg = shops[shops["From JG"] == "Y"].copy()
    jg["Serial No"] = jg["Serial No"].astype(str).str.strip()
    df_stock["SerialNo"] = df_stock["SerialNo"].astype(str).str.strip()

    # 找出在 Stocktake2 找不到的 JG 序號
    jg["found_in_stocktake"] = jg["Serial No"].isin(df_stock["SerialNo"])
    outstanding = jg[
        (~jg["found_in_stocktake"]) &   # 找不到
        (jg["Stock Take"] == "Y")       # 且標記要盤點
    ]

    st.subheader("JG 未盤點項目（在 Stocktake2 找不到）")
    st.dataframe(outstanding)
    if not outstanding.empty:
        buf2 = BytesIO()
        outstanding.to_excel(buf2, index=False, engine="openpyxl")
        st.download_button("📥 下載 JG_outstanding.xlsx", buf2.getvalue(), "JG_outstanding.xlsx")

# ---------- SharePoint 提示 ----------
st.info(
    "📤 請把兩份檔案上傳到 SharePoint："
    "[🔗 前往資料夾](https://pccw0.sharepoint.com/:f:/r/sites/BonniesTeam/Shared%20Documents/General/Maxim%27s%20stock%20take/Do_not_open?csf=1&web=1&e=arYEyY)"
)