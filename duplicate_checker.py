import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Stocktake 檢查工具", layout="wide")
st.title("📦 Stocktake 檢查工具")

# ---------- 側邊欄：永遠只放 StockTake2 ----------
with st.sidebar:
    f_stock = st.file_uploader("📤 上傳 StockTake2.xlsx", type=["xlsx"])
if not f_stock:
    st.warning("⚠️ 請先在側邊欄上傳 StockTake2.xlsx")
    st.stop()

df = pd.read_excel(f_stock).drop_duplicates()
df.columns = df.columns.str.strip()

# ---------- Tab ----------
tab1, tab2 = st.tabs(["🔍 重複 SerialNo / CMDB", "🚩 JG Outstanding"])

# ===== Tab1 =====
with tab1:
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

# ===== Tab2 =====
with tab2:
    st.header("🚩 JG Outstanding")
    f_shops = st.file_uploader("📤 上傳 All3shops.xlsx", type=["xlsx"], key="shops_in_tab2")
    if not f_shops:
        st.warning("⚠️ 請在此處上傳 All3shops.xlsx")
        st.stop()

    shops = pd.read_excel(f_shops)
    shops.columns = shops.columns.str.strip()

    # 比對邏輯
    JG = shops[shops["From JG"] == "Y"]
    JG["Serial No"] = JG["Serial No"].astype(str).str.strip()
    df["SerialNo"] = df["SerialNo"].astype(str).str.strip()

    outstanding = JG[
        (JG["Stock Take"] == "Y") &
        (~JG["Serial No"].isin(df["SerialNo"]))
    ]

    st.dataframe(outstanding)
    if not outstanding.empty:
        buf2 = BytesIO()
        outstanding.to_excel(buf2, index=False, engine="openpyxl")
        st.download_button("📥 下載 JG_outstanding.xlsx", buf2.getvalue(), "JG_outstanding.xlsx")