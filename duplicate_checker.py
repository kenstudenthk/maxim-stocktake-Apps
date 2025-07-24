import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Stocktake 檢查工具", layout="wide")
st.title("📦 Stocktake 檢查工具")

# 用 session_state 記住目前在哪個 Tab
if "active_tab" not in st.session_state:
    st.session_state.active_tab = "tab1"   # 預設第一頁

# 建立 Tab，並在切換時同步 session_state
tab1, tab2 = st.tabs(["🔍 重複 SerialNo / CMDB", "🚩 JG Outstanding"])

# --------------- 動態側邊欄 ---------------
with st.sidebar:
    if st.session_state.active_tab == "tab1":
        f_stock = st.file_uploader("📤 上傳 Stocktake2.xlsx", type=["xlsx"], key="stock_tab1")
        f_shops = None   # 讓 Tab2 的元件消失
    else:
        f_shops = st.file_uploader("📤 上傳 All3shops.xlsx", type=["xlsx"], key="shops_tab2")
        f_stock = None   # 讓 Tab1 的元件消失

# --------------- Tab 1 ---------------
with tab1:
    # 記住目前 Tab
    st.session_state.active_tab = "tab1"
    if not f_stock:
        st.warning("請先在側邊欄上傳 Stocktake2.xlsx")
        st.stop()

    df = pd.read_excel(f_stock).drop_duplicates()
    df.columns = df.columns.str.strip()

    # ===== 重複檢查邏輯 =====
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

    st.info("📤 上傳到 SharePoint：[🔗 資料夾](https://pccw0.sharepoint.com/:f:/r/sites/BonniesTeam/Shared%20Documents/General/Maxim%27s%20stock%20take/Do_not_open?csf=1&web=1&e=arYEyY)")

# --------------- Tab 2 ---------------
with tab2:
    st.session_state.active_tab = "tab2"
    if not f_shops:
        st.warning("請先在側邊欄上傳 All3shops.xlsx")
        st.stop()

    shops = pd.read_excel(f_shops)
    shops.columns = shops.columns.str.strip()

    # 若還沒上傳 Stocktake2.xlsx，在 Tab2 再要求一次
    if "f_stock_tab2" not in st.session_state:
        f_stock_tab2 = st.file_uploader("📤 重新上傳 Stocktake2.xlsx", type=["xlsx"], key="f_stock_tab2")
        if not f_stock_tab2:
            st.stop()
        df = pd.read_excel(f_stock_tab2).drop_duplicates()
        df.columns = df.columns.str.strip()
        st.session_state.f_stock_tab2 = f_stock_tab2
    else:
        df = pd.read_excel(st.session_state.f_stock_tab2).drop_duplicates()

    # ===== JG Outstanding 邏輯 =====
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

    st.info("📤 上傳到 SharePoint：[🔗 資料夾](https://pccw0.sharepoint.com/:f:/r/sites/BonniesTeam/Shared%20Documents/General/Maxim%27s%20stock%20take/Do_not_open?csf=1&web=1&e=arYEyY)")