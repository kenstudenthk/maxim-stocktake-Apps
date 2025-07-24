# ----------------------------------------------------------
# 1. 安裝依賴 (第一次運行時取消註解)
# ----------------------------------------------------------
# pip install pandas openpyxl streamlit

# ----------------------------------------------------------
# 2. 匯入套件
# ----------------------------------------------------------
import pandas as pd
import numpy as np
from pathlib import Path
import streamlit as st
from datetime import datetime, date

# ----------------------------------------------------------
# 3. Streamlit 側邊欄：選擇檔案
# ----------------------------------------------------------
st.sidebar.header("📁 請選擇檔案")
def file_selector(label: str, ext: str):
    f = st.sidebar.file_uploader(label, type=ext)
    return f

f_stock        = file_selector("Stocktake2.xlsx", ["xlsx"])
f_schedule     = file_selector("Full Schedule with Contacts.xlsx", ["xlsx"])
f_all3shops    = file_selector("All3shops.xlsx", ["xlsx"])
f_bbi          = file_selector("BBI Stocktake.xlsx", ["xlsx"])

# 日期範圍
start_date = date(2025, 7, 22)
end_date   = date(2025, 7, 22)

# ----------------------------------------------------------
# 4. 讀檔函式 (自動快取)
# ----------------------------------------------------------
@st.cache_data
def load_data():
    if not all([f_stock, f_schedule, f_all3shops, f_bbi]):
        st.stop()
    
    df       = pd.read_excel(f_stock)        .drop_duplicates()
    schedule = pd.read_excel(f_schedule, sheet_name="Schedule")
    shops    = pd.read_excel(f_all3shops)
    bbi      = pd.read_excel(f_bbi)
    
    return df, schedule, shops, bbi

df, schedule, shops, bbi = load_data()

# ----------------------------------------------------------
# 5. 資料前處理
# ----------------------------------------------------------
# 5.1 先把 schedule 的日期欄清理好
schedule = schedule.loc[~schedule["Date"].isin(["TBC", np.nan])]
schedule["Date"] = pd.to_datetime(schedule["Date"], errors="coerce").dt.date
schedule["ShopCode"] = schedule["ShopCode"].astype(str).str.zfill(5)

# 5.2 篩選日期
mask = (
    (schedule["Date"] >= start_date) &
    (schedule["Date"] <= end_date) &
    (schedule["Available"] == "Y")
)
filtered_schedule = schedule.loc[mask].dropna(how="all")

# 5.3 JG 項目
JG = shops[shops["From JG"] == "Y"]

# ----------------------------------------------------------
# 6. 過濾資料框
# ----------------------------------------------------------
df1 = df[df["Shop.Name"].isin(filtered_schedule["Shop.Name"])]
df2 = JG[JG["Shop Code"].isin(filtered_schedule["ShopCode"])]

# ----------------------------------------------------------
# 7. Check quantity (Verified & Device Not Found)
# ----------------------------------------------------------
check_quantity = (
    df1[df1["Stock.Take.Status"].isin(["Verified", "Device Not Found"])]
    .groupby("Shop.Name")
    .agg(
        max_qty   = ("TotalQty", "max"),
        row_count = ("TotalQty", "count")
    )
    .assign(check=lambda x: x["max_qty"] == x["row_count"])
)
st.subheader("📊 Check Quantity")
st.dataframe(check_quantity)

# ----------------------------------------------------------
# 8. Duplicate SerialNo & CMDB
# ----------------------------------------------------------
# 8.1 SerialNo duplicate
dup_serial = (
    df[df["SerialNo"].str.contains(r"\d", na=False)]
    .dropna(subset=["SerialNo"])
    .groupby("SerialNo")
    .filter(lambda g: len(g) > 1)
    .assign(Duplicate_Type="SerialNo")
)

# 8.2 CMDB duplicate
dup_cmdb = (
    df[df["CMDB"] != "Device Not Found"]
    .dropna(subset=["CMDB"])
    .groupby("CMDB")
    .filter(lambda g: len(g) > 1)
    .assign(Duplicate_Type="CMDB")
)

duplicate_entries = pd.concat([dup_serial, dup_cmdb]).drop_duplicates()

result2 = (
    duplicate_entries
    .merge(filtered_schedule, on="Shop.Name", how="left")
    [["Duplicate_Type", "Date", "Stock.Take.Status", "SerialNo", "CMDB",
      "Shop.Name", "Product.Type.(Eng)", "Product.Type.(Chi)", "Brand",
      "Asset.Name", "Asset.Item.ID", "IP.Address", "MX.No."]]
    .sort_values("Date")
)

st.subheader("📋 Duplicate SerialNo / CMDB")
st.dataframe(result2)

# ----------------------------------------------------------
# 9. Outstanding JG items
# ----------------------------------------------------------
df2["Serial No"]  = df2["Serial No"].astype(str).str.strip()
df1["SerialNo"]   = df1["SerialNo"].astype(str).str.strip()

df2["in_df1"] = df2["Serial No"].isin(df1["SerialNo"])
outstanding_items = df2[(df2["in_df1"] == False) & (df2["Stock Take"] == "Y")]

# 9.1 讀 roster
roster = filtered_schedule[["Main", "Assistant", "ShopCode"]].copy()
roster.columns = ["Main", "Assistant", "Shop Code"]

final = (
    outstanding_items
    .merge(roster, on="Shop Code", how="left")
    .assign(Date=lambda x: pd.to_datetime(x["Date"], errors="coerce").dt.date)
    .assign(New_Installed=lambda x: x["CreatedAt"] > x["Date"])
    .loc[lambda x: ~x["New_Installed"]]
    .drop(columns=["JG Date"], errors="ignore")
    .sort_values("Date")
)

st.subheader("🚩 Outstanding JG Items")
st.dataframe(final)

# ----------------------------------------------------------
# 10. BBI 缺失店鋪
# ----------------------------------------------------------
missing_shop = set(filtered_schedule["Shop.Name"]) - set(bbi["ShopName"])
st.subheader("❌ BBI 未包含店鋪")
st.write(missing_shop)

# ----------------------------------------------------------
# 11. 匯出 Excel (供下載)
# ----------------------------------------------------------
def to_excel(df_in):
    from io import BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_in.to_excel(writer, index=False, sheet_name="Sheet1")
    output.seek(0)
    return output.getvalue()

st.download_button(
    label="📤 下載 daily_records.xlsx",
    data=to_excel(df1[df1["Stock.Take.Status"].isin(["Verified", "New Record"]) &
                      ~df1["SerialNo"].isin(result2["SerialNo"])]),
    file_name="Shop_04134.xlsx"
)