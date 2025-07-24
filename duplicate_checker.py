# duplicate_checker.py
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="重複序號 / CMDB 檢查器", layout="wide")
st.title("📋 重複 SerialNo & CMDB 檢查器")

uploaded = st.file_uploader("📁 選擇 Stocktake2.xlsx", type=["xlsx"])
if not uploaded:
    st.stop()

# 讀檔 & 清理欄位名稱
df = pd.read_excel(uploaded).drop_duplicates()
df.columns = df.columns.str.strip()

# 1. 重複 SerialNo ------------------------------------------------------------
dup_serial = (
    df[df["SerialNo"].astype(str).str.contains(r"\d", na=False)]
    .dropna(subset=["SerialNo"])
    .groupby("SerialNo")
    .filter(lambda g: len(g) > 1)
    .assign(Duplicate_Type="SerialNo")
)

# 2. 重複 CMDB ---------------------------------------------------------------
dup_cmdb = (
    df[df["CMDB"] != "Device Not Found"]
    .dropna(subset=["CMDB"])
    .groupby("CMDB")
    .filter(lambda g: len(g) > 1)
    .assign(Duplicate_Type="CMDB")
)

# 3. 合併結果
duplicate_df = (
    pd.concat([dup_serial, dup_cmdb])
    .drop_duplicates()
    .sort_values(["Duplicate_Type", "SerialNo"])
)

# 4. 展示
st.subheader("🔍 重複結果")
st.dataframe(duplicate_df, use_container_width=True)

# 5. 下載 Excel
if not duplicate_df.empty:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        duplicate_df.to_excel(writer, index=False, sheet_name="Duplicates")
    buffer.seek(0)
    st.download_button(
        label="📥 下載結果 Excel",
        data=buffer,
        file_name="duplicate items.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.success("✅ 目前沒有重複的 SerialNo 或 CMDB")

# 6. Link for remind user upload the file to Sharepoint
st.info(
    "📤 **下一步請手動上傳到 SharePoint**  \n"
    "點擊下方連結，進入資料夾後直接 **覆蓋** 舊的 `duplicate items.xlsx` 即可：  \n"
    "[🔗 前往資料夾（SharePoint）](https://pccw0.sharepoint.com/:f:/r/sites/BonniesTeam/Shared%20Documents/General/Maxim%27s%20stock%20take/Do_not_open?csf=1&web=1&e=arYEyY)"
)