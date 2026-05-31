import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# 基本設定
# ==========================================

st.set_page_config(layout="wide")

st.title("📊 BB頁面 - Google Sheets 歷史紀錄")

# ==========================================
# Google Sheets 連線（手動版）
# ==========================================

SHEET_ID = "你的GoogleSheetID"  # 👈 改這裡

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ⚠️ Cloud 必須用公開或 service account（這裡用簡化公開讀法）

gc = gspread.service_account()

sh = gc.open_by_key(SHEET_ID)

# ==========================================
# 讀取資料
# ==========================================

try:
    tasks_ws = sh.worksheet("Tasks")
    history_ws = sh.worksheet("Data")

    tasks_df = pd.DataFrame(tasks_ws.get_all_records())
    history_df = pd.DataFrame(history_ws.get_all_records())

except Exception as e:
    st.error("❌ 無法讀取 Google Sheet，請檢查權限或 sheet 名稱")
    st.stop()

# ==========================================
# 安全處理
# ==========================================

if tasks_df.empty:
    tasks_df = pd.DataFrame()

if history_df.empty:
    history_df = pd.DataFrame()

# ==========================================
# UI
# ==========================================

tab1, tab2, tab3 = st.tabs(["📌 任務", "📦 歷史", "📊 統計"])

# ==========================================
# 📌 任務
# ==========================================

with tab1:

    st.subheader("目前任務")

    st.dataframe(tasks_df, use_container_width=True)

    search = st.text_input("搜尋任務")

    if search and not tasks_df.empty:
        filtered = tasks_df[
            tasks_df.astype(str).apply(
                lambda x: x.str.contains(search, case=False, na=False)
            ).any(axis=1)
        ]

        st.write(f"搜尋結果：{len(filtered)} 筆")
        st.dataframe(filtered, use_container_width=True)

# ==========================================
# 📦 歷史
# ==========================================

with tab2:

    st.subheader("歷史紀錄")

    st.dataframe(history_df, use_container_width=True)

    search2 = st.text_input("搜尋歷史")

    if search2 and not history_df.empty:
        filtered2 = history_df[
            history_df.astype(str).apply(
                lambda x: x.str.contains(search2, case=False, na=False)
            ).any(axis=1)
        ]

        st.write(f"搜尋結果：{len(filtered2)} 筆")
        st.dataframe(filtered2, use_container_width=True)

# ==========================================
# 📊 統計
# ==========================================

with tab3:

    st.subheader("統計")

    if not tasks_df.empty:

        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric("To Do", len(tasks_df[tasks_df["status"] == "To Do"]))

        with col2:
            st.metric("Executing", len(tasks_df[tasks_df["status"] == "In Executing"]))

        with col3:
            st.metric("Done", len(tasks_df[tasks_df["status"] == "Done"]))

    st.metric("歷史總數", len(history_df))
