import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection

# ==========================================
# 基本設定
# ==========================================

st.set_page_config(layout="wide")

st.title("📊 BB頁面 - 任務歷史紀錄中心")
st.caption("Google Sheets 任務與歷史資料檢視")

# ==========================================
# Google Sheets 連線
# ==========================================

conn = st.connection("gsheets", type=GSheetsConnection)

# 讀取目前任務
tasks_df = conn.read(worksheet="Tasks", ttl=0)

# 讀取歷史資料（封存）
history_df = conn.read(worksheet="Data", ttl=0)

# ==========================================
# 安全初始化
# ==========================================

if tasks_df is None:
    tasks_df = pd.DataFrame()

if history_df is None:
    history_df = pd.DataFrame()

# ==========================================
# Tab 分頁
# ==========================================

tab1, tab2, tab3 = st.tabs(["📌 目前任務", "📦 歷史封存", "📊 統計"])

# ==========================================
# 📌 目前任務
# ==========================================

with tab1:

    st.subheader("目前進行中的任務")

    if tasks_df.empty:
        st.info("目前沒有任務")
    else:

        st.dataframe(
            tasks_df.sort_values(by="created_time", ascending=False),
            use_container_width=True
        )

        # 搜尋
        search = st.text_input("🔍 搜尋任務（標題 / 客戶 / 負責人）")

        if search:
            filtered = tasks_df[
                tasks_df.astype(str).apply(
                    lambda x: x.str.contains(search, case=False, na=False)
                ).any(axis=1)
            ]

            st.write(f"搜尋結果：{len(filtered)} 筆")

            st.dataframe(filtered, use_container_width=True)

# ==========================================
# 📦 歷史封存
# ==========================================

with tab2:

    st.subheader("已封存任務")

    if history_df.empty:
        st.info("目前沒有歷史紀錄")
    else:

        st.dataframe(
            history_df.sort_values(by="updated_time", ascending=False),
            use_container_width=True
        )

        # 搜尋歷史
        search2 = st.text_input("🔍 搜尋歷史紀錄")

        if search2:
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

    st.subheader("任務統計")

    if not tasks_df.empty:

        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric("📌 To Do", len(tasks_df[tasks_df["status"] == "To Do"]))

        with col2:
            st.metric("🟠 In Executing", len(tasks_df[tasks_df["status"] == "In Executing"]))

        with col3:
            st.metric("🟢 Done", len(tasks_df[tasks_df["status"] == "Done"]))

    else:
        st.info("沒有任務資料可統計")

    st.write("---")

    if not history_df.empty:
        st.metric("📦 歷史總數", len(history_df))
