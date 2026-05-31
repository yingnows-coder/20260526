import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import uuid

# ==========================================
# 基本設定
# ==========================================

st.set_page_config(layout="wide")

st.title("📌 企業版：雲端 Trello 管理系統")
st.caption("edit by 林溫城")

# ==========================================
# BB頁面導航
# ==========================================

top1, top2 = st.columns([8,1])

with top2:
    if st.button("📊 BB頁面"):
        st.switch_page("pages/bb.py")

# ==========================================
# Google Sheets
# ==========================================

conn = st.connection("gsheets", type=GSheetsConnection)

df = conn.read(worksheet="Tasks", ttl=0)
history_df = conn.read(worksheet="Data", ttl=0)

# ==========================================
# 初始化
# ==========================================

if df.empty:
    df = pd.DataFrame(columns=[
        "id","department","customer","title","status",
        "owner","created_time","due_time","updated_time"
    ])
else:
    if "id" not in df.columns:
        df["id"] = [str(uuid.uuid4()) for _ in range(len(df))]

if history_df is None:
    history_df = pd.DataFrame()

# ==========================================
# 新增任務
# ==========================================

st.write("## 📝 建立新任務")

with st.form("task_form", clear_on_submit=True):

    c1, c2 = st.columns([1, 2])

    with c1:
        new_department = st.selectbox("🏢 部門", ["業務", "生產", "驗收", "售服"])

    with c2:
        new_customer = st.text_input("🏭 客戶資訊")

    c3, c4, c5 = st.columns([2, 1, 1])

    with c3:
        new_title = st.text_input("📌 任務名稱")

    with c4:
        new_status = st.selectbox("📂 狀態", ["To Do", "In Executing", "Done"])

    with c5:
        new_owner = st.text_input("👤 負責人")

    new_due_time = st.datetime_input("⏰ 預計完成時間", value=datetime.now())

    submit_btn = st.form_submit_button("✅ 建立任務")

if submit_btn and new_title and new_owner:

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    new_data = {
        "id": str(uuid.uuid4()),
        "department": new_department,
        "customer": new_customer,
        "title": new_title,
        "status": new_status,
        "owner": new_owner,
        "created_time": now,
        "due_time": new_due_time.strftime("%Y-%m-%d %H:%M:%S"),
        "updated_time": now
    }

    df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)

    conn.update(worksheet="Tasks", data=df)

    st.success("✅ 任務建立成功")
    st.rerun()

st.write("---")

# ==========================================
# 📊 主功能區（目前 + 歷史）
# ==========================================

tab1, tab2 = st.tabs(["📌 目前任務", "📦 歷史紀錄"])

# ==========================================
# 卡片函式
# ==========================================

def render_tasks(task_df):

    if task_df.empty:
        st.info("沒有資料")
        return

    for _, row in task_df.iterrows():

        real_idx = df[df["id"] == row["id"]].index[0]

        with st.container(border=True):

            st.write(f"### {row.get('title','')}")

            st.caption(f"🏢 部門：{row.get('department','')}")
            st.caption(f"🏭 客戶：{row.get('customer','')}")
            st.caption(f"👤 負責人：{row.get('owner','')}")
            st.caption(f"🕒 建立時間：{row.get('created_time','')}")
            st.caption(f"⏰ 預計完成：{row.get('due_time','')}")
            st.caption(f"🔄 更新時間：{row.get('updated_time','')}")

            new_status = st.selectbox(
                "更新狀態",
                ["To Do", "In Executing", "Done"],
                index=["To Do", "In Executing", "Done"].index(row["status"]),
                key=f"status_{row['id']}"
            )

            if st.button("💾 更新", key=f"update_{row['id']}"):

                df.loc[real_idx, "status"] = new_status
                df.loc[real_idx, "updated_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                conn.update(worksheet="Tasks", data=df)

                st.success("✅ 已更新")
                st.rerun()

            # ==========================================
            # 封存 → 移到歷史 Data
            # ==========================================

            if row["status"] == "Done":

                if st.button("📦 封存到歷史", key=f"archive_{row['id']}"):

                    row_data = df.loc[real_idx]

                    history_df_local = conn.read(worksheet="Data", ttl=0)

                    if history_df_local.empty:
                        history_df_local = pd.DataFrame(columns=df.columns)

                    history_df_local = pd.concat(
                        [history_df_local, pd.DataFrame([row_data])],
                        ignore_index=True
                    )

                    conn.update(worksheet="Data", data=history_df_local)

                    df.drop(real_idx, inplace=True)
                    df.reset_index(drop=True, inplace=True)

                    conn.update(worksheet="Tasks", data=df)

                    st.success("✅ 已移動到歷史紀錄")
                    st.rerun()

            if st.button("🗑️ 刪除", key=f"delete_{row['id']}"):

                df.drop(real_idx, inplace=True)
                df.reset_index(drop=True, inplace=True)

                conn.update(worksheet="Tasks", data=df)

                st.warning("⚠️ 已刪除")
                st.rerun()

# ==========================================
# 📌 目前任務
# ==========================================

with tab1:

    st.subheader("進行中任務")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("## 🔴 To Do")
        render_tasks(df[df["status"] == "To Do"])

    with col2:
        st.markdown("## 🟠 In Executing")
        render_tasks(df[df["status"] == "In Executing"])

    with col3:
        st.markdown("## 🟢 Done")
        render_tasks(df[df["status"] == "Done"])

# ==========================================
# 📦 歷史紀錄
# ==========================================

with tab2:

    st.subheader("📦 任務歷史紀錄")

    if history_df.empty:
        st.info("目前沒有歷史資料")
    else:
        st.dataframe(history_df, use_container_width=True)

        search = st.text_input("🔍 搜尋歷史紀錄")

        if search:
            filtered = history_df[
                history_df.astype(str).apply(
                    lambda x: x.str.contains(search, case=False, na=False)
                ).any(axis=1)
            ]

            st.write(f"搜尋結果：{len(filtered)} 筆")
            st.dataframe(filtered, use_container_width=True)
