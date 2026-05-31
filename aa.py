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

top1, top2 = st.columns([8,1])

with top2:
    st.page_link(
        "pages/bb.py",
        label="📊 BB頁面",
        icon="🚀"
    )

# ==========================================
# Google Sheets
# ==========================================

conn = st.connection("gsheets", type=GSheetsConnection)
df = conn.read(worksheet="Tasks", ttl=0)

# 初始化欄位
if df.empty:
    df = pd.DataFrame(columns=[
        "id",
        "department",
        "customer",
        "title",
        "status",
        "owner",
        "created_time",
        "due_time",
        "updated_time"
    ])
else:
    # 確保 id 存在
    if "id" not in df.columns:
        df["id"] = [str(uuid.uuid4()) for _ in range(len(df))]

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

# ==========================================
# 建立任務
# ==========================================

if submit_btn and new_title and new_owner:

    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    new_data = {
        "id": str(uuid.uuid4()),
        "department": new_department,
        "customer": new_customer,
        "title": new_title,
        "status": new_status,
        "owner": new_owner,
        "created_time": current_time,
        "due_time": new_due_time.strftime("%Y-%m-%d %H:%M:%S"),
        "updated_time": current_time
    }

    df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)

    conn.update(worksheet="Tasks", data=df)

    st.success("✅ 任務建立成功")
    st.rerun()

st.write("---")

# ==========================================
# 看板
# ==========================================

st.write("## 📊 任務管理看板")

col1, col2, col3 = st.columns(3)

# ==========================================
# 卡片渲染
# ==========================================

def render_tasks(task_df, column_name):

    if task_df.empty:
        st.info(f"目前沒有 {column_name} 任務")
        return

    for _, row in task_df.iterrows():

        real_idx = df[df["id"] == row["id"]].index[0]

        with st.container(border=True):

            st.write(f"### {row['title']}")

            st.caption(f"🏢 部門：{row.get('department','')}")
            st.caption(f"🏭 客戶：{row.get('customer','')}")
            st.caption(f"👤 負責人：{row.get('owner','')}")
            st.caption(f"🕒 建立時間：{row.get('created_time','')}")
            st.caption(f"⏰ 預計完成：{row.get('due_time','')}")
            st.caption(f"🔄 更新時間：{row.get('updated_time','')}")

            # 狀態更新
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

            if row["status"] == "Done":

                if st.button("📦 封存", key=f"archive_{row['id']}"):

                    data_df = conn.read(worksheet="Data", ttl=0)

                    if data_df.empty:
                        data_df = pd.DataFrame(columns=df.columns)

                    row_data = df.loc[real_idx]

                    data_df = pd.concat([data_df, pd.DataFrame([row_data])], ignore_index=True)

                    conn.update(worksheet="Data", data=data_df)

                    df.drop(real_idx, inplace=True)
                    df.reset_index(drop=True, inplace=True)

                    conn.update(worksheet="Tasks", data=df)

                    st.success("✅ 已封存")
                    st.rerun()

            if st.button("🗑️ 刪除", key=f"delete_{row['id']}"):

                df.drop(real_idx, inplace=True)
                df.reset_index(drop=True, inplace=True)

                conn.update(worksheet="Tasks", data=df)

                st.warning("⚠️ 已刪除")
                st.rerun()

# ==========================================
# 分欄顯示
# ==========================================

with col1:
    st.markdown("## 🔴 To Do")
    render_tasks(df[df["status"] == "To Do"], "待辦")

with col2:
    st.markdown("## 🟠 In Executing")
    render_tasks(df[df["status"] == "In Executing"], "執行中")

with col3:
    st.markdown("## 🟢 Done")
    render_tasks(df[df["status"] == "Done"], "已完成")
