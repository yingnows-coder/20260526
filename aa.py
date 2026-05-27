import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime

# ==========================================
# 基本設定
# ==========================================

st.set_page_config(layout="wide")

st.title("📌 震唯機械任務指派看板：GitHub 雲端同步 Trello 看板")
st.caption("授權標註：edit by 林溫城 | 2026052701A版")

# ==========================================
# 連接 Google Sheets
# ==========================================

conn = st.connection("gsheets", type=GSheetsConnection)

df = conn.read(worksheet="Tasks", ttl=0)

# 避免空資料錯誤
if df.empty:
    df = pd.DataFrame(columns=[
        "department",
        "customer",
        "datetime",
        "title",
        "status",
        "owner"
    ])

# ==========================================
# 區塊一：新增任務
# ==========================================

st.write("### 📝 指派新任務")

with st.form("task_input_form", clear_on_submit=True):

    c_dept, c_customer = st.columns([1, 2])

    # 部門選擇
    with c_dept:
        new_department = st.selectbox(
            "🏢 部門",
            ["業務", "生產", "驗收", "售服"]
        )

    # 客戶資訊
    with c_customer:
        new_customer = st.text_input(
            "🏭 客戶資訊",
            placeholder="輸入客戶名稱..."
        )

    c_title, c_status, c_owner = st.columns([2, 1, 1])

    # 任務名稱
    with c_title:
        new_title = st.text_input(
            "📌 任務名稱",
            placeholder="輸入任務名稱..."
        )

    # 狀態
    with c_status:
        new_status = st.selectbox(
            "📂 狀態",
            ["To Do", "In Executing", "Done"]
        )

    # 負責人
    with c_owner:
        new_owner = st.text_input(
            "👤 負責人",
            placeholder="誰來負責..."
        )

    # 日期時間
    new_datetime = st.datetime_input(
        "🕒 時間",
        value=datetime.now()
    )

    submit_btn = st.form_submit_button("✅ 確認指派並同步雲端")

# ==========================================
# 新增資料
# ==========================================

if submit_btn and new_title and new_owner:

    new_data = {
        "department": new_department,
        "customer": new_customer,
        "datetime": new_datetime.strftime("%Y-%m-%d %H:%M"),
        "title": new_title,
        "status": new_status,
        "owner": new_owner
    }

    new_row = pd.DataFrame([new_data])

    updated_df = pd.concat([df, new_row], ignore_index=True)

    conn.update(
        worksheet="Tasks",
        data=updated_df
    )

    st.success("✅ 任務已成功同步到 Google Sheets！")

    st.rerun()

st.write("---")

# ==========================================
# 區塊二：Trello 看板
# ==========================================

st.write("### 📊 看板動態狀態監控")

trello_col1, trello_col2, trello_col3 = st.columns(3)

# ==========================================
# 共用卡片函式
# ==========================================

def render_tasks(task_df, column_name):

    if not task_df.empty:

        for idx, row in task_df.iterrows():

            with st.container(border=True):

                st.write(f"### {row['title']}")

                # 顯示部門
                st.caption(f"🏢 部門：{row.get('department', '')}")

                # 顯示客戶
                st.caption(f"🏭 客戶：{row.get('customer', '')}")

                # 顯示負責人
                st.caption(f"👤 負責人：{row['owner']}")

                # 顯示時間
                st.caption(f"🕒 時間：{row.get('datetime', '')}")

                # 修改狀態
                new_status = st.selectbox(
                    "更新狀態",
                    ["To Do", "In Executing", "Done"],
                    index=["To Do", "In Executing", "Done"].index(row["status"]),
                    key=f"status_{idx}"
                )

                # 更新按鈕
                if st.button("💾 更新", key=f"update_{idx}"):

                    df.loc[idx, "status"] = new_status

                    conn.update(
                        worksheet="Tasks",
                        data=df
                    )

                    st.success("✅ 狀態已更新")

                    st.rerun()

                # 刪除按鈕
                if st.button("🗑️ 刪除任務", key=f"delete_{idx}"):

                    updated_df = df.drop(idx).reset_index(drop=True)

                    conn.update(
                        worksheet="Tasks",
                        data=updated_df
                    )

                    st.warning("⚠️ 任務已刪除")

                    st.rerun()

    else:
        st.info(f"暫無 {column_name} 任務")

# ==========================================
# 第一欄：To Do
# ==========================================

with trello_col1:

    st.markdown(
        "## <span style='color:red'>🔴 To Do</span>",
        unsafe_allow_html=True
    )

    todo_list = df[df["status"] == "To Do"]

    render_tasks(todo_list, "待辦")

# ==========================================
# 第二欄：In Executing
# ==========================================

with trello_col2:

    st.markdown(
        "## <span style='color:orange'>🟠 In Executing</span>",
        unsafe_allow_html=True
    )

    ip_list = df[df["status"] == "In Executing"]

    render_tasks(ip_list, "執行中")

# ==========================================
# 第三欄：Done
# ==========================================

with trello_col3:

    st.markdown(
        "## <span style='color:green'>🟢 Done</span>",
        unsafe_allow_html=True
    )

    done_list = df[df["status"] == "Done"]

    render_tasks(done_list, "已完成")
