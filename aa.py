import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime

# ==========================================
# 基本設定
# ==========================================

st.set_page_config(layout="wide")

st.title("📌 震唯機械企業版：雲端 Trello 管理系統")
st.caption("edit by 林溫城")

# ==========================================
# Google Sheets
# ==========================================

conn = st.connection("gsheets", type=GSheetsConnection)

df = conn.read(worksheet="Tasks", ttl=0)

# 初始化欄位
if df.empty:

    df = pd.DataFrame(columns=[
        "department",
        "customer",
        "title",
        "status",
        "owner",
        "created_time",
        "due_time",
        "updated_time"
    ])

# ==========================================
# 新增任務區
# ==========================================

st.write("## 📝 建立新任務")

with st.form("task_form", clear_on_submit=True):

    # 第一列
    c1, c2 = st.columns([1, 2])

    with c1:
        new_department = st.selectbox(
            "🏢 部門",
            ["業務", "生產", "驗收", "售服"]
        )

    with c2:
        new_customer = st.text_input(
            "🏭 客戶資訊",
            placeholder="請輸入客戶名稱"
        )

    # 第二列
    c3, c4, c5 = st.columns([2, 1, 1])

    with c3:
        new_title = st.text_input(
            "📌 任務名稱",
            placeholder="請輸入任務..."
        )

    with c4:
        new_status = st.selectbox(
            "📂 狀態",
            ["To Do", "In Executing", "Done"]
        )

    with c5:
        new_owner = st.text_input(
            "👤 負責人",
            placeholder="輸入負責人"
        )

    # 預計完成時間
    new_due_time = st.datetime_input(
        "⏰ 預計完成時間",
        value=datetime.now()
    )

    submit_btn = st.form_submit_button("✅ 建立任務")

# ==========================================
# 建立任務
# ==========================================

if submit_btn and new_title and new_owner:

    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    new_data = {

        "department": new_department,
        "customer": new_customer,
        "title": new_title,
        "status": new_status,
        "owner": new_owner,

        # 自動建立時間
        "created_time": current_time,

        # 預計完成時間
        "due_time": new_due_time.strftime("%Y-%m-%d %H:%M:%S"),

        # 初始更新時間
        "updated_time": current_time
    }

    new_row = pd.DataFrame([new_data])

    updated_df = pd.concat([df, new_row], ignore_index=True)

    conn.update(
        worksheet="Tasks",
        data=updated_df
    )

    st.success("✅ 任務建立成功")

    st.rerun()

st.write("---")

# ==========================================
# Trello 看板
# ==========================================

st.write("## 📊 任務管理看板")

col1, col2, col3 = st.columns(3)

# ==========================================
# 卡片函式
# ==========================================

def render_tasks(task_df, column_name):

    if not task_df.empty:

        for idx, row in task_df.iterrows():

            with st.container(border=True):

                st.write(f"### {row['title']}")

                st.caption(f"🏢 部門：{row.get('department', '')}")
                st.caption(f"🏭 客戶：{row.get('customer', '')}")
                st.caption(f"👤 負責人：{row['owner']}")

                st.caption(f"🕒 建立時間：{row.get('created_time', '')}")

                st.caption(f"⏰ 預計完成：{row.get('due_time', '')}")

                st.caption(f"🔄 最後更新：{row.get('updated_time', '')}")

                # 更新狀態
                new_status = st.selectbox(
                    "更新狀態",
                    ["To Do", "In Executing", "Done"],
                    index=["To Do", "In Executing", "Done"].index(row["status"]),
                    key=f"status_{idx}"
                )

                # 更新按鈕
                if st.button("💾 更新", key=f"update_{idx}"):

                    current_update_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    # 更新狀態
                    df.loc[idx, "status"] = new_status

                    # 更新最後時間
                    df.loc[idx, "updated_time"] = current_update_time

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
        st.info(f"目前沒有 {column_name} 任務")

# ==========================================
# To Do
# ==========================================

with col1:

    st.markdown(
        "## <span style='color:red'>🔴 To Do</span>",
        unsafe_allow_html=True
    )

    todo_df = df[df["status"] == "To Do"]

    render_tasks(todo_df, "待辦")

# ==========================================
# In Executing
# ==========================================

with col2:

    st.markdown(
        "## <span style='color:orange'>🟠 In Executing</span>",
        unsafe_allow_html=True
    )

    executing_df = df[df["status"] == "In Executing"]

    render_tasks(executing_df, "執行中")

# ==========================================
# Done
# ==========================================

with col3:

    st.markdown(
        "## <span style='color:green'>🟢 Done</span>",
        unsafe_allow_html=True
    )

    done_df = df[df["status"] == "Done"]

    render_tasks(done_df, "已完成")

                # ==========================================
                # 封存按鈕（只在 Done 顯示）
                # ==========================================

                if row["status"] == "Done":

                    if st.button("📦 封存到 Data", key=f"archive_{idx}"):

                        # 讀取 Data 工作頁
                        data_df = conn.read(
                            worksheet="Data",
                            ttl=0
                        )

                        # 若 Data 工作頁為空
                        if data_df.empty:

                            data_df = pd.DataFrame(columns=df.columns)

                        # 加入封存資料
                        archived_df = pd.concat(
                            [data_df, pd.DataFrame([row])],
                            ignore_index=True
                        )

                        # 更新 Data 工作頁
                        conn.update(
                            worksheet="Data",
                            data=archived_df
                        )

                        # 從 Tasks 移除
                        updated_df = df.drop(idx).reset_index(drop=True)

                        # 更新 Tasks 工作頁
                        conn.update(
                            worksheet="Tasks",
                            data=updated_df
                        )

                        st.success("✅ 已封存到 Data 工作頁")

                        st.rerun()
