import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import plotly.express as px

# ==========================================
# 基本設定
# ==========================================

st.set_page_config(layout="wide")

st.title("📌 企業級 Trello 雲端管理系統")
st.caption("edit by 闕河正")

# ==========================================
# Google Sheets
# ==========================================

conn = st.connection("gsheets", type=GSheetsConnection)

df = conn.read(worksheet="Tasks", ttl=0)

# ==========================================
# 初始化
# ==========================================

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
# 搜尋功能
# ==========================================

st.sidebar.title("🔍 搜尋任務")

search_text = st.sidebar.text_input(
    "輸入關鍵字",
    placeholder="搜尋任務/客戶/負責人"
)

# 搜尋過濾
if search_text:

    df = df[
        df.astype(str)
        .apply(lambda row: row.str.contains(search_text, case=False).any(), axis=1)
    ]

# ==========================================
# KPI 統計
# ==========================================

st.write("## 📈 KPI 統計")

total_tasks = len(df)

done_tasks = len(df[df["status"] == "Done"])

todo_tasks = len(df[df["status"] == "To Do"])

executing_tasks = len(df[df["status"] == "In Executing"])

completion_rate = 0

if total_tasks > 0:
    completion_rate = round((done_tasks / total_tasks) * 100, 1)

k1, k2, k3, k4 = st.columns(4)

k1.metric("📌 總任務", total_tasks)
k2.metric("🟢 已完成", done_tasks)
k3.metric("🟠 執行中", executing_tasks)
k4.metric("✅ 完成率", f"{completion_rate}%")

# ==========================================
# KPI 圖表
# ==========================================

chart_df = pd.DataFrame({
    "狀態": ["To Do", "In Executing", "Done"],
    "數量": [todo_tasks, executing_tasks, done_tasks]
})

fig = px.pie(
    chart_df,
    names="狀態",
    values="數量",
    title="📊 任務狀態分布"
)

st.plotly_chart(fig, use_container_width=True)

st.write("---")

# ==========================================
# 新增任務
# ==========================================

st.write("## 📝 建立新任務")

with st.form("task_form", clear_on_submit=True):

    c1, c2 = st.columns([1, 2])

    with c1:

        new_department = st.selectbox(
            "🏢 部門",
            ["業務", "生產", "驗收", "售服"]
        )

    with c2:

        new_customer = st.text_input(
            "🏭 客戶資訊"
        )

    c3, c4, c5 = st.columns([2, 1, 1])

    with c3:

        new_title = st.text_input(
            "📌 任務名稱"
        )

    with c4:

        new_status = st.selectbox(
            "📂 狀態",
            ["To Do", "In Executing", "Done"]
        )

    with c5:

        new_owner = st.text_input(
            "👤 負責人"
        )

    new_due_time = st.datetime_input(
        "⏰ 預計完成時間",
        value=datetime.now()
    )

    submit_btn = st.form_submit_button("✅ 建立任務")

# ==========================================
# 建立資料
# ==========================================

if submit_btn and new_title and new_owner:

    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    new_data = {

        "department": new_department,
        "customer": new_customer,
        "title": new_title,
        "status": new_status,
        "owner": new_owner,
        "created_time": current_time,
        "due_time": new_due_time.strftime("%Y-%m-%d %H:%M:%S"),
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

            # ==================================
            # 逾期判斷
            # ==================================

            overdue = False

            try:

                due_time = datetime.strptime(
                    str(row.get("due_time", "")),
                    "%Y-%m-%d %H:%M:%S"
                )

                if due_time < datetime.now() and row["status"] != "Done":
                    overdue = True

            except:
                pass

            # ==================================
            # 卡片
            # ==================================

            with st.container(border=True):

                # 逾期紅字
                if overdue:
                    st.error("🚨 任務已逾期")

                st.write(f"### {row['title']}")

                st.caption(f"🏢 部門：{row.get('department', '')}")
                st.caption(f"🏭 客戶：{row.get('customer', '')}")
                st.caption(f"👤 負責人：{row['owner']}")

                st.caption(f"🕒 建立時間：{row.get('created_time', '')}")

                st.caption(f"⏰ 預計完成：{row.get('due_time', '')}")

                st.caption(f"🔄 最後更新：{row.get('updated_time', '')}")

                # 剩餘時間
                try:

                    remaining = due_time - datetime.now()

                    days = remaining.days

                    if days >= 0 and not overdue:
                        st.info(f"⏳ 剩餘 {days} 天")

                except:
                    pass

                # 狀態更新
                new_status = st.selectbox(
                    "更新狀態",
                    ["To Do", "In Executing", "Done"],
                    index=["To Do", "In Executing", "Done"].index(row["status"]),
                    key=f"status_{idx}"
                )

                # 更新按鈕
                if st.button("💾 更新", key=f"update_{idx}"):

                    current_update_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    df.loc[idx, "status"] = new_status

                    df.loc[idx, "updated_time"] = current_update_time

                    conn.update(
                        worksheet="Tasks",
                        data=df
                    )

                    st.success("✅ 狀態已更新")

                    st.rerun()

                # 刪除
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
