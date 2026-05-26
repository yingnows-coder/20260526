import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from streamlit_sortables import sort_items

# ==========================================
# 基本設定
# ==========================================

st.set_page_config(layout="wide")

st.title("📌 真正拖拉式 Trello 看板")
st.caption("Drag & Drop + Google Sheets 雲端同步")

# ==========================================
# Google Sheets 連線
# ==========================================

conn = st.connection("gsheets", type=GSheetsConnection)

df = conn.read(worksheet="Tasks", ttl=0)

# 空表防呆
if df.empty:
    df = pd.DataFrame(columns=["title", "status", "owner"])

# ==========================================
# 新增任務
# ==========================================

with st.form("new_task", clear_on_submit=True):

    col1, col2, col3 = st.columns([2,1,1])

    with col1:
        title = st.text_input("任務名稱")

    with col2:
        owner = st.text_input("負責人")

    with col3:
        status = st.selectbox(
            "狀態",
            ["To Do", "In Executing", "Done"]
        )

    submit = st.form_submit_button("新增任務")

if submit and title and owner:

    new_row = pd.DataFrame([{
        "title": title,
        "status": status,
        "owner": owner
    }])

    df = pd.concat([df, new_row], ignore_index=True)

    conn.update(
        worksheet="Tasks",
        data=df
    )

    st.success("✅ 任務已新增")

    st.rerun()

st.divider()

# ==========================================
# 建立拖拉資料
# ==========================================

todo_tasks = df[df["status"] == "To Do"]["title"].tolist()

doing_tasks = df[df["status"] == "In Executing"]["title"].tolist()

done_tasks = df[df["status"] == "Done"]["title"].tolist()

# ==========================================
# Trello Drag UI
# ==========================================

board = [

    {
        "header": "🔴 To Do",
        "items": todo_tasks
    },

    {
        "header": "🟠 In Executing",
        "items": doing_tasks
    },

    {
        "header": "🟢 Done",
        "items": done_tasks
    }
]

# ==========================================
# 顯示拖拉元件
# ==========================================

sorted_board = sort_items(
    board,
    multi_containers=True
)

# ==========================================
# 同步回 Google Sheets
# ==========================================

status_map = {
    "🔴 To Do": "To Do",
    "🟠 In Executing": "In Executing",
    "🟢 Done": "Done"
}

updated_rows = []

for column in sorted_board:

    status = status_map[column["header"]]

    for task_title in column["items"]:

        # 找原始資料
        task_data = df[df["title"] == task_title].iloc[0]

        updated_rows.append({
            "title": task_title,
            "status": status,
            "owner": task_data["owner"]
        })

updated_df = pd.DataFrame(updated_rows)

# ==========================================
# 更新資料
# ==========================================

if not updated_df.equals(df):

    conn.update(
        worksheet="Tasks",
        data=updated_df
    )

    st.success("✅ 看板已同步更新")

# ==========================================
# 額外功能：刪除任務
# ==========================================

st.divider()

st.write("### 🗑️ 刪除任務")

task_to_delete = st.selectbox(
    "選擇任務",
    df["title"].tolist()
)

if st.button("刪除"):

    df = df[df["title"] != task_to_delete]

    conn.update(
        worksheet="Tasks",
        data=df
    )

    st.warning("⚠️ 任務已刪除")

    st.rerun()
