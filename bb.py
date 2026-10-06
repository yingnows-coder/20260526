import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime, date
import uuid

# ==========================================
# 基本設定
# ==========================================

st.set_page_config(
    page_title="企業版：雲端 Trello 管理系統",
    page_icon="📌",
    layout="wide"
)

st.title("📌 企業版：雲端 Trello 管理系統")
st.caption("edit by 林溫城")

# ==========================================
# Google Sheets
# ==========================================

conn = st.connection(
    "gsheets",
    type=GSheetsConnection
)

# ==========================================
# 工作表欄位
# ==========================================

TASK_COLUMNS = [
    "id",
    "department",
    "customer",
    "title",
    "status",
    "owner",
    "created_time",
    "due_time",
    "updated_time"
]

STATUS_LIST = [
    "To Do",
    "In Executing",
    "Done"
]

DEPARTMENT_LIST = [
    "業務",
    "生產",
    "驗收",
    "售服"
]

# ==========================================
# 讀取 Google Sheets
# ==========================================

try:
    df = conn.read(
        worksheet="Tasks",
        ttl=0
    )
except Exception as e:
    st.error(f"❌ 無法讀取 Tasks 工作表：{e}")
    df = pd.DataFrame(columns=TASK_COLUMNS)

try:
    history_df = conn.read(
        worksheet="Data",
        ttl=0
    )
except Exception:
    history_df = pd.DataFrame(columns=TASK_COLUMNS)


# ==========================================
# 資料初始化
# ==========================================

if df is None or df.empty:

    df = pd.DataFrame(columns=TASK_COLUMNS)

else:

    # 補齊缺少欄位
    for col in TASK_COLUMNS:

        if col not in df.columns:

            if col == "id":
                df[col] = [
                    str(uuid.uuid4())
                    for _ in range(len(df))
                ]
            else:
                df[col] = ""

    df = df[TASK_COLUMNS]


# ==========================================
# 初始化歷史資料
# ==========================================

if history_df is None or history_df.empty:

    history_df = pd.DataFrame(
        columns=TASK_COLUMNS
    )

else:

    for col in TASK_COLUMNS:

        if col not in history_df.columns:
            history_df[col] = ""

    history_df = history_df[TASK_COLUMNS]


# ==========================================
# BB 頁面
# ==========================================

top1, top2 = st.columns([8, 1])

with top2:

    if st.button(
        "📊 BB頁面",
        use_container_width=True
    ):

        st.switch_page("pages/bb.py")


# ==========================================
# 📊 任務統計
# ==========================================

st.subheader("📊 任務總覽")

total_count = len(df)

todo_count = len(
    df[df["status"] == "To Do"]
)

executing_count = len(
    df[df["status"] == "In Executing"]
)

done_count = len(
    df[df["status"] == "Done"]
)

# ==========================================
# 逾期計算
# ==========================================

today = date.today()

overdue_count = 0

if not df.empty:

    due_dates = pd.to_datetime(
        df["due_time"],
        errors="coerce"
    )

    overdue_mask = (
        (due_dates.dt.date < today)
        & (df["status"] != "Done")
    )

    overdue_count = overdue_mask.sum()


# ==========================================
# 統計卡片
# ==========================================

m1, m2, m3, m4, m5 = st.columns(5)

with m1:
    st.metric(
        "📋 全部任務",
        total_count
    )

with m2:
    st.metric(
        "🔴 To Do",
        todo_count
    )

with m3:
    st.metric(
        "🟠 執行中",
        executing_count
    )

with m4:
    st.metric(
        "🟢 已完成",
        done_count
    )

with m5:
    st.metric(
        "⚠️ 已逾期",
        overdue_count
    )


st.divider()


# ==========================================
# 📝 建立新任務
# ==========================================

st.subheader("📝 建立新任務")

with st.form(
    "task_form",
    clear_on_submit=True
):

    c1, c2 = st.columns([1, 2])

    with c1:

        new_department = st.selectbox(
            "🏢 部門",
            DEPARTMENT_LIST
        )

    with c2:

        new_customer = st.text_input(
            "🏭 客戶資訊"
        )

    c3, c4, c5 = st.columns(
        [2, 1, 1]
    )

    with c3:

        new_title = st.text_input(
            "📌 任務名稱"
        )

    with c4:

        new_status = st.selectbox(
            "📂 狀態",
            STATUS_LIST
        )

    with c5:

        new_owner = st.text_input(
            "👤 負責人"
        )

    new_due_time = st.datetime_input(
        "⏰ 預計完成時間",
        value=datetime.now()
    )

    submit_btn = st.form_submit_button(
        "✅ 建立任務",
        use_container_width=True
    )


# ==========================================
# 新增任務
# ==========================================

if submit_btn:

    if not new_title.strip():

        st.error(
            "⚠️ 請輸入任務名稱"
        )

    elif not new_owner.strip():

        st.error(
            "⚠️ 請輸入負責人"
        )

    else:

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        new_data = {

            "id": str(uuid.uuid4()),

            "department":
                new_department,

            "customer":
                new_customer,

            "title":
                new_title,

            "status":
                new_status,

            "owner":
                new_owner,

            "created_time":
                now,

            "due_time":
                new_due_time.strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),

            "updated_time":
                now
        }

        df = pd.concat(
            [
                df,
                pd.DataFrame([new_data])
            ],
            ignore_index=True
        )

        conn.update(
            worksheet="Tasks",
            data=df
        )

        st.success(
            "✅ 任務建立成功"
        )

        st.rerun()


st.divider()


# ==========================================
# 主功能區
# ==========================================

tab1, tab2 = st.tabs(
    [
        "📌 目前任務",
        "📦 歷史紀錄"
    ]
)


# ==========================================
# 任務卡片
# ==========================================

def render_tasks(task_df):

    if task_df.empty:

        st.info(
            "目前沒有符合條件的任務"
        )

        return


    for _, row in task_df.iterrows():

        task_id = row["id"]

        matched = df[
            df["id"] == task_id
        ]

        if matched.empty:
            continue

        real_idx = matched.index[0]


        # --------------------------------------
        # 日期判斷
        # --------------------------------------

        due_value = row.get(
            "due_time",
            ""
        )

        due_date = pd.to_datetime(
            due_value,
            errors="coerce"
        )

        is_overdue = False

        if pd.notna(due_date):

            is_overdue = (
                due_date.date() < today
                and row["status"] != "Done"
            )


        # --------------------------------------
        # 卡片
        # --------------------------------------

        with st.container(
            border=True
        ):

            # ----------------------------------
            # 標題
            # ----------------------------------

            if is_overdue:

                st.markdown(
                    f"### ⚠️ {row.get('title', '')}"
                )

                st.error(
                    "🚨 此任務已逾期"
                )

            else:

                st.markdown(
                    f"### 📌 {row.get('title', '')}"
                )


            # ----------------------------------
            # 基本資訊
            # ----------------------------------

            info1, info2 = st.columns(2)

            with info1:

                st.caption(
                    f"🏢 部門：{row.get('department', '')}"
                )

                st.caption(
                    f"🏭 客戶：{row.get('customer', '')}"
                )

                st.caption(
                    f"👤 負責人：{row.get('owner', '')}"
                )

            with info2:

                st.caption(
                    f"🕒 建立時間：{row.get('created_time', '')}"
                )

                st.caption(
                    f"⏰ 預計完成：{row.get('due_time', '')}"
                )

                st.caption(
                    f"🔄 更新時間：{row.get('updated_time', '')}"
                )


            # ----------------------------------
            # 狀態
            # ----------------------------------

            new_status = st.selectbox(

                "📂 更新狀態",

                STATUS_LIST,

                index=(
                    STATUS_LIST.index(
                        row["status"]
                    )
                    if row["status"] in STATUS_LIST
                    else 0
                ),

                key=f"status_{task_id}"
            )


            # ----------------------------------
            # 編輯區
            # ----------------------------------

            with st.expander(
                "✏️ 編輯任務"
            ):

                edit_department = st.selectbox(

                    "🏢 部門",

                    DEPARTMENT_LIST,

                    index=(
                        DEPARTMENT_LIST.index(
                            row["department"]
                        )
                        if row["department"]
                        in DEPARTMENT_LIST
                        else 0
                    ),

                    key=f"edit_department_{task_id}"
                )


                edit_customer = st.text_input(

                    "🏭 客戶資訊",

                    value=str(
                        row.get(
                            "customer",
                            ""
                        )
                    ),

                    key=f"edit_customer_{task_id}"
                )


                edit_title = st.text_input(

                    "📌 任務名稱",

                    value=str(
                        row.get(
                            "title",
                            ""
                        )
                    ),

                    key=f"edit_title_{task_id}"
                )


                edit_owner = st.text_input(

                    "👤 負責人",

                    value=str(
                        row.get(
                            "owner",
                            ""
                        )
                    ),

                    key=f"edit_owner_{task_id}"
                )


                existing_due = pd.to_datetime(

                    row.get(
                        "due_time",
                        ""
                    ),

                    errors="coerce"
                )


                if pd.isna(existing_due):

                    existing_due = datetime.now()


                edit_due_time = st.datetime_input(

                    "⏰ 預計完成時間",

                    value=existing_due.to_pydatetime(),

                    key=f"edit_due_{task_id}"
                )


                if st.button(

                    "💾 儲存編輯",

                    key=f"save_edit_{task_id}",

                    use_container_width=True
                ):

                    now = datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )

                    df.loc[
                        real_idx,
                        "department"
                    ] = edit_department

                    df.loc[
                        real_idx,
                        "customer"
                    ] = edit_customer

                    df.loc[
                        real_idx,
                        "title"
                    ] = edit_title

                    df.loc[
                        real_idx,
                        "owner"
                    ] = edit_owner

                    df.loc[
                        real_idx,
                        "due_time"
                    ] = edit_due_time.strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )

                    df.loc[
                        real_idx,
                        "updated_time"
                    ] = now

                    conn.update(
                        worksheet="Tasks",
                        data=df
                    )

                    st.success(
                        "✅ 任務內容已更新"
                    )

                    st.rerun()


            # ----------------------------------
            # 按鈕
            # ----------------------------------

            b1, b2, b3 = st.columns(3)


            # 更新狀態
            with b1:

                if st.button(

                    "💾 更新",

                    key=f"update_{task_id}",

                    use_container_width=True
                ):

                    df.loc[
                        real_idx,
                        "status"
                    ] = new_status

                    df.loc[
                        real_idx,
                        "updated_time"
                    ] = datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )

                    conn.update(
                        worksheet="Tasks",
                        data=df
                    )

                    st.success(
                        "✅ 狀態已更新"
                    )

                    st.rerun()


            # ----------------------------------
            # 封存
            # ----------------------------------

            with b2:

                if row["status"] == "Done":

                    if st.button(

                        "📦 封存",

                        key=f"archive_{task_id}",

                        use_container_width=True
                    ):

                        row_data = df.loc[
                            real_idx
                        ].copy()


                        try:

                            history_df_local = conn.read(
                                worksheet="Data",
                                ttl=0
                            )

                        except Exception:

                            history_df_local = pd.DataFrame(
                                columns=TASK_COLUMNS
                            )


                        if (
                            history_df_local is None
                            or history_df_local.empty
                        ):

                            history_df_local = pd.DataFrame(
                                columns=TASK_COLUMNS
                            )


                        for col in TASK_COLUMNS:

                            if col not in history_df_local.columns:

                                history_df_local[col] = ""


                        history_df_local = (
                            history_df_local[TASK_COLUMNS]
                        )


                        history_df_local = pd.concat(
                            [
                                history_df_local,
                                pd.DataFrame(
                                    [row_data]
                                )
                            ],
                            ignore_index=True
                        )


                        # 更新歷史
                        conn.update(
                            worksheet="Data",
                            data=history_df_local
                        )


                        # 從目前任務刪除
                        df.drop(
                            real_idx,
                            inplace=True
                        )

                        df.reset_index(
                            drop=True,
                            inplace=True
                        )


                        conn.update(
                            worksheet="Tasks",
                            data=df
                        )


                        st.success(
                            "📦 已移動到歷史紀錄"
                        )

                        st.rerun()


            # ----------------------------------
            # 刪除
            # ----------------------------------

            with b3:

                if st.button(

                    "🗑️ 刪除",

                    key=f"delete_{task_id}",

                    use_container_width=True
                ):

                    df.drop(
                        real_idx,
                        inplace=True
                    )

                    df.reset_index(
                        drop=True,
                        inplace=True
                    )

                    conn.update(
                        worksheet="Tasks",
                        data=df
                    )

                    st.warning(
                        "🗑️ 任務已刪除"
                    )

                    st.rerun()


# ==========================================
# 📌 目前任務
# ==========================================

with tab1:

    st.subheader(
        "📌 進行中任務"
    )


    # ======================================
    # 🔍 快捷搜尋
    # ======================================

    st.markdown(
        "### 🔍 快捷搜尋與篩選"
    )

    search_task = st.text_input(

        "搜尋",

        placeholder=
        "輸入任務名稱、客戶、負責人、部門...",

        key="task_search"
    )


    f1, f2, f3, f4 = st.columns(4)


    # ======================================
    # 部門
    # ======================================

    with f1:

        department_options = sorted(
            df["department"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        selected_departments = st.multiselect(

            "🏢 部門",

            department_options,

            key="filter_department"
        )


    # ======================================
    # 負責人
    # ======================================

    with f2:

        owner_options = sorted(
            df["owner"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        selected_owners = st.multiselect(

            "👤 負責人",

            owner_options,

            key="filter_owner"
        )


    # ======================================
    # 客戶
    # ======================================

    with f3:

        customer_options = sorted(
            df["customer"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        selected_customers = st.multiselect(

            "🏭 客戶",

            customer_options,

            key="filter_customer"
        )


    # ======================================
    # 日期
    # ======================================

    with f4:

        date_filter = st.selectbox(

            "📅 日期",

            [
                "全部",
                "今天到期",
                "7天內到期",
                "已逾期"
            ],

            key="filter_date"
        )


    # ======================================
    # 搜尋資料
    # ======================================

    display_df = df.copy()


    # 關鍵字搜尋

    if search_task.strip():

        keyword = search_task.strip()

        search_mask = (
            display_df
            .astype(str)
            .apply(
                lambda x:
                x.str.contains(
                    keyword,
                    case=False,
                    na=False,
                    regex=False
                )
            )
            .any(axis=1)
        )

        display_df = display_df[
            search_mask
        ]


    # 部門

    if selected_departments:

        display_df = display_df[
            display_df["department"]
            .isin(selected_departments)
        ]


    # 負責人

    if selected_owners:

        display_df = display_df[
            display_df["owner"]
            .isin(selected_owners)
        ]


    # 客戶

    if selected_customers:

        display_df = display_df[
            display_df["customer"]
            .isin(selected_customers)
        ]


    # ======================================
    # 日期篩選
    # ======================================

    if date_filter != "全部":

        temp_due = pd.to_datetime(
            display_df["due_time"],
            errors="coerce"
        )


        if date_filter == "今天到期":

            display_df = display_df[
                temp_due.dt.date == today
            ]


        elif date_filter == "7天內到期":

            max_date = today + pd.Timedelta(
                days=7
            )

            display_df = display_df[
                (
                    temp_due.dt.date >= today
                )
                &
                (
                    temp_due.dt.date <= max_date
                )
                &
                (
                    display_df["status"]
                    != "Done"
                )
            ]


        elif date_filter == "已逾期":

            display_df = display_df[
                (
                    temp_due.dt.date < today
                )
                &
                (
                    display_df["status"]
                    != "Done"
                )
            ]


    # ======================================
    # 搜尋結果
    # ======================================

    st.caption(
        f"🔎 目前符合條件："
        f"{len(display_df)} 筆任務"
    )


    # ======================================
    # 三欄看板
    # ======================================

    col1, col2, col3 = st.columns(3)


    # ======================================
    # To Do
    # ======================================

    with col1:

        st.markdown(
            f"## 🔴 To Do "
            f"({len(display_df[display_df['status'] == 'To Do'])})"
        )

        render_tasks(
            display_df[
                display_df["status"]
                == "To Do"
            ]
        )


    # ======================================
    # In Executing
    # ======================================

    with col2:

        st.markdown(
            f"## 🟠 In Executing "
            f"({len(display_df[display_df['status'] == 'In Executing'])})"
        )

        render_tasks(
            display_df[
                display_df["status"]
                == "In Executing"
            ]
        )


    # ======================================
    # Done
    # ======================================

    with col3:

        st.markdown(
            f"## 🟢 Done "
            f"({len(display_df[display_df['status'] == 'Done'])})"
        )

        render_tasks(
            display_df[
                display_df["status"]
                == "Done"
            ]
        )


# ==========================================
# 📦 歷史紀錄
# ==========================================

with tab2:

    st.subheader(
        "📦 任務歷史紀錄"
    )


    # ======================================
    # 重新讀取最新歷史
    # ======================================

    try:

        history_df = conn.read(
            worksheet="Data",
            ttl=0
        )

    except Exception:

        history_df = pd.DataFrame(
            columns=TASK_COLUMNS
        )


    if (
        history_df is None
        or history_df.empty
    ):

        st.info(
            "目前沒有歷史資料"
        )

    else:

        # ==================================
        # 歷史搜尋
        # ==================================

        history_search = st.text_input(

            "🔍 搜尋歷史紀錄",

            placeholder=
            "搜尋任務、客戶、負責人、部門...",

            key="history_search"
        )


        history_display = history_df.copy()


        if history_search.strip():

            history_mask = (
                history_display
                .astype(str)
                .apply(
                    lambda x:
                    x.str.contains(
                        history_search.strip(),
                        case=False,
                        na=False,
                        regex=False
                    )
                )
                .any(axis=1)
            )

            history_display = (
                history_display[
                    history_mask
                ]
            )


        st.caption(
            f"📦 共 {len(history_display)} 筆歷史紀錄"
        )


        st.dataframe(

            history_display,

            use_container_width=True,

            hide_index=True
        )


        # ==================================
        # 歷史統計
        # ==================================

        if "department" in history_display.columns:

            st.markdown(
                "### 📊 歷史任務部門統計"
            )

            department_count = (
                history_display[
                    "department"
                ]
                .value_counts()
            )

            st.bar_chart(
                department_count
            )
