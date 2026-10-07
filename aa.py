import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime, timedelta
import uuid

# ==========================================
# 基本設定
# ==========================================

st.set_page_config(
    page_title="企業版 RFQ 詢價追蹤系統",
    layout="wide"
)

st.title("📌 企業版：RFQ 詢價追蹤管理系統")
st.caption("edit by 林溫城")

# ==========================================
# BB頁面導航
# ==========================================

top1, top2 = st.columns([8, 1])

with top2:
    st.link_button(
        "📊 BB頁面",
        "https://aazzyybb.streamlit.app/"
    )

# ==========================================
# Google Sheets
# ==========================================

conn = st.connection(
    "gsheets",
    type=GSheetsConnection
)

df = conn.read(
    worksheet="Tasks",
    ttl=0
)

# ==========================================
# RFQ 欄位
# ==========================================

RFQ_COLUMNS = [
    "id",
    "RFQ_ID",
    "version",
    "department",
    "customer",
    "title",
    "status",
    "next_step",
    "owner",
    "created_time",
    "due_time",
    "updated_time",
    "evidence",
    "exception",
    "approved_quote",
    "quote_sent",
    "first_followup_due"
]

# ==========================================
# 初始化資料
# ==========================================

if df.empty:

    df = pd.DataFrame(columns=RFQ_COLUMNS)

else:

    # 補上舊資料缺少的欄位
    for col in RFQ_COLUMNS:

        if col not in df.columns:

            if col == "version":
                df[col] = "V1"

            elif col == "owner":
                df[col] = "業務承辦"

            elif col in ["approved_quote", "quote_sent"]:
                df[col] = False

            else:
                df[col] = ""

    # id不存在時補ID
    if "id" not in df.columns:
        df["id"] = [
            str(uuid.uuid4())
            for _ in range(len(df))
        ]

    # RFQ_ID不存在時補RFQ_ID
    if "RFQ_ID" not in df.columns:
        df["RFQ_ID"] = [
            f"RFQ-{datetime.now().strftime('%Y%m%d')}-{str(uuid.uuid4())[:4].upper()}"
            for _ in range(len(df))
        ]

# ==========================================
# 狀態設定
# ==========================================

RFQ_STATUS = [
    "新詢價",
    "待補件",
    "工程評估",
    "待核價",
    "已報價",
    "追蹤中",
    "結案"
]

# ==========================================
# 計算下一個工作天
# ==========================================

def next_workday(date_value):

    next_day = date_value + timedelta(days=1)

    while next_day.weekday() >= 5:
        next_day += timedelta(days=1)

    return next_day


# ==========================================
# 狀態規則
# ==========================================

STATUS_NEXT_STEP = {

    "新詢價":
        "確認詢價資料完整性",

    "待補件":
        "向客戶補問缺漏資料",

    "工程評估":
        "請工程確認規格與可製造性",

    "待核價":
        "準備並取得核準報價",

    "已報價":
        "確認客戶收到報價並追蹤",

    "追蹤中":
        "追蹤客戶回覆",

    "結案":
        "完成案件結案紀錄"
}


# ==========================================
# 驗證狀態
# ==========================================

def validate_status(row, new_status):

    # 已報價防呆
    if new_status == "已報價":

        approved = str(
            row.get("approved_quote", "")
        ).lower() in ["true", "1", "yes", "是"]

        sent = str(
            row.get("quote_sent", "")
        ).lower() in ["true", "1", "yes", "是"]

        if not approved or not sent:

            return False, (
                "❌ 不可標記「已報價」："
                "必須同時有「核準報價」與「實際寄送紀錄」。"
            )

    return True, ""


# ==========================================
# 狀態變更紀錄
# ==========================================

def save_status_log(
    rfq_id,
    old_status,
    new_status,
    evidence
):

    try:

        log_df = conn.read(
            worksheet="StatusLog",
            ttl=0
        )

    except Exception:

        log_df = pd.DataFrame()

    log_columns = [
        "log_id",
        "RFQ_ID",
        "change_time",
        "old_status",
        "new_status",
        "evidence"
    ]

    if log_df.empty:

        log_df = pd.DataFrame(
            columns=log_columns
        )

    log_data = {
        "log_id": str(uuid.uuid4()),
        "RFQ_ID": rfq_id,
        "change_time":
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
        "old_status": old_status,
        "new_status": new_status,
        "evidence": evidence
    }

    log_df = pd.concat(
        [
            log_df,
            pd.DataFrame([log_data])
        ],
        ignore_index=True
    )

    conn.update(
        worksheet="StatusLog",
        data=log_df
    )


# ==========================================
# 新增 RFQ
# ==========================================

st.write("## 📝 建立新詢價 RFQ")

with st.form(
    "rfq_form",
    clear_on_submit=True
):

    c1, c2, c3 = st.columns(
        [1, 2, 1]
    )

    with c1:

        new_department = st.selectbox(
            "🏢 部門",
            [
                "業務",
                "生產",
                "驗收",
                "售服"
            ]
        )

    with c2:

        new_customer = st.text_input(
            "🏭 客戶資訊"
        )

    with c3:

        new_version = st.text_input(
            "📑 版本",
            value="V1"
        )

    c4, c5 = st.columns(
        [2, 1]
    )

    with c4:

        new_title = st.text_input(
            "📌 詢價名稱"
        )

    with c5:

        new_owner = st.text_input(
            "👤 負責人",
            value="業務承辦"
        )

    new_evidence = st.text_area(
        "📎 證據／詢價來源",
        placeholder="例如：客戶 Email、圖面、詢價單、會議紀錄..."
    )

    new_exception = st.text_area(
        "⚠️ 異常／缺漏",
        placeholder="若沒有，請填「無」；資料不足請填「待確認」"
    )

    new_due_time = st.date_input(
        "⏰ 到期日",
        value=None
    )

    submit_btn = st.form_submit_button(
        "✅ 建立 RFQ"
    )


# ==========================================
# 建立 RFQ
# ==========================================

if submit_btn:

    if not new_title:

        st.error(
            "❌ 請輸入詢價名稱"
        )

    else:

        now_dt = datetime.now()

        now = now_dt.strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        rfq_id = (
            f"RFQ-"
            f"{now_dt.strftime('%Y%m%d')}-"
            f"{str(uuid.uuid4())[:4].upper()}"
        )

        # 首次補問：下一個工作天
        first_followup = next_workday(
            now_dt.date()
        )

        # 如果使用者沒有輸入到期日
        # 不自行指定
        if new_due_time is None:
            due_time = "待確認"
        else:
            due_time = new_due_time.strftime(
                "%Y-%m-%d"
            )

        owner = (
            new_owner.strip()
            if new_owner.strip()
            else "業務承辦"
        )

        new_data = {

            "id":
                str(uuid.uuid4()),

            "RFQ_ID":
                rfq_id,

            "version":
                new_version.strip()
                if new_version.strip()
                else "V1",

            "department":
                new_department,

            "customer":
                new_customer,

            "title":
                new_title,

            "status":
                "新詢價",

            "next_step":
                STATUS_NEXT_STEP["新詢價"],

            "owner":
                owner,

            "created_time":
                now,

            "due_time":
                due_time,

            "updated_time":
                now,

            "evidence":
                new_evidence
                if new_evidence.strip()
                else "待確認",

            "exception":
                new_exception
                if new_exception.strip()
                else "待確認",

            "approved_quote":
                False,

            "quote_sent":
                False,

            "first_followup_due":
                first_followup.strftime(
                    "%Y-%m-%d"
                )
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

        # 建立初始狀態紀錄
        save_status_log(
            rfq_id,
            "無",
            "新詢價",
            new_evidence
            if new_evidence.strip()
            else "待確認"
        )

        st.success(
            f"✅ RFQ 建立成功：{rfq_id}"
        )

        st.rerun()


# ==========================================
# 快捷搜尋
# ==========================================

st.write("---")

st.write("## 🔎 RFQ 快捷搜尋")

search_keyword = st.text_input(
    "輸入 RFQ、客戶、詢價名稱、負責人、證據、異常",
    placeholder="例如：RFQ-20261007、台積電、業務承辦..."
)

if search_keyword:

    search_mask = (
        df.astype(str)
        .apply(
            lambda col:
            col.str.contains(
                search_keyword,
                case=False,
                na=False
            )
        )
        .any(axis=1)
    )

    display_df = df[
        search_mask
    ]

else:

    display_df = df


# ==========================================
# RFQ 統計
# ==========================================

st.write("## 📊 RFQ 狀態統計")

stat_cols = st.columns(
    len(RFQ_STATUS)
)

for col, status in zip(
    stat_cols,
    RFQ_STATUS
):

    count = len(
        df[df["status"] == status]
    )

    with col:

        st.metric(
            status,
            count
        )


# ==========================================
# RFQ 追蹤表
# ==========================================

st.write("## 📋 RFQ 追蹤表")

tracking_columns = [
    "RFQ_ID",
    "version",
    "customer",
    "title",
    "status",
    "next_step",
    "owner",
    "due_time",
    "updated_time",
    "evidence",
    "exception"
]

if display_df.empty:

    st.info("目前沒有符合條件的 RFQ")

else:

    st.dataframe(
        display_df[
            tracking_columns
        ],
        use_container_width=True,
        hide_index=True
    )


# ==========================================
# 看板
# ==========================================

st.write("---")
st.write("## 📊 RFQ 管理看板")

board_cols = st.columns(
    len(RFQ_STATUS)
)


# ==========================================
# 卡片函式
# ==========================================

def render_tasks(
    task_df,
    column_name
):

    if task_df.empty:

        st.info(
            f"目前沒有 {column_name}"
        )

        return

    for _, row in task_df.iterrows():

        matches = df[
            df["id"] == row["id"]
        ].index

        if len(matches) == 0:
            continue

        real_idx = matches[0]

        with st.container(
            border=True
        ):

            st.markdown(
                f"### 📌 {row['title']}"
            )

            st.caption(
                f"🆔 RFQ：{row['RFQ_ID']}"
            )

            st.caption(
                f"📑 版本：{row['version']}"
            )

            st.caption(
                f"🏢 部門：{row['department']}"
            )

            st.caption(
                f"🏭 客戶：{row['customer']}"
            )

            st.caption(
                f"👤 負責人：{row['owner']}"
            )

            st.caption(
                f"📅 到期日：{row['due_time']}"
            )

            st.caption(
                f"🔄 最後更新：{row['updated_time']}"
            )

            st.caption(
                f"➡️ 下一步：{row['next_step']}"
            )

            st.caption(
                f"📎 證據：{row['evidence']}"
            )

            st.caption(
                f"⚠️ 異常：{row['exception']}"
            )

            # ==================================
            # 首次補問
            # ==================================

            if row["status"] in [
                "新詢價",
                "待補件"
            ]:

                st.info(
                    "📅 首次補問期限："
                    f"{row['first_followup_due']}"
                )

            # ==================================
            # 報價條件
            # ==================================

            if row["status"] in [
                "待核價",
                "已報價",
                "追蹤中"
            ]:

                q1, q2 = st.columns(2)

                with q1:

                    approved = st.checkbox(
                        "✅ 核準報價",
                        value=(
                            str(
                                row.get(
                                    "approved_quote",
                                    False
                                )
                            ).lower()
                            in [
                                "true",
                                "1",
                                "yes",
                                "是"
                            ]
                        ),
                        key=f"approved_{row['id']}"
                    )

                with q2:

                    sent = st.checkbox(
                        "📤 實際寄送",
                        value=(
                            str(
                                row.get(
                                    "quote_sent",
                                    False
                                )
                            ).lower()
                            in [
                                "true",
                                "1",
                                "yes",
                                "是"
                            ]
                        ),
                        key=f"sent_{row['id']}"
                    )

            else:

                approved = row.get(
                    "approved_quote",
                    False
                )

                sent = row.get(
                    "quote_sent",
                    False
                )

            # ==================================
            # 狀態更新
            # ==================================

            current_status = row["status"]

            new_status = st.selectbox(
                "📂 更新狀態",
                RFQ_STATUS,
                index=RFQ_STATUS.index(
                    current_status
                ),
                key=f"status_{row['id']}"
            )

            evidence_update = st.text_area(
                "📝 狀態變更依據",
                value="",
                placeholder="請輸入本次狀態變更的證據，例如：客戶 Email、工程確認、核價單...",
                key=f"evidence_{row['id']}"
            )

            next_step_update = st.text_input(
                "➡️ 下一步",
                value=row["next_step"],
                key=f"next_{row['id']}"
            )

            exception_update = st.text_area(
                "⚠️ 異常",
                value=row["exception"],
                key=f"exception_{row['id']}"
            )

            # ==================================
            # 更新
            # ==================================

            if st.button(
                "💾 更新 RFQ",
                key=f"update_{row['id']}"
            ):

                temp_row = row.copy()

                temp_row["approved_quote"] = approved
                temp_row["quote_sent"] = sent

                valid, error_message = validate_status(
                    temp_row,
                    new_status
                )

                if not valid:

                    st.error(
                        error_message
                    )

                else:

                    old_status = current_status

                    df.loc[
                        real_idx,
                        "status"
                    ] = new_status

                    df.loc[
                        real_idx,
                        "next_step"
                    ] = (
                        next_step_update
                        if next_step_update.strip()
                        else STATUS_NEXT_STEP[
                            new_status
                        ]
                    )

                    df.loc[
                        real_idx,
                        "approved_quote"
                    ] = approved

                    df.loc[
                        real_idx,
                        "quote_sent"
                    ] = sent

                    df.loc[
                        real_idx,
                        "evidence"
                    ] = (
                        evidence_update
                        if evidence_update.strip()
                        else row["evidence"]
                    )

                    df.loc[
                        real_idx,
                        "exception"
                    ] = (
                        exception_update
                        if exception_update.strip()
                        else row["exception"]
                    )

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

                    # 狀態有變更才寫入Log
                    if old_status != new_status:

                        save_status_log(
                            row["RFQ_ID"],
                            old_status,
                            new_status,
                            evidence_update
                            if evidence_update.strip()
                            else "待確認"
                        )

                    st.success(
                        "✅ RFQ 已更新"
                    )

                    st.rerun()

            # ==================================
            # 封存
            # ==================================

            if row["status"] == "結案":

                if st.button(
                    "📦 封存",
                    key=f"archive_{row['id']}"
                ):

                    try:

                        data_df = conn.read(
                            worksheet="Data",
                            ttl=0
                        )

                    except Exception:

                        data_df = pd.DataFrame()

                    if data_df.empty:

                        data_df = pd.DataFrame(
                            columns=df.columns
                        )

                    row_data = df.loc[
                        real_idx
                    ]

                    data_df = pd.concat(
                        [
                            data_df,
                            pd.DataFrame(
                                [row_data]
                            )
                        ],
                        ignore_index=True
                    )

                    conn.update(
                        worksheet="Data",
                        data=data_df
                    )

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
                        "✅ RFQ 已封存"
                    )

                    st.rerun()

            # ==================================
            # 刪除
            # ==================================

            if st.button(
                "🗑️ 刪除",
                key=f"delete_{row['id']}"
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
                    "⚠️ RFQ 已刪除"
                )

                st.rerun()


# ==========================================
# 顯示看板
# ==========================================

for col, status in zip(
    board_cols,
    RFQ_STATUS
):

    with col:

        st.markdown(
            f"## {status}"
        )

        render_tasks(
            df[df["status"] == status],
            status
        )
