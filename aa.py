import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime, timedelta
import uuid
import io

# =========================================================
# 基本設定
# =========================================================

st.set_page_config(
    page_title="企業版 RFQ 詢價追蹤系統",
    page_icon="📌",
    layout="wide"
)

st.title("📌 企業版：RFQ 詢價追蹤管理系統")
st.caption("RFQ Workflow + 報價產生器 | edit by 林溫城")


# =========================================================
# BB 頁面
# =========================================================

top1, top2 = st.columns([8, 1])

with top2:
    st.link_button(
        "📊 BB頁面",
        "https://aazzyybb.streamlit.app/"
    )


# =========================================================
# Google Sheets
# =========================================================

conn = st.connection(
    "gsheets",
    type=GSheetsConnection
)


# =========================================================
# RFQ 欄位
# =========================================================

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


# =========================================================
# 報價欄位
# =========================================================

QUOTE_COLUMNS = [
    "quote_id",
    "RFQ_ID",
    "quote_version",
    "quote_date",
    "customer",
    "product",
    "drawing_version",
    "quantity",
    "currency",
    "unit",
    "material_cost",
    "processing_cost",
    "outsourcing_cost",
    "surface_cost",
    "packing_cost",
    "transport_cost",
    "other_cost",
    "total_cost",
    "management_rate",
    "management_cost",
    "profit_rate",
    "profit_amount",
    "discount",
    "quote_unit_price",
    "quote_total",
    "tax_rate",
    "tax_amount",
    "grand_total",
    "payment_terms",
    "incoterms",
    "delivery",
    "lead_time",
    "validity",
    "remark",
    "approved",
    "approved_time",
    "sent",
    "sent_time",
    "created_time"
]


# =========================================================
# 狀態
# =========================================================

RFQ_STATUS = [
    "新詢價",
    "待補件",
    "工程評估",
    "待核價",
    "已報價",
    "追蹤中",
    "結案"
]


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


# =========================================================
# Boolean 工具
# =========================================================

BOOLEAN_COLUMNS = [
    "approved_quote",
    "quote_sent"
]


def to_bool(value):

    if pd.isna(value):
        return False

    if isinstance(value, bool):
        return value

    value = str(value).strip().lower()

    return value in [
        "true",
        "1",
        "yes",
        "y",
        "是",
        "已核准",
        "已寄送"
    ]


def normalize_boolean_columns(dataframe):

    dataframe = dataframe.copy()

    for col in BOOLEAN_COLUMNS:

        if col not in dataframe.columns:

            dataframe[col] = False

        dataframe[col] = (
            dataframe[col]
            .apply(to_bool)
            .astype(bool)
        )

    return dataframe


# =========================================================
# 數字工具
# =========================================================

def to_float(value):

    if value is None:
        return 0.0

    if isinstance(value, bool):
        return 0.0

    try:

        if pd.isna(value):
            return 0.0

    except Exception:
        pass

    try:

        return float(
            str(value)
            .replace(",", "")
            .replace("%", "")
            .strip()
        )

    except Exception:

        return 0.0


def money(value):

    return f"{to_float(value):,.2f}"


# =========================================================
# 下一個工作日
# =========================================================

def next_workday(date_value):

    next_day = date_value + timedelta(days=1)

    while next_day.weekday() >= 5:

        next_day += timedelta(days=1)

    return next_day


# =========================================================
# RFQ DataFrame 初始化
# =========================================================

try:

    df = conn.read(
        worksheet="Tasks",
        ttl=0
    )

except Exception:

    df = pd.DataFrame()


if df.empty:

    df = pd.DataFrame(
        columns=RFQ_COLUMNS
    )

else:

    # -----------------------------------------
    # 補欄位
    # -----------------------------------------

    for col in RFQ_COLUMNS:

        if col not in df.columns:

            if col == "version":

                df[col] = "V1"

            elif col == "owner":

                df[col] = "業務承辦"

            elif col in BOOLEAN_COLUMNS:

                df[col] = False

            else:

                df[col] = ""

    # -----------------------------------------
    # ID
    # -----------------------------------------

    if "id" not in df.columns:

        df["id"] = [
            str(uuid.uuid4())
            for _ in range(len(df))
        ]

    # -----------------------------------------
    # RFQ ID
    # -----------------------------------------

    if "RFQ_ID" not in df.columns:

        df["RFQ_ID"] = [

            f"RFQ-{datetime.now().strftime('%Y%m%d')}-"
            f"{str(uuid.uuid4())[:4].upper()}"

            for _ in range(len(df))
        ]


# =========================================================
# Boolean 修正
# =========================================================

df = normalize_boolean_columns(df)


# =========================================================
# 狀態驗證
# =========================================================

def validate_status(row, new_status):

    if new_status == "已報價":

        approved = to_bool(
            row.get(
                "approved_quote",
                False
            )
        )

        sent = to_bool(
            row.get(
                "quote_sent",
                False
            )
        )

        if not approved:

            return False, (
                "❌ 不可標記「已報價」："
                "報價尚未核准。"
            )

        if not sent:

            return False, (
                "❌ 不可標記「已報價」："
                "尚未記錄實際寄送。"
            )

    return True, ""


# =========================================================
# Status Log
# =========================================================

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

        "log_id":
            str(uuid.uuid4()),

        "RFQ_ID":
            rfq_id,

        "change_time":
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),

        "old_status":
            old_status,

        "new_status":
            new_status,

        "evidence":
            evidence
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


# =========================================================
# Tabs
# =========================================================

tab_rfq, tab_board, tab_quote, tab_log = st.tabs(
    [
        "📋 RFQ追蹤表",
        "📊 RFQ看板",
        "💰 報價產生器",
        "📜 StatusLog"
    ]
)


# =========================================================
# TAB 1：RFQ 追蹤
# =========================================================

with tab_rfq:

    st.subheader("📝 建立新詢價 RFQ")

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
            placeholder=(
                "例如：客戶 Email、圖面、"
                "詢價單、會議紀錄..."
            )
        )


        new_exception = st.text_area(
            "⚠️ 異常／缺漏",
            placeholder=(
                "若沒有，請填「無」；"
                "資料不足請填「待確認」"
            )
        )


        has_due_date = st.checkbox(
            "設定到期日",
            value=False
        )


        if has_due_date:

            new_due_time = st.date_input(
                "⏰ 到期日"
            )

        else:

            new_due_time = None


        submit_btn = st.form_submit_button(
            "✅ 建立 RFQ"
        )


    # =====================================================
    # 建立 RFQ
    # =====================================================

    if submit_btn:

        if not new_title.strip():

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


            first_followup = next_workday(
                now_dt.date()
            )


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
                    STATUS_NEXT_STEP[
                        "新詢價"
                    ],

                "owner":
                    owner,

                "created_time":
                    now,

                "due_time":
                    due_time,

                "updated_time":
                    now,

                "evidence":
                    new_evidence.strip()
                    if new_evidence.strip()
                    else "待確認",

                "exception":
                    new_exception.strip()
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


            df = normalize_boolean_columns(
                df
            )


            conn.update(
                worksheet="Tasks",
                data=df
            )


            save_status_log(
                rfq_id,
                "無",
                "新詢價",
                new_evidence.strip()
                if new_evidence.strip()
                else "待確認"
            )


            st.success(
                f"✅ RFQ 建立成功：{rfq_id}"
            )

            st.rerun()


    st.divider()


    # =====================================================
    # 快捷搜尋
    # =====================================================

    st.subheader("🔎 RFQ 快捷搜尋")

    search_keyword = st.text_input(
        "搜尋 RFQ、客戶、產品、負責人、證據、異常",
        placeholder=(
            "例如：RFQ-20261007、"
            "ABC、業務承辦..."
        )
    )


    if search_keyword:

        search_mask = (

            df.astype(str)
            .apply(
                lambda col:
                col.str.contains(
                    search_keyword,
                    case=False,
                    na=False,
                    regex=False
                )
            )
            .any(axis=1)
        )

        display_df = df[
            search_mask
        ]

    else:

        display_df = df


    # =====================================================
    # 統計
    # =====================================================

    st.subheader("📊 RFQ 狀態統計")

    stat_cols = st.columns(
        len(RFQ_STATUS)
    )


    for col, status in zip(
        stat_cols,
        RFQ_STATUS
    ):

        count = len(
            df[
                df["status"] == status
            ]
        )

        with col:

            st.metric(
                status,
                count
            )


    # =====================================================
    # 追蹤表
    # =====================================================

    st.subheader("📋 RFQ 追蹤表")


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

        st.info(
            "目前沒有符合條件的 RFQ"
        )

    else:

        st.dataframe(
            display_df[
                tracking_columns
            ],
            use_container_width=True,
            hide_index=True
        )


# =========================================================
# TAB 2：看板
# =========================================================

with tab_board:

    st.subheader("📊 RFQ 管理看板")

    board_cols = st.columns(
        len(RFQ_STATUS)
    )


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
                    f"🔄 最後更新："
                    f"{row['updated_time']}"
                )

                st.caption(
                    f"➡️ 下一步："
                    f"{row['next_step']}"
                )


                # =========================================
                # 首次補問
                # =========================================

                if row["status"] in [
                    "新詢價",
                    "待補件"
                ]:

                    st.info(
                        "📅 首次補問期限："
                        f"{row['first_followup_due']}"
                    )


                # =========================================
                # 報價條件
                # =========================================

                if row["status"] in [
                    "待核價",
                    "已報價",
                    "追蹤中"
                ]:

                    q1, q2 = st.columns(2)


                    with q1:

                        approved = st.checkbox(
                            "✅ 核準報價",
                            value=to_bool(
                                row.get(
                                    "approved_quote",
                                    False
                                )
                            ),
                            key=(
                                f"approved_"
                                f"{row['id']}"
                            )
                        )


                    with q2:

                        sent = st.checkbox(
                            "📤 實際寄送",
                            value=to_bool(
                                row.get(
                                    "quote_sent",
                                    False
                                )
                            ),
                            key=(
                                f"sent_"
                                f"{row['id']}"
                            )
                        )

                else:

                    approved = to_bool(
                        row.get(
                            "approved_quote",
                            False
                        )
                    )

                    sent = to_bool(
                        row.get(
                            "quote_sent",
                            False
                        )
                    )


                # =========================================
                # 狀態
                # =========================================

                current_status = row["status"]


                new_status = st.selectbox(
                    "📂 更新狀態",
                    RFQ_STATUS,
                    index=RFQ_STATUS.index(
                        current_status
                    ),
                    key=(
                        f"status_"
                        f"{row['id']}"
                    )
                )


                evidence_update = st.text_area(
                    "📝 狀態變更依據",
                    value="",
                    placeholder=(
                        "請輸入本次狀態變更的證據，"
                        "例如：客戶 Email、工程確認、核價單..."
                    ),
                    key=(
                        f"evidence_"
                        f"{row['id']}"
                    )
                )


                next_step_update = st.text_input(
                    "➡️ 下一步",
                    value=row["next_step"],
                    key=(
                        f"next_"
                        f"{row['id']}"
                    )
                )


                exception_update = st.text_area(
                    "⚠️ 異常",
                    value=row["exception"],
                    key=(
                        f"exception_"
                        f"{row['id']}"
                    )
                )


                # =========================================
                # 更新
                # =========================================

                if st.button(
                    "💾 更新 RFQ",
                    key=(
                        f"update_"
                        f"{row['id']}"
                    )
                ):

                    temp_row = row.copy()

                    temp_row[
                        "approved_quote"
                    ] = bool(approved)

                    temp_row[
                        "quote_sent"
                    ] = bool(sent)


                    valid, error_message = (
                        validate_status(
                            temp_row,
                            new_status
                        )
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
                            next_step_update.strip()
                            if next_step_update.strip()
                            else STATUS_NEXT_STEP[
                                new_status
                            ]
                        )


                        # ---------------------------------
                        # 關鍵：強制 Boolean
                        # ---------------------------------

                        df.loc[
                            real_idx,
                            "approved_quote"
                        ] = bool(approved)


                        df.loc[
                            real_idx,
                            "quote_sent"
                        ] = bool(sent)


                        df.loc[
                            real_idx,
                            "evidence"
                        ] = (
                            evidence_update.strip()
                            if evidence_update.strip()
                            else row["evidence"]
                        )


                        df.loc[
                            real_idx,
                            "exception"
                        ] = (
                            exception_update.strip()
                            if exception_update.strip()
                            else row["exception"]
                        )


                        df.loc[
                            real_idx,
                            "updated_time"
                        ] = datetime.now().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        )


                        # ---------------------------------
                        # 儲存前重新標準化 Boolean
                        # ---------------------------------

                        df = normalize_boolean_columns(
                            df
                        )


                        conn.update(
                            worksheet="Tasks",
                            data=df
                        )


                        # ---------------------------------
                        # StatusLog
                        # ---------------------------------

                        if old_status != new_status:

                            save_status_log(
                                row["RFQ_ID"],
                                old_status,
                                new_status,
                                (
                                    evidence_update.strip()
                                    if evidence_update.strip()
                                    else "待確認"
                                )
                            )


                        st.success(
                            "✅ RFQ 已更新"
                        )

                        st.rerun()


                # =========================================
                # 顯示證據
                # =========================================

                with st.expander(
                    "📎 查看證據／異常"
                ):

                    st.write(
                        "**證據：**"
                    )

                    st.write(
                        row["evidence"]
                    )

                    st.write(
                        "**異常：**"
                    )

                    st.write(
                        row["exception"]
                    )


                # =========================================
                # 封存
                # =========================================

                if row["status"] == "結案":

                    if st.button(
                        "📦 封存",
                        key=(
                            f"archive_"
                            f"{row['id']}"
                        )
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


                        df = normalize_boolean_columns(
                            df
                        )


                        conn.update(
                            worksheet="Tasks",
                            data=df
                        )


                        st.success(
                            "✅ RFQ 已封存"
                        )

                        st.rerun()


                # =========================================
                # 刪除
                # =========================================

                if st.button(
                    "🗑️ 刪除",
                    key=(
                        f"delete_"
                        f"{row['id']}"
                    )
                ):

                    df.drop(
                        real_idx,
                        inplace=True
                    )


                    df.reset_index(
                        drop=True,
                        inplace=True
                    )


                    df = normalize_boolean_columns(
                        df
                    )


                    conn.update(
                        worksheet="Tasks",
                        data=df
                    )


                    st.warning(
                        "⚠️ RFQ 已刪除"
                    )

                    st.rerun()


    # =============================================
    # 看板顯示
    # =============================================

    for col, status in zip(
        board_cols,
        RFQ_STATUS
    ):

        with col:

            st.markdown(
                f"## {status}"
            )

            render_tasks(
                df[
                    df["status"] == status
                ],
                status
            )


# =========================================================
# TAB 3：報價產生器
# =========================================================

with tab_quote:

    st.subheader("💰 RFQ 報價產生器")

    if df.empty:

        st.info(
            "目前沒有 RFQ，請先建立 RFQ。"
        )

    else:

        rfq_options = df[
            "RFQ_ID"
        ].dropna().astype(str).tolist()


        selected_rfq = st.selectbox(
            "選擇 RFQ",
            rfq_options
        )


        rfq_matches = df[
            df["RFQ_ID"].astype(str)
            == str(selected_rfq)
        ]


        if not rfq_matches.empty:

            rfq = rfq_matches.iloc[0]


            # =============================================
            # RFQ 基本資料
            # =============================================

            st.markdown("### 📋 RFQ 基本資料")

            info1, info2, info3, info4 = st.columns(4)


            with info1:

                st.metric(
                    "RFQ",
                    rfq["RFQ_ID"]
                )


            with info2:

                st.metric(
                    "客戶",
                    rfq["customer"]
                )


            with info3:

                st.metric(
                    "產品",
                    rfq["title"]
                )


            with info4:

                st.metric(
                    "目前狀態",
                    rfq["status"]
                )


            st.divider()


            # =============================================
            # 報價表單
            # =============================================

            st.markdown(
                "### 🧮 成本與報價計算"
            )


            c1, c2, c3 = st.columns(3)


            with c1:

                quote_version = st.text_input(
                    "報價版本",
                    value="V1"
                )

                quote_date = st.date_input(
                    "報價日期",
                    value=datetime.now().date()
                )

                quantity = st.number_input(
                    "數量",
                    min_value=1.0,
                    value=1.0,
                    step=1.0
                )


            with c2:

                currency = st.selectbox(
                    "幣別",
                    [
                        "TWD",
                        "USD",
                        "EUR",
                        "JPY",
                        "CNY"
                    ]
                )

                unit = st.text_input(
                    "單位",
                    value="PCS"
                )

                drawing_version = st.text_input(
                    "圖面版本",
                    value=""
                )


            with c3:

                payment_terms = st.text_input(
                    "付款條件",
                    placeholder="例如：T/T 30 days"
                )

                incoterms = st.text_input(
                    "交易條件",
                    placeholder="例如：FOB / CIF / EXW"
                )

                delivery = st.text_input(
                    "交貨地",
                    placeholder="待確認"
                )


            st.markdown("#### 💵 成本")

            cost1, cost2, cost3, cost4 = st.columns(4)


            with cost1:

                material_cost = st.number_input(
                    "材料成本",
                    min_value=0.0,
                    value=0.0,
                    step=100.0
                )

                processing_cost = st.number_input(
                    "加工成本",
                    min_value=0.0,
                    value=0.0,
                    step=100.0
                )


            with cost2:

                outsourcing_cost = st.number_input(
                    "外包成本",
                    min_value=0.0,
                    value=0.0,
                    step=100.0
                )

                surface_cost = st.number_input(
                    "表面處理",
                    min_value=0.0,
                    value=0.0,
                    step=100.0
                )


            with cost3:

                packing_cost = st.number_input(
                    "包裝成本",
                    min_value=0.0,
                    value=0.0,
                    step=100.0
                )

                transport_cost = st.number_input(
                    "運輸成本",
                    min_value=0.0,
                    value=0.0,
                    step=100.0
                )


            with cost4:

                other_cost = st.number_input(
                    "其他成本",
                    min_value=0.0,
                    value=0.0,
                    step=100.0
                )


            # =============================================
            # 計算
            # =============================================

            total_cost = (
                material_cost
                + processing_cost
                + outsourcing_cost
                + surface_cost
                + packing_cost
                + transport_cost
                + other_cost
            )


            st.markdown("#### 📈 利潤與價格")

            p1, p2, p3, p4 = st.columns(4)


            with p1:

                management_rate = st.number_input(
                    "管理費 %",
                    min_value=0.0,
                    max_value=100.0,
                    value=0.0,
                    step=1.0
                )


            with p2:

                profit_rate = st.number_input(
                    "利潤率 %",
                    min_value=0.0,
                    max_value=100.0,
                    value=20.0,
                    step=1.0
                )


            with p3:

                discount = st.number_input(
                    "折扣",
                    min_value=0.0,
                    value=0.0,
                    step=100.0
                )


            with p4:

                tax_rate = st.number_input(
                    "稅率 %",
                    min_value=0.0,
                    max_value=100.0,
                    value=5.0,
                    step=1.0
                )


            management_cost = (
                total_cost
                * management_rate
                / 100
            )


            profit_base = (
                total_cost
                + management_cost
            )


            profit_amount = (
                profit_base
                * profit_rate
                / 100
            )


            quote_total_before_discount = (
                profit_base
                + profit_amount
            )


            quote_total = max(
                0,
                quote_total_before_discount
                - discount
            )


            if quantity > 0:

                quote_unit_price = (
                    quote_total
                    / quantity
                )

            else:

                quote_unit_price = 0


            tax_amount = (
                quote_total
                * tax_rate
                / 100
            )


            grand_total = (
                quote_total
                + tax_amount
            )


            # =============================================
            # 報價結果
            # =============================================

            st.divider()

            st.markdown(
                "### 📊 報價結果"
            )


            result1, result2, result3, result4 = (
                st.columns(4)
            )


            with result1:

                st.metric(
                    "總成本",
                    f"{currency} "
                    f"{money(total_cost)}"
                )


            with result2:

                st.metric(
                    "毛利",
                    f"{currency} "
                    f"{money(profit_amount)}"
                )


            with result3:

                st.metric(
                    "未稅報價",
                    f"{currency} "
                    f"{money(quote_total)}"
                )


            with result4:

                st.metric(
                    "含稅總額",
                    f"{currency} "
                    f"{money(grand_total)}"
                )


            gross_margin = 0

            if quote_total > 0:

                gross_margin = (
                    quote_total
                    - total_cost
                ) / quote_total * 100


            st.info(
                f"📌 毛利率：{gross_margin:.2f}%  "
                f"| 單價：{currency} "
                f"{money(quote_unit_price)} / {unit}"
            )


            # =============================================
            # 交期
            # =============================================

            lead_time = st.text_input(
                "⏱️ 交期",
                placeholder=(
                    "例如：規格與訂單確認後 "
                    "依生產排程評估"
                )
            )


            validity = st.text_input(
                "📅 報價有效期限",
                value="30 days"
            )


            remark = st.text_area(
                "📝 報價備註",
                placeholder=(
                    "請輸入客戶版報價需要顯示的備註"
                )
            )


            # =============================================
            # 產生報價
            # =============================================

            st.divider()

            if st.button(
                "💾 儲存報價草稿",
                type="primary",
                use_container_width=True
            ):

                # -----------------------------------------
                # RFQ 基本防呆
                # -----------------------------------------

                if rfq["status"] not in [
                    "待核價",
                    "已報價",
                    "追蹤中"
                ]:

                    st.warning(
                        "⚠️ 目前 RFQ 狀態不是「待核價」。"
                        "仍可儲存報價草稿，但正式報價前請確認工程評估完成。"
                    )


                quote_id = (
                    f"QT-"
                    f"{datetime.now().strftime('%Y%m%d')}-"
                    f"{str(uuid.uuid4())[:6].upper()}"
                )


                quote_data = {

                    "quote_id":
                        quote_id,

                    "RFQ_ID":
                        rfq["RFQ_ID"],

                    "quote_version":
                        quote_version,

                    "quote_date":
                        quote_date.strftime(
                            "%Y-%m-%d"
                        ),

                    "customer":
                        rfq["customer"],

                    "product":
                        rfq["title"],

                    "drawing_version":
                        drawing_version
                        if drawing_version.strip()
                        else "待確認",

                    "quantity":
                        quantity,

                    "currency":
                        currency,

                    "unit":
                        unit,

                    "material_cost":
                        material_cost,

                    "processing_cost":
                        processing_cost,

                    "outsourcing_cost":
                        outsourcing_cost,

                    "surface_cost":
                        surface_cost,

                    "packing_cost":
                        packing_cost,

                    "transport_cost":
                        transport_cost,

                    "other_cost":
                        other_cost,

                    "total_cost":
                        total_cost,

                    "management_rate":
                        management_rate,

                    "management_cost":
                        management_cost,

                    "profit_rate":
                        profit_rate,

                    "profit_amount":
                        profit_amount,

                    "discount":
                        discount,

                    "quote_unit_price":
                        quote_unit_price,

                    "quote_total":
                        quote_total,

                    "tax_rate":
                        tax_rate,

                    "tax_amount":
                        tax_amount,

                    "grand_total":
                        grand_total,

                    "payment_terms":
                        payment_terms,

                    "incoterms":
                        incoterms,

                    "delivery":
                        delivery
                        if delivery.strip()
                        else "待確認",

                    "lead_time":
                        lead_time
                        if lead_time.strip()
                        else "待確認",

                    "validity":
                        validity,

                    "remark":
                        remark,

                    "approved":
                        False,

                    "approved_time":
                        "",

                    "sent":
                        False,

                    "sent_time":
                        "",

                    "created_time":
                        datetime.now().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        )
                }


                try:

                    quote_df = conn.read(
                        worksheet="Quotes",
                        ttl=0
                    )

                except Exception:

                    quote_df = pd.DataFrame()


                if quote_df.empty:

                    quote_df = pd.DataFrame(
                        columns=QUOTE_COLUMNS
                    )

                else:

                    for col in QUOTE_COLUMNS:

                        if col not in quote_df.columns:

                            quote_df[col] = ""


                quote_df = pd.concat(
                    [
                        quote_df,
                        pd.DataFrame(
                            [quote_data]
                        )
                    ],
                    ignore_index=True
                )


                conn.update(
                    worksheet="Quotes",
                    data=quote_df
                )


                st.success(
                    f"✅ 報價草稿已儲存：{quote_id}"
                )


            # =============================================
            # 報價明細
            # =============================================

            st.divider()

            st.markdown(
                "### 📄 報價摘要"
            )


            summary_df = pd.DataFrame(
                {
                    "項目": [
                        "RFQ",
                        "客戶",
                        "產品",
                        "數量",
                        "幣別",
                        "總成本",
                        "管理費",
                        "利潤",
                        "折扣",
                        "未稅報價",
                        "稅額",
                        "含稅總額",
                        "單價",
                        "交期",
                        "有效期限"
                    ],

                    "內容": [
                        rfq["RFQ_ID"],
                        rfq["customer"],
                        rfq["title"],
                        quantity,
                        currency,
                        money(total_cost),
                        money(management_cost),
                        money(profit_amount),
                        money(discount),
                        money(quote_total),
                        money(tax_amount),
                        money(grand_total),
                        money(quote_unit_price),
                        lead_time
                        if lead_time.strip()
                        else "待確認",
                        validity
                    ]
                }
            )


            st.dataframe(
                summary_df,
                use_container_width=True,
                hide_index=True
            )


            # =============================================
            # Excel 下載
            # =============================================

            export_df = pd.DataFrame(
                [quote_data]
            )


            excel_buffer = io.BytesIO()


            try:

                with pd.ExcelWriter(
                    excel_buffer,
                    engine="openpyxl"
                ) as writer:

                    export_df.to_excel(
                        writer,
                        index=False,
                        sheet_name="Quotation"
                    )


                st.download_button(
                    "📊 下載 Excel 報價明細",
                    data=excel_buffer.getvalue(),
                    file_name=(
                        f"{quote_id}.xlsx"
                    ),
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    )
                )

            except Exception as e:

                st.warning(
                    f"Excel 產生失敗：{e}"
                )


# =========================================================
# TAB 4：StatusLog
# =========================================================

with tab_log:

    st.subheader("📜 RFQ 狀態異動紀錄")


    try:

        log_df = conn.read(
            worksheet="StatusLog",
            ttl=0
        )

    except Exception:

        log_df = pd.DataFrame()


    if log_df.empty:

        st.info(
            "目前沒有狀態紀錄"
        )

    else:

        st.dataframe(
            log_df.sort_values(
                by="change_time",
                ascending=False
            ),
            use_container_width=True,
            hide_index=True
        )
