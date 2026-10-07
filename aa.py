import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime, timedelta
import uuid
import io


# =========================================================
# 1. 基本設定
# =========================================================

st.set_page_config(
    page_title="企業版 RFQ 詢價追蹤系統",
    page_icon="📌",
    layout="wide"
)

st.title("📌 企業版：RFQ 詢價追蹤管理系統")
st.caption("RFQ Workflow + 報價產生器 | edit by 林溫城")


# =========================================================
# 2. BB 頁面
# =========================================================

top1, top2 = st.columns([8, 1])

with top2:

    st.link_button(
        "📊 BB頁面",
        "https://aazzyybb.streamlit.app/"
    )


# =========================================================
# 3. Google Sheets
# =========================================================

conn = st.connection(
    "gsheets",
    type=GSheetsConnection
)


# =========================================================
# 4. RFQ 欄位
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
# 5. Quote 欄位
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
# 6. RFQ 狀態
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
# 7. Boolean 工具
# =========================================================

BOOLEAN_COLUMNS = [

    "approved_quote",
    "quote_sent"

]


def to_bool(value):

    if value is None:
        return False

    if isinstance(value, bool):
        return value

    try:

        if pd.isna(value):
            return False

    except Exception:
        pass

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
# 8. 數字工具
# =========================================================

def to_float(value):

    if value is None:
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
# 9. 下一個工作日
# =========================================================

def next_workday(date_value):

    next_day = date_value + timedelta(days=1)

    while next_day.weekday() >= 5:

        next_day += timedelta(days=1)

    return next_day


# =========================================================
# 10. Status 驗證
# =========================================================

def validate_status(row, new_status):

    if new_status != "已報價":

        return True, ""


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
# 11. 讀取 Tasks
# =========================================================

try:

    df = conn.read(
        worksheet="Tasks",
        ttl=0
    )

except Exception as e:

    st.error(
        "❌ 無法讀取 Google Sheets 的 Tasks"
    )

    st.exception(e)

    st.stop()


# =========================================================
# 12. 初始化 Tasks
# =========================================================

if df is None:

    df = pd.DataFrame()


if df.empty:

    df = pd.DataFrame(
        columns=RFQ_COLUMNS
    )

else:

    # -----------------------------------------------------
    # 補欄位
    # -----------------------------------------------------

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


# =========================================================
# 13. ID 防呆
# =========================================================

if "id" not in df.columns:

    df["id"] = ""


for idx in df.index:

    current_id = df.loc[idx, "id"]

    if (
        current_id is None
        or str(current_id).strip() == ""
        or str(current_id).lower() == "nan"
    ):

        df.loc[idx, "id"] = str(
            uuid.uuid4()
        )


# =========================================================
# 14. RFQ_ID 防呆
# =========================================================

if "RFQ_ID" not in df.columns:

    df["RFQ_ID"] = ""


for idx in df.index:

    current_rfq = df.loc[
        idx,
        "RFQ_ID"
    ]

    if (
        current_rfq is None
        or str(current_rfq).strip() == ""
        or str(current_rfq).lower() == "nan"
    ):

        df.loc[
            idx,
            "RFQ_ID"
        ] = (
            f"RFQ-"
            f"{datetime.now().strftime('%Y%m%d')}-"
            f"{str(uuid.uuid4())[:4].upper()}"
        )


# =========================================================
# 15. 狀態防呆
# =========================================================

if "status" not in df.columns:

    df["status"] = "新詢價"


df["status"] = (
    df["status"]
    .fillna("新詢價")
    .astype(str)
)


# =========================================================
# 16. Boolean 正規化
# =========================================================

df = normalize_boolean_columns(df)


# =========================================================
# 17. StatusLog
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


    if log_df is None:

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


    for col in log_columns:

        if col not in log_df.columns:

            log_df[col] = ""


    new_log = {

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
            pd.DataFrame(
                [new_log]
            )
        ],

        ignore_index=True

    )


    try:

        conn.update(
            worksheet="StatusLog",
            data=log_df
        )

    except Exception as e:

        st.warning(
            "⚠️ StatusLog 寫入失敗"
        )

        st.exception(e)


# =========================================================
# 18. 頁面 Tabs
# =========================================================

tab_rfq, tab_board, tab_quote, tab_log = st.tabs(

    [
        "📋 RFQ追蹤表",
        "📊 RFQ看板",
        "💰 報價產生器",
        "📜 StatusLog"
    ]

)


# ################################################################
# TAB 1：RFQ
# ################################################################

with tab_rfq:

    st.subheader(
        "📝 建立新詢價 RFQ"
    )


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
            "設定到期日"
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

                "RFQ-"
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
                    (
                        new_version.strip()
                        if new_version.strip()
                        else "V1"
                    ),

                "department":
                    new_department,

                "customer":
                    new_customer.strip(),

                "title":
                    new_title.strip(),

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
                    (
                        new_evidence.strip()
                        if new_evidence.strip()
                        else "待確認"
                    ),

                "exception":
                    (
                        new_exception.strip()
                        if new_exception.strip()
                        else "待確認"
                    ),

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
                    pd.DataFrame(
                        [new_data]
                    )
                ],

                ignore_index=True

            )


            df = normalize_boolean_columns(
                df
            )


            try:

                conn.update(
                    worksheet="Tasks",
                    data=df
                )


                save_status_log(

                    rfq_id,
                    "無",
                    "新詢價",

                    (
                        new_evidence.strip()
                        if new_evidence.strip()
                        else "待確認"
                    )

                )


                st.success(
                    f"✅ RFQ 建立成功：{rfq_id}"
                )


                st.rerun()


            except Exception as e:

                st.error(
                    "❌ RFQ 寫入 Google Sheets 失敗"
                )

                st.exception(e)


    st.divider()


    # =====================================================
    # 快捷搜尋
    # =====================================================

    st.subheader(
        "🔎 RFQ 快捷搜尋"
    )


    search_keyword = st.text_input(

        "搜尋 RFQ、客戶、詢價名稱、負責人、證據、異常",

        placeholder=(
            "例如：RFQ-20261007、"
            "ABC、業務承辦"
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
        ].copy()

    else:

        display_df = df.copy()


    # =====================================================
    # 統計
    # =====================================================

    st.subheader(
        "📊 RFQ 狀態統計"
    )


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

    st.subheader(
        "📋 RFQ 追蹤表"
    )


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

        # 防止舊資料缺欄位造成 KeyError

        safe_columns = [

            col
            for col in tracking_columns
            if col in display_df.columns

        ]


        st.dataframe(

            display_df[
                safe_columns
            ],

            use_container_width=True,

            hide_index=True

        )


# ################################################################
# TAB 2：KANBAN
# ################################################################

with tab_board:

    st.subheader(
        "📊 RFQ 管理看板"
    )


    board_cols = st.columns(
        len(RFQ_STATUS)
    )


    # =========================================================
    # Render Tasks
    #
    # 重要：
    # 不再依賴 render_tasks 外部重新搜尋 df。
    # 直接使用傳入的 full_df。
    # =========================================================

    def render_tasks(
        task_df,
        column_name,
        full_df
    ):

        # -----------------------------------------------------
        # 防呆 1
        # -----------------------------------------------------

        if task_df is None:

            st.warning(
                f"⚠️ {column_name} 資料為 None"
            )

            return


        # -----------------------------------------------------
        # 防呆 2
        # -----------------------------------------------------

        if task_df.empty:

            st.info(
                f"目前沒有 {column_name}"
            )

            return


        # -----------------------------------------------------
        # 防呆 3
        # -----------------------------------------------------

        if full_df is None:

            st.error(
                "❌ 完整 RFQ DataFrame 不存在"
            )

            return


        # -----------------------------------------------------
        # 防呆 4
        # -----------------------------------------------------

        if "id" not in full_df.columns:

            st.error(
                "❌ Tasks 缺少 id 欄位"
            )

            st.write(
                "目前 Tasks 欄位：",
                list(full_df.columns)
            )

            return


        # -----------------------------------------------------
        # 每張卡片
        # -----------------------------------------------------

        for _, row in task_df.iterrows():

            row_id = row.get(
                "id",
                ""
            )


            # -------------------------------------------------
            # 找原始資料 index
            # -------------------------------------------------

            try:

                matches = full_df[
                    full_df["id"].astype(str)
                    == str(row_id)
                ].index

            except Exception as e:

                st.error(
                    "❌ 搜尋 RFQ id 時發生錯誤"
                )

                st.exception(e)

                continue


            if len(matches) == 0:

                st.warning(
                    f"⚠️ 找不到 RFQ："
                    f"{row.get('RFQ_ID', '未知')}"
                )

                continue


            real_idx = matches[0]


            # -------------------------------------------------
            # 卡片
            # -------------------------------------------------

            with st.container(
                border=True
            ):

                st.markdown(
                    f"### 📌 {row.get('title', '未命名')}"
                )


                st.caption(
                    f"🆔 RFQ："
                    f"{row.get('RFQ_ID', '')}"
                )


                st.caption(
                    f"📑 版本："
                    f"{row.get('version', '')}"
                )


                st.caption(
                    f"🏢 部門："
                    f"{row.get('department', '')}"
                )


                st.caption(
                    f"🏭 客戶："
                    f"{row.get('customer', '')}"
                )


                st.caption(
                    f"👤 負責人："
                    f"{row.get('owner', '')}"
                )


                st.caption(
                    f"📅 到期日："
                    f"{row.get('due_time', '待確認')}"
                )


                st.caption(
                    f"🔄 最後更新："
                    f"{row.get('updated_time', '')}"
                )


                st.caption(
                    f"➡️ 下一步："
                    f"{row.get('next_step', '')}"
                )


                # =================================================
                # 首次補問
                # =================================================

                if row.get("status") in [

                    "新詢價",
                    "待補件"

                ]:

                    st.info(

                        "📅 首次補問期限："
                        f"{row.get('first_followup_due', '待確認')}"

                    )


                # =================================================
                # 報價 Boolean
                # =================================================

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


                if row.get("status") in [

                    "待核價",
                    "已報價",
                    "追蹤中"

                ]:

                    q1, q2 = st.columns(2)


                    with q1:

                        approved = st.checkbox(

                            "✅ 核準報價",

                            value=approved,

                            key=(
                                "approved_"
                                f"{row_id}"
                            )

                        )


                    with q2:

                        sent = st.checkbox(

                            "📤 實際寄送",

                            value=sent,

                            key=(
                                "sent_"
                                f"{row_id}"
                            )

                        )


                # =================================================
                # 狀態
                # =================================================

                current_status = str(

                    row.get(
                        "status",
                        "新詢價"
                    )

                )


                if current_status not in RFQ_STATUS:

                    current_status = "新詢價"


                new_status = st.selectbox(

                    "📂 更新狀態",

                    RFQ_STATUS,

                    index=RFQ_STATUS.index(
                        current_status
                    ),

                    key=(
                        "status_"
                        f"{row_id}"
                    )

                )


                evidence_update = st.text_area(

                    "📝 狀態變更依據",

                    value="",

                    placeholder=(
                        "例如：客戶 Email、"
                        "工程確認、核價單..."
                    ),

                    key=(
                        "evidence_"
                        f"{row_id}"
                    )

                )


                next_step_update = st.text_input(

                    "➡️ 下一步",

                    value=str(
                        row.get(
                            "next_step",
                            ""
                        )
                    ),

                    key=(
                        "next_"
                        f"{row_id}"
                    )

                )


                exception_update = st.text_area(

                    "⚠️ 異常",

                    value=str(
                        row.get(
                            "exception",
                            ""
                        )
                    ),

                    key=(
                        "exception_"
                        f"{row_id}"
                    )

                )


                # =================================================
                # 更新 RFQ
                # =================================================

                if st.button(

                    "💾 更新 RFQ",

                    key=(
                        "update_"
                        f"{row_id}"
                    )

                ):

                    # ---------------------------------------------
                    # 建立暫存資料
                    # ---------------------------------------------

                    temp_row = row.copy()


                    temp_row[
                        "approved_quote"
                    ] = bool(approved)


                    temp_row[
                        "quote_sent"
                    ] = bool(sent)


                    # ---------------------------------------------
                    # 驗證
                    # ---------------------------------------------

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


                        # -----------------------------------------
                        # 使用 full_df
                        # -----------------------------------------

                        full_df.loc[
                            real_idx,
                            "status"
                        ] = new_status


                        full_df.loc[
                            real_idx,
                            "next_step"
                        ] = (

                            next_step_update.strip()

                            if next_step_update.strip()

                            else STATUS_NEXT_STEP[
                                new_status
                            ]

                        )


                        # -----------------------------------------
                        # Boolean
                        # -----------------------------------------

                        full_df.loc[
                            real_idx,
                            "approved_quote"
                        ] = bool(approved)


                        full_df.loc[
                            real_idx,
                            "quote_sent"
                        ] = bool(sent)


                        # -----------------------------------------
                        # 證據
                        # -----------------------------------------

                        if evidence_update.strip():

                            full_df.loc[
                                real_idx,
                                "evidence"
                            ] = (
                                evidence_update.strip()
                            )


                        # -----------------------------------------
                        # 異常
                        # -----------------------------------------

                        if exception_update.strip():

                            full_df.loc[
                                real_idx,
                                "exception"
                            ] = (
                                exception_update.strip()
                            )


                        # -----------------------------------------
                        # 更新時間
                        # -----------------------------------------

                        full_df.loc[
                            real_idx,
                            "updated_time"
                        ] = datetime.now().strftime(

                            "%Y-%m-%d %H:%M:%S"

                        )


                        # -----------------------------------------
                        # Boolean 再次正規化
                        # -----------------------------------------

                        full_df = (
                            normalize_boolean_columns(
                                full_df
                            )
                        )


                        # -----------------------------------------
                        # 寫入 Google Sheets
                        # -----------------------------------------

                        try:

                            conn.update(

                                worksheet="Tasks",

                                data=full_df

                            )


                            # -------------------------------------
                            # Status Log
                            # -------------------------------------

                            if old_status != new_status:

                                save_status_log(

                                    row.get(
                                        "RFQ_ID",
                                        ""
                                    ),

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


                        except Exception as e:

                            st.error(
                                "❌ RFQ 更新失敗"
                            )

                            st.exception(e)


                # =================================================
                # 查看證據
                # =================================================

                with st.expander(
                    "📎 查看證據／異常"
                ):

                    st.write(
                        "**證據：**"
                    )

                    st.write(
                        row.get(
                            "evidence",
                            "無"
                        )
                    )


                    st.write(
                        "**異常：**"
                    )

                    st.write(
                        row.get(
                            "exception",
                            "無"
                        )
                    )


                # =================================================
                # 封存
                # =================================================

                if current_status == "結案":

                    if st.button(

                        "📦 封存",

                        key=(
                            "archive_"
                            f"{row_id}"
                        )

                    ):

                        try:

                            data_df = conn.read(

                                worksheet="Data",

                                ttl=0

                            )

                        except Exception:

                            data_df = pd.DataFrame()


                        if data_df is None:

                            data_df = pd.DataFrame()


                        if data_df.empty:

                            data_df = pd.DataFrame(
                                columns=full_df.columns
                            )


                        row_data = full_df.loc[
                            real_idx
                        ].copy()


                        data_df = pd.concat(

                            [

                                data_df,

                                pd.DataFrame(
                                    [row_data]
                                )

                            ],

                            ignore_index=True

                        )


                        try:

                            conn.update(

                                worksheet="Data",

                                data=data_df

                            )


                            full_df = full_df.drop(
                                index=real_idx
                            )


                            full_df = full_df.reset_index(
                                drop=True
                            )


                            full_df = (
                                normalize_boolean_columns(
                                    full_df
                                )
                            )


                            conn.update(

                                worksheet="Tasks",

                                data=full_df

                            )


                            st.success(
                                "✅ RFQ 已封存"
                            )


                            st.rerun()


                        except Exception as e:

                            st.error(
                                "❌ RFQ 封存失敗"
                            )

                            st.exception(e)


                # =================================================
                # 刪除
                # =================================================

                if st.button(

                    "🗑️ 刪除",

                    key=(
                        "delete_"
                        f"{row_id}"
                    )

                ):

                    try:

                        full_df = full_df.drop(
                            index=real_idx
                        )


                        full_df = full_df.reset_index(
                            drop=True
                        )


                        full_df = (
                            normalize_boolean_columns(
                                full_df
                            )
                        )


                        conn.update(

                            worksheet="Tasks",

                            data=full_df

                        )


                        st.warning(
                            "⚠️ RFQ 已刪除"
                        )


                        st.rerun()


                    except Exception as e:

                        st.error(
                            "❌ RFQ 刪除失敗"
                        )

                        st.exception(e)


    # =========================================================
    # 顯示 Kanban
    # =========================================================

    for col, status in zip(

        board_cols,

        RFQ_STATUS

    ):

        with col:

            st.markdown(
                f"## {status}"
            )


            # ---------------------------------------------
            # 重要：
            # task_df 是副本
            # full_df 是完整資料
            # ---------------------------------------------

            task_df = df[
                df["status"] == status
            ].copy()


            render_tasks(

                task_df,

                status,

                df

            )


# ################################################################
# TAB 3：報價產生器
# ################################################################

with tab_quote:

    st.subheader(
        "💰 RFQ 報價產生器"
    )


    if df.empty:

        st.info(
            "目前沒有 RFQ，請先建立 RFQ。"
        )

    else:

        rfq_options = (

            df["RFQ_ID"]
            .dropna()
            .astype(str)
            .tolist()

        )


        if not rfq_options:

            st.info(
                "目前沒有有效 RFQ_ID。"
            )

        else:

            selected_rfq = st.selectbox(

                "選擇 RFQ",

                rfq_options

            )


            rfq_matches = df[

                df["RFQ_ID"].astype(str)
                == str(selected_rfq)

            ]


            if rfq_matches.empty:

                st.error(
                    "❌ 找不到選擇的 RFQ"
                )

            else:

                rfq = rfq_matches.iloc[0]


                # =================================================
                # RFQ 資料
                # =================================================

                st.markdown(
                    "### 📋 RFQ 基本資料"
                )


                c1, c2, c3, c4 = st.columns(4)


                with c1:

                    st.metric(
                        "RFQ",
                        str(
                            rfq.get(
                                "RFQ_ID",
                                ""
                            )
                        )
                    )


                with c2:

                    st.metric(
                        "客戶",
                        str(
                            rfq.get(
                                "customer",
                                ""
                            )
                        )
                    )


                with c3:

                    st.metric(
                        "詢價名稱",
                        str(
                            rfq.get(
                                "title",
                                ""
                            )
                        )
                    )


                with c4:

                    st.metric(
                        "目前狀態",
                        str(
                            rfq.get(
                                "status",
                                ""
                            )
                        )
                    )


                st.divider()


                # =================================================
                # 報價基本資料
                # =================================================

                st.markdown(
                    "### 🧾 報價基本資料"
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
                        "圖面版本"
                    )


                with c3:

                    payment_terms = st.text_input(
                        "付款條件"
                    )


                    incoterms = st.text_input(
                        "交易條件"
                    )


                    delivery = st.text_input(
                        "交貨地"
                    )


                # =================================================
                # 成本
                # =================================================

                st.markdown(
                    "### 💵 成本"
                )


                c1, c2, c3, c4 = st.columns(4)


                with c1:

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


                with c2:

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


                with c3:

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


                with c4:

                    other_cost = st.number_input(
                        "其他成本",
                        min_value=0.0,
                        value=0.0,
                        step=100.0
                    )


                # =================================================
                # 計算
                # =================================================

                total_cost = (

                    material_cost
                    + processing_cost
                    + outsourcing_cost
                    + surface_cost
                    + packing_cost
                    + transport_cost
                    + other_cost

                )


                st.markdown(
                    "### 📈 報價參數"
                )


                c1, c2, c3, c4 = st.columns(4)


                with c1:

                    management_rate = st.number_input(
                        "管理費 %",
                        min_value=0.0,
                        max_value=100.0,
                        value=0.0,
                        step=1.0
                    )


                with c2:

                    profit_rate = st.number_input(
                        "利潤率 %",
                        min_value=0.0,
                        max_value=100.0,
                        value=20.0,
                        step=1.0
                    )


                with c3:

                    discount = st.number_input(
                        "折扣金額",
                        min_value=0.0,
                        value=0.0,
                        step=100.0
                    )


                with c4:

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


                before_discount = (

                    profit_base
                    + profit_amount

                )


                quote_total = max(

                    0,

                    before_discount
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


                gross_margin = 0


                if quote_total > 0:

                    gross_margin = (

                        quote_total
                        - total_cost

                    ) / quote_total * 100


                # =================================================
                # 報價結果
                # =================================================

                st.divider()


                st.markdown(
                    "### 📊 報價結果"
                )


                c1, c2, c3, c4 = st.columns(4)


                with c1:

                    st.metric(
                        "總成本",
                        f"{currency} "
                        f"{money(total_cost)}"
                    )


                with c2:

                    st.metric(
                        "利潤",
                        f"{currency} "
                        f"{money(profit_amount)}"
                    )


                with c3:

                    st.metric(
                        "未稅報價",
                        f"{currency} "
                        f"{money(quote_total)}"
                    )


                with c4:

                    st.metric(
                        "含稅總額",
                        f"{currency} "
                        f"{money(grand_total)}"
                    )


                st.info(

                    f"📌 毛利率："
                    f"{gross_margin:.2f}%"
                    f"　|　單價："
                    f"{currency} "
                    f"{money(quote_unit_price)}"
                    f" / {unit}"

                )


                # =================================================
                # 交期
                # =================================================

                lead_time = st.text_input(
                    "⏱️ 交期",
                    placeholder="例如：確認訂單後 30 天"
                )


                validity = st.text_input(
                    "📅 報價有效期限",
                    value="30 days"
                )


                remark = st.text_area(
                    "📝 報價備註"
                )


                # =================================================
                # 儲存報價
                # =================================================

                if st.button(

                    "💾 儲存報價草稿",

                    type="primary",

                    use_container_width=True

                ):

                    quote_id = (

                        "QT-"
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
                            rfq.get(
                                "customer",
                                ""
                            ),

                        "product":
                            rfq.get(
                                "title",
                                ""
                            ),

                        "drawing_version":
                            (
                                drawing_version.strip()
                                if drawing_version.strip()
                                else "待確認"
                            ),

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
                            (
                                delivery.strip()
                                if delivery.strip()
                                else "待確認"
                            ),

                        "lead_time":
                            (
                                lead_time.strip()
                                if lead_time.strip()
                                else "待確認"
                            ),

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


                    if quote_df is None:

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


                    try:

                        conn.update(

                            worksheet="Quotes",

                            data=quote_df

                        )


                        st.success(

                            f"✅ 報價草稿建立成功："
                            f"{quote_id}"

                        )


                    except Exception as e:

                        st.error(
                            "❌ 報價寫入 Google Sheets 失敗"
                        )

                        st.exception(e)


                # =================================================
                # 報價摘要
                # =================================================

                st.divider()


                st.markdown(
                    "### 📄 報價摘要"
                )


                summary_df = pd.DataFrame({

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

                        rfq.get(
                            "RFQ_ID",
                            ""
                        ),

                        rfq.get(
                            "customer",
                            ""
                        ),

                        rfq.get(
                            "title",
                            ""
                        ),

                        quantity,

                        currency,

                        money(
                            total_cost
                        ),

                        money(
                            management_cost
                        ),

                        money(
                            profit_amount
                        ),

                        money(
                            discount
                        ),

                        money(
                            quote_total
                        ),

                        money(
                            tax_amount
                        ),

                        money(
                            grand_total
                        ),

                        money(
                            quote_unit_price
                        ),

                        lead_time
                        if lead_time.strip()
                        else "待確認",

                        validity

                    ]

                })


                st.dataframe(

                    summary_df,

                    use_container_width=True,

                    hide_index=True

                )


                # =================================================
                # Excel
                # =================================================

                export_df = pd.DataFrame(
                    [quote_data]
                ) if "quote_data" in locals() else pd.DataFrame()


                if not export_df.empty:

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

                            "📊 下載 Excel 報價",

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


# ################################################################
# TAB 4：STATUS LOG
# ################################################################

with tab_log:

    st.subheader(
        "📜 RFQ 狀態異動紀錄"
    )


    try:

        log_df = conn.read(

            worksheet="StatusLog",

            ttl=0

        )

    except Exception:

        log_df = pd.DataFrame()


    if log_df is None:

        log_df = pd.DataFrame()


    if log_df.empty:

        st.info(
            "目前沒有狀態紀錄"
        )

    else:

        if "change_time" in log_df.columns:

            log_df = log_df.sort_values(

                by="change_time",

                ascending=False

            )


        st.dataframe(

            log_df,

            use_container_width=True,

            hide_index=True

        )
