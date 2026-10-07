import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime, timedelta
import uuid
import io
import os

# Excel 匯出為選用功能；即使環境尚未安裝 openpyxl，主系統仍可正常執行
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    from openpyxl.drawing.image import Image as XLImage
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# =========================================================
# 1. 基本設定
# =========================================================

st.set_page_config(
    page_title="企業版 RFQ 詢價追蹤系統",
    page_icon="📌",
    layout="wide"
)

# =========================================================
# 系統標題 + JENN-WEI LOGO
# =========================================================

logo_path = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "jenn_wei_logo.png"
)

header_logo, header_title = st.columns(
    [1, 9],
    vertical_alignment="center"
)

with header_logo:
    if os.path.exists(logo_path):
        st.image(logo_path, width=85)
    else:
        st.warning("Logo 找不到")

with header_title:
    st.markdown(
        """
        <h1 style="
            margin-top:0;
            margin-bottom:2px;
            padding-top:0;
        ">
            企業版：RFQ 詢價追蹤管理系統
        </h1>
        """,
        unsafe_allow_html=True
    )
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
# TAB 3：報價產生器（簡化正式版）
# ################################################################

with tab_quote:

    st.subheader("💰 RFQ 報價產生器")
    st.caption("簡化流程：RFQ → 商務條件 → 報價品項 → 規格 → 備註 → 正式報價")

    if "quote_items" not in st.session_state:
        st.session_state.quote_items = [
            {"品名": "", "規格/說明": "", "數量": 1.0, "單位": "台", "單價": 0.0}
        ]

    if "generated_quote" not in st.session_state:
        st.session_state.generated_quote = None

    # 常用選單：可自行增修
    VALIDITY_OPTIONS = ["一個月", "30天", "60天", "90天", "自訂"]
    PAYMENT_OPTIONS = [
        "訂金30%現金，餘款70%出貨前付清",
        "訂金50%，餘款50%出貨前付清",
        "月結30天",
        "月結60天",
        "自訂",
    ]
    TRADE_OPTIONS = ["FOR TAIWAN", "EXW", "FOB", "CIF", "自訂"]
    LEAD_OPTIONS = [
        "收到訂金後約30個工作天交貨",
        "收到訂金後約45個工作天交貨",
        "收到訂金後約60個工作天交貨",
        "自訂",
    ]
    UNIT_OPTIONS = ["台", "組", "套", "件", "PCS", "SET"]

    PRODUCT_SPECS = {
        "不顯示規格": [],
        "JCP 90 鐵屑擠壓機": [
            "壓縮能力：100噸",
            "最大使用壓力：190 kg/cm²",
            "產出壓縮塊：Ø90 mm × 70L",
            "一次循環時間：約30秒（不含入料時間）",
            "零件保固一年（消耗品除外）",
        ],
        "KLFM-200 自動鐵屑抬舉機": [
            "抬升高度：1200 mm",
            "馬力：1/2 HP",
            "減速比：200:1",
            "PLC 自動控制上下料",
            "具行程保護及安全門鎖保護",
        ],
        "自訂規格": [],
    }

    if df.empty:
        st.info("目前沒有 RFQ，請先建立 RFQ。")
    else:
        rfq_options = df["RFQ_ID"].dropna().astype(str).tolist()

        if not rfq_options:
            st.info("目前沒有有效 RFQ_ID。")
        else:
            # ---------------------------------------------------------
            # 1. RFQ
            # ---------------------------------------------------------
            st.markdown("### 1️⃣ 選擇 RFQ")

            selected_rfq = st.selectbox(
                "RFQ",
                rfq_options,
                key="quote_rfq_select",
            )

            rfq_matches = df[df["RFQ_ID"].astype(str) == str(selected_rfq)]

            if rfq_matches.empty:
                st.error("❌ 找不到選擇的 RFQ")
            else:
                rfq = rfq_matches.iloc[0]
                rfq_customer = str(rfq.get("customer", ""))
                rfq_product = str(rfq.get("title", ""))

                c1, c2, c3 = st.columns(3)
                c1.metric("RFQ", selected_rfq)
                c2.metric("客戶", rfq_customer)
                c3.metric("詢價名稱", rfq_product)

                # -----------------------------------------------------
                # 2. 商務條件
                # -----------------------------------------------------
                st.divider()
                st.markdown("### 2️⃣ 報價基本資料 / 商務條件")

                c1, c2, c3 = st.columns(3)

                with c1:
                    quote_date = st.date_input(
                        "報價日期",
                        value=datetime.now().date(),
                        key="simple_quote_date",
                    )
                    contact_person = st.text_input(
                        "Attn / 聯絡人",
                        key="simple_contact_person",
                    )
                    currency = st.selectbox(
                        "幣別",
                        ["TWD", "USD", "EUR", "JPY", "CNY"],
                        key="simple_currency",
                    )

                with c2:
                    validity_choice = st.selectbox(
                        "有效期限",
                        VALIDITY_OPTIONS,
                        key="simple_validity_choice",
                    )
                    validity_custom = ""
                    if validity_choice == "自訂":
                        validity_custom = st.text_input(
                            "自訂有效期限",
                            key="simple_validity_custom",
                        )

                    payment_choice = st.selectbox(
                        "付款條件",
                        PAYMENT_OPTIONS,
                        key="simple_payment_choice",
                    )
                    payment_custom = ""
                    if payment_choice == "自訂":
                        payment_custom = st.text_input(
                            "自訂付款條件",
                            key="simple_payment_custom",
                        )

                with c3:
                    trade_choice = st.selectbox(
                        "報價條件",
                        TRADE_OPTIONS,
                        key="simple_trade_choice",
                    )
                    trade_custom = ""
                    if trade_choice == "自訂":
                        trade_custom = st.text_input(
                            "自訂報價條件",
                            key="simple_trade_custom",
                        )

                    lead_choice = st.selectbox(
                        "交貨日期",
                        LEAD_OPTIONS,
                        key="simple_lead_choice",
                    )
                    lead_custom = ""
                    if lead_choice == "自訂":
                        lead_custom = st.text_input(
                            "自訂交期",
                            key="simple_lead_custom",
                        )

                validity = validity_custom.strip() if validity_choice == "自訂" else validity_choice
                payment_terms = payment_custom.strip() if payment_choice == "自訂" else payment_choice
                trade_terms = trade_custom.strip() if trade_choice == "自訂" else trade_choice
                lead_time = lead_custom.strip() if lead_choice == "自訂" else lead_choice

                validity = validity or "待確認"
                payment_terms = payment_terms or "待確認"
                trade_terms = trade_terms or "待確認"
                lead_time = lead_time or "待確認"

                # -----------------------------------------------------
                # 3. 多品項
                # -----------------------------------------------------
                st.divider()
                st.markdown("### 3️⃣ 報價內容")
                st.caption("客戶版只顯示品名、規格/說明、數量、單位、單價與金額；成本與利潤不輸出。")

                # 第一次選 RFQ 時，把詢價名稱帶入第一列
                if (
                    len(st.session_state.quote_items) == 1
                    and not st.session_state.quote_items[0].get("品名", "").strip()
                ):
                    st.session_state.quote_items[0]["品名"] = rfq_product

                item_df = pd.DataFrame(st.session_state.quote_items)

                edited_items = st.data_editor(
                    item_df,
                    num_rows="dynamic",
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "品名": st.column_config.TextColumn("品名", required=True),
                        "規格/說明": st.column_config.TextColumn("規格/說明"),
                        "數量": st.column_config.NumberColumn("數量", min_value=0.0, step=1.0),
                        "單位": st.column_config.SelectboxColumn("單位", options=UNIT_OPTIONS),
                        "單價": st.column_config.NumberColumn("單價", min_value=0.0, step=100.0, format="%.2f"),
                    },
                    key="quote_items_editor",
                )

                # 正規化並計算
                for col in ["品名", "規格/說明", "單位"]:
                    if col not in edited_items.columns:
                        edited_items[col] = ""
                for col in ["數量", "單價"]:
                    if col not in edited_items.columns:
                        edited_items[col] = 0.0

                edited_items["數量"] = pd.to_numeric(edited_items["數量"], errors="coerce").fillna(0.0)
                edited_items["單價"] = pd.to_numeric(edited_items["單價"], errors="coerce").fillna(0.0)
                edited_items["金額"] = edited_items["數量"] * edited_items["單價"]

                st.session_state.quote_items = edited_items[
                    ["品名", "規格/說明", "數量", "單位", "單價"]
                ].to_dict("records")

                subtotal = float(edited_items["金額"].sum())

                tax_mode = st.radio(
                    "營業稅",
                    ["未稅報價（不含5%營業稅）", "加計5%營業稅"],
                    horizontal=True,
                    key="simple_tax_mode",
                )
                tax_rate = 5.0
                tax_amount = subtotal * tax_rate / 100 if tax_mode == "加計5%營業稅" else 0.0
                grand_total = subtotal + tax_amount

                m1, m2, m3 = st.columns(3)
                m1.metric("未稅合計", f"{currency} {money(subtotal)}")
                m2.metric("營業稅", f"{currency} {money(tax_amount)}")
                m3.metric("報價總額", f"{currency} {money(grand_total)}")

                # -----------------------------------------------------
                # 4. 規格
                # -----------------------------------------------------
                st.divider()
                st.markdown("### 4️⃣ 產品規格")

                spec_product = st.selectbox(
                    "規格範本",
                    list(PRODUCT_SPECS.keys()),
                    key="simple_spec_product",
                )

                default_spec_text = "\n".join(PRODUCT_SPECS.get(spec_product, []))

                if spec_product == "不顯示規格":
                    specification = ""
                    st.caption("此報價單不顯示設備規格。")
                else:
                    specification = st.text_area(
                        "規格內容（可直接修改）",
                        value=default_spec_text,
                        height=160,
                        key=f"simple_spec_text_{spec_product}",
                    )

                # -----------------------------------------------------
                # 5. 備註
                # -----------------------------------------------------
                st.divider()
                st.markdown("### 5️⃣ 備註")

                b1, b2, b3 = st.columns(3)
                with b1:
                    note_tax = st.checkbox(
                        "以上報價不含5%營業稅",
                        value=(tax_mode == "未稅報價（不含5%營業稅）"),
                        key="note_tax",
                    )
                    note_pack = st.checkbox("以上報價不含包裝", value=True, key="note_pack")
                with b2:
                    note_shipping = st.checkbox("運費另計", key="note_shipping")
                    note_install = st.checkbox("安裝費另計", key="note_install")
                with b3:
                    note_test = st.checkbox("試車費另計", key="note_test")
                    note_other = st.checkbox("其他備註", key="note_other")

                other_note = ""
                if note_other:
                    other_note = st.text_area("其他備註內容", key="simple_other_note")

                notes = []
                if note_tax:
                    notes.append("以上報價不含5%營業稅")
                if note_pack:
                    notes.append("以上報價不含包裝")
                if note_shipping:
                    notes.append("運費另計")
                if note_install:
                    notes.append("安裝費另計")
                if note_test:
                    notes.append("試車費另計")
                if other_note.strip():
                    notes.append(other_note.strip())

                # -----------------------------------------------------
                # 生成正式報價
                # -----------------------------------------------------
                st.divider()

                if st.button(
                    "🚀 生成正式報價單",
                    type="primary",
                    use_container_width=True,
                    key="generate_simple_quote",
                ):
                    valid_items = edited_items[
                        edited_items["品名"].fillna("").astype(str).str.strip() != ""
                    ].copy()

                    if valid_items.empty:
                        st.error("❌ 請至少輸入一個報價品項。")
                    elif (valid_items["數量"] <= 0).any():
                        st.error("❌ 報價品項數量必須大於 0。")
                    else:
                        quote_id = (
                            "Q-"
                            + datetime.now().strftime("%Y%m%d")
                            + str(uuid.uuid4())[:2].upper()
                        )

                        generated_quote = {
                            "quote_id": quote_id,
                            "RFQ_ID": selected_rfq,
                            "quote_version": "V1",
                            "quote_date": quote_date.strftime("%Y/%m/%d"),
                            "customer": rfq_customer,
                            "contact_person": contact_person.strip() or "待確認",
                            "currency": currency,
                            "validity": validity,
                            "payment_terms": payment_terms,
                            "incoterms": trade_terms,
                            "lead_time": lead_time,
                            "specification": specification.strip(),
                            "notes": notes,
                            "items": valid_items.to_dict("records"),
                            "quote_total": subtotal,
                            "tax_rate": tax_rate,
                            "tax_amount": tax_amount,
                            "grand_total": grand_total,
                            "created_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        }

                        st.session_state.generated_quote = generated_quote

                        # Quotes 主表保留一筆報價摘要，明細以文字保存，避免破壞既有 Quotes 結構
                        quote_row = {
                            "quote_id": quote_id,
                            "RFQ_ID": selected_rfq,
                            "quote_version": "V1",
                            "quote_date": quote_date.strftime("%Y-%m-%d"),
                            "customer": rfq_customer,
                            "product": " / ".join(valid_items["品名"].astype(str).tolist()),
                            "drawing_version": "待確認",
                            "quantity": float(valid_items["數量"].sum()),
                            "currency": currency,
                            "unit": "多品項" if len(valid_items) > 1 else str(valid_items.iloc[0]["單位"]),
                            "material_cost": 0.0,
                            "processing_cost": 0.0,
                            "outsourcing_cost": 0.0,
                            "surface_cost": 0.0,
                            "packing_cost": 0.0,
                            "transport_cost": 0.0,
                            "other_cost": 0.0,
                            "total_cost": 0.0,
                            "management_rate": 0.0,
                            "management_cost": 0.0,
                            "profit_rate": 0.0,
                            "profit_amount": 0.0,
                            "discount": 0.0,
                            "quote_unit_price": 0.0,
                            "quote_total": subtotal,
                            "tax_rate": tax_rate if tax_amount > 0 else 0.0,
                            "tax_amount": tax_amount,
                            "grand_total": grand_total,
                            "payment_terms": payment_terms,
                            "incoterms": trade_terms,
                            "delivery": "待確認",
                            "lead_time": lead_time,
                            "validity": validity,
                            "remark": "；".join(notes),
                            "approved": False,
                            "approved_time": "",
                            "sent": False,
                            "sent_time": "",
                            "created_time": generated_quote["created_time"],
                        }

                        try:
                            quote_df = conn.read(worksheet="Quotes", ttl=0)
                        except Exception:
                            quote_df = pd.DataFrame()

                        if quote_df is None or quote_df.empty:
                            quote_df = pd.DataFrame(columns=QUOTE_COLUMNS)

                        for col in QUOTE_COLUMNS:
                            if col not in quote_df.columns:
                                quote_df[col] = ""

                        quote_df = pd.concat(
                            [quote_df, pd.DataFrame([quote_row])],
                            ignore_index=True,
                        )

                        try:
                            conn.update(worksheet="Quotes", data=quote_df)
                            st.success(f"✅ 正式報價單已建立：{quote_id}")
                        except Exception as e:
                            st.warning("⚠️ 報價單已在畫面產生，但 Quotes 工作表寫入失敗。")
                            st.exception(e)

                # -----------------------------------------------------
                # 正式預覽 / Excel
                # -----------------------------------------------------
                generated = st.session_state.generated_quote

                if generated:
                    st.divider()
                    st.markdown("## 📄 報價單預覽")
                    st.markdown("### 震唯機械股份有限公司")
                    st.caption("報 價 單 / QUOTATION")

                    h1, h2 = st.columns(2)
                    with h1:
                        st.write(f"**客戶名稱：** {generated['customer']}")
                        st.write(f"**Attn：** {generated['contact_person']}")
                        st.write(f"**有效期限：** {generated['validity']}")
                        st.write(f"**付款條件：** {generated['payment_terms']}")
                        st.write(f"**交貨日期：** {generated['lead_time']}")
                    with h2:
                        st.write(f"**REF. NO.：** {generated['quote_id']}")
                        st.write(f"**日期：** {generated['quote_date']}")
                        st.write(f"**報價條件：** {generated['incoterms']}")
                        st.write(f"**幣別：** {generated['currency']}")

                    preview_items = pd.DataFrame(generated["items"])
                    preview_items.insert(0, "NO.", range(1, len(preview_items) + 1))
                    preview_items["單價"] = preview_items["單價"].apply(money)
                    preview_items["金額"] = preview_items["金額"].apply(money)

                    st.dataframe(
                        preview_items[["NO.", "品名", "規格/說明", "數量", "單位", "單價", "金額"]],
                        use_container_width=True,
                        hide_index=True,
                    )

                    t1, t2, t3 = st.columns(3)
                    t1.metric("未稅合計", f"{generated['currency']} {money(generated['quote_total'])}")
                    t2.metric("營業稅", f"{generated['currency']} {money(generated['tax_amount'])}")
                    t3.metric("總額", f"{generated['currency']} {money(generated['grand_total'])}")

                    if generated["specification"]:
                        st.markdown("### 產品規格")
                        for line in generated["specification"].splitlines():
                            if line.strip():
                                st.write(f"• {line.strip()}")

                    if generated["notes"]:
                        st.markdown("### 備註")
                        for note in generated["notes"]:
                            st.write(f"• {note}")

                    st.markdown("**客戶確認：________________　核准：________　覆核：________　製單：________**")

                    # Excel：正式客戶版，不輸出內部成本/利潤
                    excel_buffer = io.BytesIO()
                    try:
                        if not OPENPYXL_AVAILABLE:
                            raise ImportError(
                                "尚未安裝 openpyxl。請在 requirements.txt 加入 openpyxl 後重新部署。"
                            )

                        wb = Workbook()
                        ws = wb.active
                        ws.title = "Quotation"

                        # =====================================================
                        # 公司 LOGO（jenn_wei_logo.png 與 aa.py 放同一層）
                        # =====================================================
                        logo_path = os.path.join(
                            os.path.dirname(os.path.abspath(__file__)),
                            "jenn_wei_logo.png",
                        )

                        if os.path.exists(logo_path):
                            try:
                                logo = XLImage(logo_path)
                                # 保持接近原始 Logo 比例，避免變形
                                logo.width = 115
                                logo.height = 98
                                ws.add_image(logo, "A1")
                            except Exception as logo_error:
                                # Logo 載入失敗不影響報價單本身產生
                                print(f"Logo 載入失敗：{logo_error}")

                        # =====================================================
                        # 公司抬頭
                        # =====================================================
                        # 公司表頭：B:G 合併並水平/垂直置中
                        company_header = [
                            ("B1:G1", "B1", "震唯機械股份有限公司", 18, True),
                            ("B2:G2", "B2", "台中市烏日區溪壩里溪南路一段680巷239號", 10, False),
                            ("B3:G3", "B3", "TEL:886-4-23352368   FAX:886-4-23353880", 10, False),
                            ("B4:G4", "B4", "E-mail: L3352368@ms49.hinet.net", 10, False),
                            ("B5:G5", "B5", "Website: www.jennjwei.com", 10, False),
                        ]

                        for merge_range, cell_ref, text_value, font_size, is_bold in company_header:
                            ws.merge_cells(merge_range)
                            cell = ws[cell_ref]
                            cell.value = text_value
                            cell.font = Font(size=font_size, bold=is_bold)
                            cell.alignment = Alignment(
                                horizontal="center",
                                vertical="center"
                            )

                        ws.row_dimensions[1].height = 26
                        ws.row_dimensions[2].height = 18
                        ws.row_dimensions[3].height = 18
                        ws.row_dimensions[4].height = 18
                        ws.row_dimensions[5].height = 18

                        ws.merge_cells("A7:G7")
                        ws["A7"] = "報  價  單"
                        ws["A7"].font = Font(size=18, bold=True)
                        ws["A7"].alignment = Alignment(horizontal="center", vertical="center")
                        ws.row_dimensions[7].height = 28

                        # =====================================================
                        # 客戶 / 商務條件
                        # =====================================================
                        ws["A9"] = "客戶名稱"
                        ws.merge_cells("B9:D9")
                        ws["B9"] = generated["customer"]
                        ws["E9"] = "REF. NO."
                        ws.merge_cells("F9:G9")
                        ws["F9"] = generated["quote_id"]

                        ws["A10"] = "Attn"
                        ws.merge_cells("B10:D10")
                        ws["B10"] = generated["contact_person"]
                        ws["E10"] = "日期"
                        ws.merge_cells("F10:G10")
                        ws["F10"] = generated["quote_date"]

                        ws["A11"] = "有效期限"
                        ws.merge_cells("B11:D11")
                        ws["B11"] = generated["validity"]
                        ws["E11"] = "報價條件"
                        ws.merge_cells("F11:G11")
                        ws["F11"] = generated["incoterms"]

                        ws["A12"] = "付款條件"
                        ws.merge_cells("B12:G12")
                        ws["B12"] = generated["payment_terms"]

                        ws["A13"] = "交貨日期"
                        ws.merge_cells("B13:G13")
                        ws["B13"] = generated["lead_time"]

                        for row_no in range(9, 14):
                            for cell in ws[row_no]:
                                cell.alignment = Alignment(vertical="center", wrap_text=True)

                        # =====================================================
                        # 報價品項
                        # =====================================================
                        headers = ["NO.", "品名", "規格/說明", "數量", "單位", "單價", "金額"]
                        start_row = 15
                        for col_no, header in enumerate(headers, 1):
                            cell = ws.cell(start_row, col_no, header)
                            cell.font = Font(bold=True)
                            cell.alignment = Alignment(horizontal="center", vertical="center")

                        raw_items = pd.DataFrame(generated["items"])
                        for i, (_, item) in enumerate(raw_items.iterrows(), start=1):
                            r = start_row + i
                            ws.cell(r, 1, i)
                            ws.cell(r, 2, str(item.get("品名", "")))
                            ws.cell(r, 3, str(item.get("規格/說明", "")))
                            ws.cell(r, 4, float(item.get("數量", 0)))
                            ws.cell(r, 5, str(item.get("單位", "")))
                            ws.cell(r, 6, float(item.get("單價", 0)))
                            ws.cell(r, 7, float(item.get("金額", 0)))
                            ws.cell(r, 6).number_format = '#,##0.00'
                            ws.cell(r, 7).number_format = '#,##0.00'

                        total_row = start_row + len(raw_items) + 1
                        ws.cell(total_row, 6, "合計")
                        ws.cell(total_row, 6).font = Font(bold=True)
                        ws.cell(total_row, 7, generated["quote_total"])
                        ws.cell(total_row, 7).font = Font(bold=True)
                        ws.cell(total_row, 7).number_format = '#,##0.00'

                        # 若選擇加稅，額外列出稅額與含稅總額
                        if float(generated.get("tax_amount", 0)) > 0:
                            tax_row = total_row + 1
                            grand_row = total_row + 2
                            ws.cell(tax_row, 6, "營業稅")
                            ws.cell(tax_row, 7, generated["tax_amount"])
                            ws.cell(tax_row, 7).number_format = '#,##0.00'
                            ws.cell(grand_row, 6, "含稅總額")
                            ws.cell(grand_row, 6).font = Font(bold=True)
                            ws.cell(grand_row, 7, generated["grand_total"])
                            ws.cell(grand_row, 7).font = Font(bold=True)
                            ws.cell(grand_row, 7).number_format = '#,##0.00'
                            table_end_row = grand_row
                        else:
                            table_end_row = total_row

                        current_row = table_end_row + 2
                        if generated["specification"]:
                            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
                            ws.cell(current_row, 1, "產品規格")
                            ws.cell(current_row, 1).font = Font(bold=True, underline="single")
                            current_row += 1
                            for line in generated["specification"].splitlines():
                                if line.strip():
                                    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
                                    ws.cell(current_row, 1, f"• {line.strip()}")
                                    ws.cell(current_row, 1).alignment = Alignment(wrap_text=True, vertical="top")
                                    current_row += 1

                        current_row += 1
                        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
                        ws.cell(current_row, 1, "備註")
                        ws.cell(current_row, 1).font = Font(bold=True, underline="single")
                        current_row += 1
                        for note in generated["notes"]:
                            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
                            ws.cell(current_row, 1, f"• {note}")
                            ws.cell(current_row, 1).alignment = Alignment(wrap_text=True)
                            current_row += 1

                        # =====================================================
                        # 簽核區
                        # =====================================================
                        current_row += 2
                        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
                        ws.cell(current_row, 1, "客戶確認：________________")
                        ws.merge_cells(start_row=current_row, start_column=4, end_row=current_row, end_column=7)
                        ws.cell(current_row, 4, "核准：________   覆核：________   製單：________")

                        # =====================================================
                        # 表格框線與欄寬
                        # =====================================================
                        thin = Side(style="thin")
                        for row in ws.iter_rows(min_row=start_row, max_row=table_end_row, min_col=1, max_col=7):
                            for cell in row:
                                cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
                                cell.alignment = Alignment(vertical="center", wrap_text=True)

                        widths = {"A": 7, "B": 25, "C": 28, "D": 10, "E": 10, "F": 16, "G": 16}
                        for col_letter, width in widths.items():
                            ws.column_dimensions[col_letter].width = width

                        # A4 列印設定
                        ws.page_setup.paperSize = ws.PAPERSIZE_A4
                        ws.page_setup.orientation = "portrait"
                        ws.page_setup.fitToWidth = 1
                        ws.page_setup.fitToHeight = 0
                        ws.sheet_properties.pageSetUpPr.fitToPage = True
                        ws.print_options.horizontalCentered = True
                        ws.page_margins.left = 0.25
                        ws.page_margins.right = 0.25
                        ws.page_margins.top = 0.35
                        ws.page_margins.bottom = 0.35
                        ws.print_area = f"A1:G{current_row}"

                        wb.save(excel_buffer)

                        st.download_button(
                            "📊 下載正式 Excel 報價單",
                            data=excel_buffer.getvalue(),
                            file_name=f"{generated['quote_id']}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key="download_simple_quote_excel",
                        )
                    except Exception as e:
                        st.warning(f"Excel 報價單產生失敗：{e}")

                    # =====================================================
                    # PDF：正式客戶版
                    # =====================================================
                    pdf_buffer = io.BytesIO()

                    try:
                        from reportlab.lib import colors
                        from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
                        from reportlab.lib.pagesizes import A4
                        from reportlab.lib.styles import ParagraphStyle
                        from reportlab.lib.units import mm
                        from reportlab.pdfbase import pdfmetrics
                        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
                        from reportlab.platypus import (
                            SimpleDocTemplate,
                            Paragraph,
                            Spacer,
                            Table,
                            TableStyle,
                            Image as RLImage,
                            KeepTogether,
                        )

                        # 繁體中文 CID Font，不需額外上傳字型檔
                        pdfmetrics.registerFont(UnicodeCIDFont("MSung-Light"))
                        PDF_FONT = "MSung-Light"

                        doc = SimpleDocTemplate(
                            pdf_buffer,
                            pagesize=A4,
                            rightMargin=10 * mm,
                            leftMargin=10 * mm,
                            topMargin=9 * mm,
                            bottomMargin=10 * mm,
                            title=f"Quotation {generated['quote_id']}",
                            author="Jenn-Wei Machinery Co., Ltd.",
                        )

                        styles = {
                            "company": ParagraphStyle(
                                "company",
                                fontName=PDF_FONT,
                                fontSize=17,
                                leading=21,
                                alignment=TA_CENTER,
                                spaceAfter=2,
                            ),
                            "company_info": ParagraphStyle(
                                "company_info",
                                fontName=PDF_FONT,
                                fontSize=9,
                                leading=12,
                                alignment=TA_CENTER,
                            ),
                            "title": ParagraphStyle(
                                "title",
                                fontName=PDF_FONT,
                                fontSize=18,
                                leading=23,
                                alignment=TA_CENTER,
                                spaceBefore=5,
                                spaceAfter=8,
                            ),
                            "normal": ParagraphStyle(
                                "normal",
                                fontName=PDF_FONT,
                                fontSize=9.5,
                                leading=13,
                                alignment=TA_LEFT,
                            ),
                            "small": ParagraphStyle(
                                "small",
                                fontName=PDF_FONT,
                                fontSize=8.5,
                                leading=11,
                                alignment=TA_LEFT,
                            ),
                            "section": ParagraphStyle(
                                "section",
                                fontName=PDF_FONT,
                                fontSize=10,
                                leading=14,
                                alignment=TA_LEFT,
                                spaceBefore=3,
                                spaceAfter=3,
                            ),
                            "right": ParagraphStyle(
                                "right",
                                fontName=PDF_FONT,
                                fontSize=9.5,
                                leading=13,
                                alignment=TA_RIGHT,
                            ),
                            "center": ParagraphStyle(
                                "center",
                                fontName=PDF_FONT,
                                fontSize=9.5,
                                leading=13,
                                alignment=TA_CENTER,
                            ),
                        }

                        story = []

                        # -------------------------------------------------
                        # PDF 公司表頭 + Logo
                        # -------------------------------------------------
                        pdf_logo_path = os.path.join(
                            os.path.dirname(os.path.abspath(__file__)),
                            "jenn_wei_logo.png",
                        )

                        logo_flowable = ""
                        if os.path.exists(pdf_logo_path):
                            try:
                                logo_flowable = RLImage(
                                    pdf_logo_path,
                                    width=29 * mm,
                                    height=25 * mm,
                                )
                            except Exception:
                                logo_flowable = ""

                        company_info = [
                            Paragraph("震唯機械股份有限公司", styles["company"]),
                            Paragraph(
                                "台中市烏日區溪壩里溪南路一段680巷239號",
                                styles["company_info"],
                            ),
                            Paragraph(
                                "TEL: 886-4-23352368　FAX: 886-4-23353880",
                                styles["company_info"],
                            ),
                            Paragraph(
                                "E-mail: L3352368@ms49.hinet.net",
                                styles["company_info"],
                            ),
                            Paragraph(
                                "Website: www.jennjwei.com",
                                styles["company_info"],
                            ),
                        ]

                        # 三欄對稱配置：
                        # 左側 Logo 34mm + 中央公司資料 108mm + 右側留白 34mm
                        # 這樣公司資料是以整張報價內容區為基準真正置中，
                        # 不會因左側 Logo 而向右偏移。
                        header_table = Table(
                            [[logo_flowable, company_info, ""]],
                            colWidths=[34 * mm, 108 * mm, 34 * mm],
                        )
                        header_table.setStyle(
                            TableStyle([
                                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                                ("ALIGN", (0, 0), (0, 0), "CENTER"),
                                ("ALIGN", (1, 0), (1, 0), "CENTER"),
                                ("ALIGN", (2, 0), (2, 0), "CENTER"),
                                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                ("TOPPADDING", (0, 0), (-1, -1), 0),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                            ])
                        )
                        story.append(header_table)
                        story.append(Spacer(1, 3 * mm))
                        story.append(Paragraph("報　價　單", styles["title"]))

                        # -------------------------------------------------
                        # 客戶 / 商務條件
                        # -------------------------------------------------
                        info_data = [
                            [
                                Paragraph("客戶名稱", styles["normal"]),
                                Paragraph(str(generated["customer"]), styles["normal"]),
                                Paragraph("REF. NO.", styles["normal"]),
                                Paragraph(str(generated["quote_id"]), styles["normal"]),
                            ],
                            [
                                Paragraph("Attn", styles["normal"]),
                                Paragraph(str(generated["contact_person"]), styles["normal"]),
                                Paragraph("日期", styles["normal"]),
                                Paragraph(str(generated["quote_date"]), styles["normal"]),
                            ],
                            [
                                Paragraph("有效期限", styles["normal"]),
                                Paragraph(str(generated["validity"]), styles["normal"]),
                                Paragraph("報價條件", styles["normal"]),
                                Paragraph(str(generated["incoterms"]), styles["normal"]),
                            ],
                            [
                                Paragraph("付款條件", styles["normal"]),
                                Paragraph(str(generated["payment_terms"]), styles["normal"]),
                                "",
                                "",
                            ],
                            [
                                Paragraph("交貨日期", styles["normal"]),
                                Paragraph(str(generated["lead_time"]), styles["normal"]),
                                "",
                                "",
                            ],
                        ]

                        info_table = Table(
                            info_data,
                            colWidths=[20 * mm, 77 * mm, 20 * mm, 59 * mm],
                        )
                        info_table.setStyle(
                            TableStyle([
                                ("FONTNAME", (0, 0), (-1, -1), PDF_FONT),
                                ("FONTSIZE", (0, 0), (-1, -1), 9.5),
                                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                                ("SPAN", (1, 3), (3, 3)),
                                ("SPAN", (1, 4), (3, 4)),
                                ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor("#BFBFBF")),
                                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                                ("TOPPADDING", (0, 0), (-1, -1), 3),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                            ])
                        )
                        story.append(info_table)
                        story.append(Spacer(1, 4 * mm))

                        # -------------------------------------------------
                        # 報價品項
                        # -------------------------------------------------
                        pdf_item_rows = [[
                            Paragraph("NO.", styles["center"]),
                            Paragraph("品名", styles["center"]),
                            Paragraph("規格/說明", styles["center"]),
                            Paragraph("數量", styles["center"]),
                            Paragraph("單位", styles["center"]),
                            Paragraph("單價", styles["center"]),
                            Paragraph("金額", styles["center"]),
                        ]]

                        for i, item in enumerate(generated["items"], start=1):
                            qty = to_float(item.get("數量", 0))
                            unit_price = to_float(item.get("單價", 0))
                            amount = to_float(item.get("金額", qty * unit_price))
                            pdf_item_rows.append([
                                Paragraph(str(i), styles["center"]),
                                Paragraph(str(item.get("品名", "")), styles["normal"]),
                                Paragraph(str(item.get("規格/說明", "")), styles["small"]),
                                Paragraph(f"{qty:g}", styles["center"]),
                                Paragraph(str(item.get("單位", "")), styles["center"]),
                                Paragraph(f"{unit_price:,.2f}", styles["right"]),
                                Paragraph(f"{amount:,.2f}", styles["right"]),
                            ])

                        pdf_item_rows.append([
                            "", "", "", "", "",
                            Paragraph("合計", styles["right"]),
                            Paragraph(f"{to_float(generated['quote_total']):,.2f}", styles["right"]),
                        ])

                        if to_float(generated.get("tax_amount", 0)) > 0:
                            pdf_item_rows.append([
                                "", "", "", "", "",
                                Paragraph("營業稅", styles["right"]),
                                Paragraph(f"{to_float(generated['tax_amount']):,.2f}", styles["right"]),
                            ])
                            pdf_item_rows.append([
                                "", "", "", "", "",
                                Paragraph("含稅總額", styles["right"]),
                                Paragraph(f"{to_float(generated['grand_total']):,.2f}", styles["right"]),
                            ])

                        item_table = Table(
                            pdf_item_rows,
                            colWidths=[
                                10 * mm,
                                40 * mm,
                                47 * mm,
                                14 * mm,
                                14 * mm,
                                25 * mm,
                                26 * mm,
                            ],
                            repeatRows=1,
                        )
                        item_table.setStyle(
                            TableStyle([
                                ("FONTNAME", (0, 0), (-1, -1), PDF_FONT),
                                ("FONTSIZE", (0, 0), (-1, -1), 9),
                                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
                                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                                ("TOPPADDING", (0, 0), (-1, -1), 3),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                            ])
                        )
                        story.append(item_table)
                        story.append(Spacer(1, 4 * mm))

                        # -------------------------------------------------
                        # 規格
                        # -------------------------------------------------
                        if generated.get("specification", "").strip():
                            spec_flowables = [
                                Paragraph("<u>產品規格</u>", styles["section"])
                            ]
                            for line in generated["specification"].splitlines():
                                if line.strip():
                                    spec_flowables.append(
                                        Paragraph(
                                            f"• {line.strip()}",
                                            styles["small"],
                                        )
                                    )
                            story.append(KeepTogether(spec_flowables))
                            story.append(Spacer(1, 2 * mm))

                        # -------------------------------------------------
                        # 備註
                        # -------------------------------------------------
                        note_flowables = [
                            Paragraph("<u>備註</u>", styles["section"])
                        ]
                        for note in generated.get("notes", []):
                            note_flowables.append(
                                Paragraph(f"• {str(note)}", styles["small"])
                            )
                        story.append(KeepTogether(note_flowables))
                        story.append(Spacer(1, 7 * mm))

                        # -------------------------------------------------
                        # 簽核
                        # -------------------------------------------------
                        sign_table = Table(
                            [[
                                Paragraph("客戶確認：________________", styles["normal"]),
                                Paragraph(
                                    "核准：________　覆核：________　製單：________",
                                    styles["right"],
                                ),
                            ]],
                            colWidths=[78 * mm, 98 * mm],
                        )
                        sign_table.setStyle(
                            TableStyle([
                                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                            ])
                        )
                        story.append(sign_table)

                        doc.build(story)
                        pdf_buffer.seek(0)

                        st.download_button(
                            "📄 下載正式 PDF 報價單",
                            data=pdf_buffer.getvalue(),
                            file_name=f"{generated['quote_id']}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                            key="download_simple_quote_pdf",
                        )

                    except ImportError:
                        st.warning(
                            "PDF 報價單需要 reportlab。"
                            "請在 requirements.txt 加入 reportlab 後重新部署。"
                        )
                    except Exception as e:
                        st.warning(f"PDF 報價單產生失敗：{e}")


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
