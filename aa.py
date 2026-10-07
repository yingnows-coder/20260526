import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime, timedelta
import uuid
import io
import os
import plotly.express as px
import plotly.graph_objects as go

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
        "📋 雲端 Trello 管理系統",
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
    "first_followup_due",
    "product_category",
    "product_model",
    "sales_result",
    "deal_date",
    "deal_amount",
    "lost_reason",
    "continent",
    "country"
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
    "成交生產",
    "交貨",
    "收款",
    "結案"

]


# =========================================================
# 產品分類 / 型號主檔
# =========================================================
PRODUCT_MODELS = {
    "CNC Multi-Axis Turning Milling Center": ["JMY-4220", "JMY-6020"],
    "CNC Double Spindle Turning Center": ["TMY-25", "TMY-42"],
    "High Precision CNC Lathe": ["JW-42B3", "JW-42B3S", "JW-42B3M"],
    "High Precision CNC Mult Function Lathe": ["JW-25", "JW-32", "JFKL-400"],
    "High Precision Slant Bed CNC Lathe": [
        "JW-16", "JW-16M", "JW-16ST", "JW-16STM",
        "JW-21", "JW-21M", "JW21-ST", "JW-21STM"
    ],
    "Heavy Duty CNC Lathe": [
        "JW-31", "JW-31M", "JW-31ST", "JW-31STM",
        "JW-41", "JW-41M", "JW-41ST", "JW-41STM", "JW-41MLX2"
    ],
    "Chip Compactor": ["JCP", "JCP S"],
    "Crush Machine": ["CCM-600", "CCM-700"],
    "Lifting Machine": ["KLFM-200"],
}
PRODUCT_CATEGORIES = list(PRODUCT_MODELS.keys())


# =========================================================
# 全球洲別 / 國家主檔
# =========================================================
CONTINENT_COUNTRIES = {
    "亞洲": [
        "台灣", "中國", "日本", "韓國", "北韓", "蒙古",
        "香港", "澳門",
        "新加坡", "馬來西亞", "泰國", "越南", "菲律賓", "印尼",
        "汶萊", "柬埔寨", "寮國", "緬甸", "東帝汶",
        "印度", "巴基斯坦", "孟加拉", "斯里蘭卡", "尼泊爾",
        "不丹", "馬爾地夫", "阿富汗",
        "哈薩克", "烏茲別克", "土庫曼", "吉爾吉斯", "塔吉克",
        "亞美尼亞", "亞塞拜然", "喬治亞",
        "土耳其", "以色列", "巴勒斯坦", "約旦", "黎巴嫩", "敘利亞",
        "伊拉克", "伊朗", "沙烏地阿拉伯", "阿拉伯聯合大公國",
        "卡達", "科威特", "巴林", "阿曼", "葉門"
    ],
    "歐洲": [
        "英國", "愛爾蘭", "法國", "德國", "義大利", "西班牙",
        "葡萄牙", "荷蘭", "比利時", "盧森堡", "瑞士", "奧地利",
        "丹麥", "瑞典", "挪威", "芬蘭", "冰島",
        "波蘭", "捷克", "斯洛伐克", "匈牙利", "羅馬尼亞",
        "保加利亞", "希臘", "克羅埃西亞", "斯洛維尼亞",
        "塞爾維亞", "波士尼亞與赫塞哥維納", "蒙特內哥羅",
        "北馬其頓", "阿爾巴尼亞", "科索沃",
        "愛沙尼亞", "拉脫維亞", "立陶宛", "烏克蘭", "白俄羅斯",
        "摩爾多瓦", "俄羅斯",
        "賽普勒斯", "馬爾他", "安道爾", "摩納哥", "列支敦斯登",
        "聖馬利諾", "梵蒂岡"
    ],
    "北美洲": [
        "美國", "加拿大", "墨西哥",
        "瓜地馬拉", "貝里斯", "宏都拉斯", "薩爾瓦多",
        "尼加拉瓜", "哥斯大黎加", "巴拿馬",
        "巴哈馬", "古巴", "牙買加", "海地", "多明尼加",
        "安地卡及巴布達", "巴貝多", "多米尼克", "格瑞那達",
        "聖克里斯多福及尼維斯", "聖露西亞",
        "聖文森及格瑞那丁", "千里達及托巴哥"
    ],
    "南美洲": [
        "巴西", "阿根廷", "智利", "秘魯", "哥倫比亞",
        "委內瑞拉", "厄瓜多", "玻利維亞", "巴拉圭",
        "烏拉圭", "蓋亞那", "蘇利南"
    ],
    "非洲": [
        "南非", "埃及", "摩洛哥", "阿爾及利亞", "突尼西亞",
        "利比亞", "蘇丹", "南蘇丹", "衣索比亞", "厄利垂亞",
        "吉布地", "索馬利亞", "肯亞", "烏干達", "坦尚尼亞",
        "盧安達", "蒲隆地",
        "奈及利亞", "迦納", "象牙海岸", "塞內加爾", "馬利",
        "尼日", "查德", "茅利塔尼亞", "甘比亞", "幾內亞",
        "幾內亞比索", "獅子山", "賴比瑞亞", "布吉納法索",
        "貝南", "多哥", "維德角",
        "喀麥隆", "中非共和國", "赤道幾內亞", "加彭",
        "剛果共和國", "剛果民主共和國", "聖多美普林西比",
        "安哥拉", "尚比亞", "辛巴威", "馬拉威", "莫三比克",
        "納米比亞", "波札那", "賴索托", "史瓦帝尼",
        "馬達加斯加", "模里西斯", "塞席爾", "葛摩"
    ],
    "大洋洲": [
        "澳洲", "紐西蘭", "巴布亞紐幾內亞", "斐濟",
        "索羅門群島", "萬那杜", "薩摩亞", "東加",
        "吉里巴斯", "吐瓦魯", "諾魯", "帛琉",
        "密克羅尼西亞聯邦", "馬紹爾群島"
    ]
}
CONTINENTS = list(CONTINENT_COUNTRIES.keys())


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
        ttl=GSHEETS_READ_TTL
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

# 流程相容：案件進入成交生產 / 交貨 / 收款後，視為已成交
if "sales_result" in df.columns and "status" in df.columns:
    won_workflow_mask = df["status"].astype(str).isin(["成交生產", "交貨", "收款"])
    df.loc[won_workflow_mask, "sales_result"] = "已成交"


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
            ttl=GSHEETS_READ_TTL
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

tab_rfq, tab_board, tab_quote, tab_kpi, tab_gantt, tab_sales, tab_intelligence, tab_log = st.tabs(

    [
        "📋 RFQ追蹤表",
        "📊 RFQ看板",
        "💰 報價產生器",
        "📈 KPI儀表板",
        "📅 甘特圖",
        "📊 銷售分析",
        "🎯 Sales Intelligence",
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


    # 產品分類 / 型號必須放在 st.form 外面，
    # 才能在切換分類時立即重新執行並更新第二層型號選單。
    p1, p2 = st.columns([1.6, 1])

    with p1:
        new_product_category = st.selectbox(
            "🏷️ 產品分類",
            ["請選擇"] + PRODUCT_CATEGORIES,
            key="new_product_category"
        )

    with p2:
        model_options = (
            PRODUCT_MODELS.get(new_product_category, [])
            if new_product_category != "請選擇"
            else []
        )

        new_product_model = st.selectbox(
            "⚙️ 產品型號",
            ["請選擇"] + model_options,
            key="new_product_model"
        )

    g1, g2 = st.columns([1, 1.5])

    with g1:
        new_continent = st.selectbox(
            "🌍 洲別",
            ["請選擇"] + CONTINENTS,
            key="new_continent"
        )

    with g2:
        country_options = (
            CONTINENT_COUNTRIES.get(new_continent, [])
            if new_continent != "請選擇"
            else []
        )
        new_country = st.selectbox(
            "🌐 國家",
            ["請選擇"] + country_options,
            key="new_country"
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
            default_title = (
                new_product_model
                if new_product_model != "請選擇"
                else ""
            )
            new_title = st.text_input(
                "📌 詢價名稱",
                value=default_title,
                key=f"new_title_{new_product_model}"
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
                    ),

                "product_category":
                    (
                        new_product_category
                        if new_product_category != "請選擇"
                        else ""
                    ),

                "product_model":
                    (
                        new_product_model
                        if new_product_model != "請選擇"
                        else ""
                    ),

                "continent":
                    (
                        new_continent
                        if new_continent != "請選擇"
                        else ""
                    ),

                "country":
                    (
                        new_country
                        if new_country != "請選擇"
                        else ""
                    ),

                "sales_result":
                    "進行中",

                "deal_date":
                    "",

                "deal_amount":
                    0.0,

                "lost_reason":
                    ""

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

    # =========================================================
    # RFQ 管理看板：業務下拉篩選
    # =========================================================
    board_owner_source = df.copy()

    board_owner_options = sorted(
        [
            str(x).strip()
            for x in board_owner_source.get("owner", pd.Series(dtype=str))
                .dropna()
                .astype(str)
                .unique()
            if str(x).strip()
        ]
    )

    board_owner_filter = st.selectbox(
        "👤 選擇業務",
        ["全部業務"] + board_owner_options,
        key="rfq_board_owner_filter"
    )

    if board_owner_filter == "全部業務":
        board_df = board_owner_source.copy()
    else:
        board_df = board_owner_source[
            board_owner_source["owner"].fillna("").astype(str).str.strip()
            == board_owner_filter
        ].copy()

    st.caption(
        f"目前顯示：{board_owner_filter}｜共 {len(board_df)} 筆 RFQ"
    )


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

                                ttl=GSHEETS_READ_TTL

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

    # =========================================================
    # 報價單語言版本
    # =========================================================
    quote_language = st.radio(
        "🌐 報價單版本",
        ["中文版", "English Version"],
        horizontal=True,
        key="quote_language_selector",
    )
    quote_is_en = quote_language == "English Version"


    st.subheader("💰 RFQ 報價產生器")
    st.caption("正式報價流程：RFQ → 商務條件 → 報價品項 → 規格 → 備註 → 中／英文正式報價單")

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

    def translate_payment_term_en(value):
        """Convert the app's standard Chinese payment terms to formal English."""
        raw = "" if value is None else str(value).strip()
        compact = raw.replace(" ", "").replace("％", "%").replace("，", ",")
        mapping = {
            "訂金30%現金,餘款70%出貨前付清": "30% cash deposit, 70% balance before shipment",
            "訂金50%,餘款50%出貨前付清": "50% deposit, 50% balance before shipment",
            "訂金30%,出貨前付清70%": "30% deposit, 70% balance before shipment",
            "訂金50%,出貨前付清50%": "50% deposit, 50% balance before shipment",
            "訂金30%,交機前付清70%": "30% deposit, 70% balance before delivery",
            "訂金50%,交機前付清50%": "50% deposit, 50% balance before delivery",
            "月結30天": "Net 30 days",
            "月結60天": "Net 60 days",
            "現金": "Cash",
            "貨到付款": "Payment upon delivery",
            "出貨前付清": "Full payment before shipment",
            "交機前付清": "Full payment before delivery",
            "待確認": "To be confirmed",
        }
        return mapping.get(compact, raw)

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
                        "Payment Terms" if quote_is_en else "付款條件",
                        PAYMENT_OPTIONS,
                        key="simple_payment_choice",
                        format_func=lambda x: (
                            "Custom" if x == "自訂"
                            else translate_payment_term_en(x)
                        ) if quote_is_en else x,
                    )
                    payment_custom = ""
                    if payment_choice == "自訂":
                        payment_custom = st.text_input(
                            "Custom Payment Terms" if quote_is_en else "自訂付款條件",
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

                    delivery_days = st.number_input(
                        "📦 收到訂金後交貨天數",
                        min_value=1,
                        max_value=999,
                        value=30,
                        step=1,
                        key="quote_delivery_days"
                    )
                    lead_choice = (
                        f"Within {int(delivery_days)} days after receipt of deposit"
                        if quote_is_en
                        else f"收到訂金後 {int(delivery_days)} 天"
                    )
                    lead_custom = ""
                validity = validity_custom.strip() if validity_choice == "自訂" else validity_choice
                payment_terms = payment_custom.strip() if payment_choice == "自訂" else payment_choice
                trade_terms = trade_custom.strip() if trade_choice == "自訂" else trade_choice
                lead_time = lead_custom.strip() if lead_choice == "自訂" else lead_choice

                validity = validity or "待確認"
                payment_terms = payment_terms or "待確認"
                trade_terms = trade_terms or "待確認"
                lead_time = lead_time or "待確認"

                # 英文版標準商務條件
                if quote_is_en:
                    validity_map_en = {
                        "一個月": "One month",
                        "30天": "30 days",
                        "60天": "60 days",
                        "90天": "90 days",
                        "待確認": "To be confirmed",
                    }
                    payment_terms = translate_payment_term_en(payment_terms)
                    trade_terms = "To be confirmed" if trade_terms == "待確認" else trade_terms

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

                if quote_is_en:
                    note_map_en = {
                        "以上報價不含5%營業稅": "The above quotation excludes 5% VAT.",
                        "以上報價不含包裝": "The above quotation excludes packing charges.",
                        "運費另計": "Freight charges are not included.",
                        "安裝費另計": "Installation charges are not included.",
                        "試車費另計": "Commissioning charges are not included.",
                    }
                    notes = [note_map_en.get(note, note) for note in notes]

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
                            "quote_language": quote_language,
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
                            quote_df = conn.read(worksheet="Quotes", ttl=GSHEETS_READ_TTL)
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
                            clear_gsheets_cache()
                            st.success(f"✅ 正式報價單已建立：{quote_id}")
                        except Exception as e:
                            st.warning("⚠️ 報價單已在畫面產生，但 Quotes 工作表寫入失敗。")
                            st.exception(e)

                # -----------------------------------------------------
                # 正式預覽 / Excel
                # -----------------------------------------------------
                generated = st.session_state.generated_quote

                if generated:
                    generated_language = generated.get("quote_language", quote_language)
                    generated_is_en = generated_language == "English Version"

                    # Output-time normalization: old session-state quotes may still contain
                    # Chinese payment terms. Force English before Preview / Excel / PDF.
                    if generated_is_en:
                        generated["payment_terms"] = translate_payment_term_en(
                            generated.get("payment_terms", "")
                        )

                    st.divider()
                    st.markdown("## 📄 " + ("Quotation Preview" if generated_is_en else "報價單預覽"))
                    st.markdown("### " + ("JENN-WEI MACHINERY CO., LTD." if generated_is_en else "震唯機械股份有限公司"))
                    st.caption("QUOTATION" if generated_is_en else "報 價 單")

                    h1, h2 = st.columns(2)
                    with h1:
                        st.write(f"**{'Customer' if generated_is_en else '客戶名稱'}：** {generated['customer']}")
                        st.write(f"**Attn：** {generated['contact_person']}")
                        st.write(f"**{'Validity' if generated_is_en else '有效期限'}：** {generated['validity']}")
                        st.write(f"**{'Payment Terms' if generated_is_en else '付款條件'}：** {generated['payment_terms']}")
                        st.write(f"**{'Delivery' if generated_is_en else '交貨日期'}：** {generated['lead_time']}")
                    with h2:
                        st.write(f"**REF. NO.：** {generated['quote_id']}")
                        st.write(f"**{'Date' if generated_is_en else '日期'}：** {generated['quote_date']}")
                        st.write(f"**{'Trade Terms' if generated_is_en else '報價條件'}：** {generated['incoterms']}")
                        st.write(f"**{'Currency' if generated_is_en else '幣別'}：** {generated['currency']}")

                    preview_items = pd.DataFrame(generated["items"])
                    preview_items.insert(0, "NO.", range(1, len(preview_items) + 1))
                    preview_items["單價"] = preview_items["單價"].apply(money)
                    preview_items["金額"] = preview_items["金額"].apply(money)

                    preview_display = preview_items[["NO.", "品名", "規格/說明", "數量", "單位", "單價", "金額"]].copy()
                    if generated_is_en:
                        preview_display.columns = ["NO.", "Description", "Specification / Description", "Qty", "Unit", "Unit Price", "Amount"]

                    st.dataframe(
                        preview_display,
                        use_container_width=True,
                        hide_index=True,
                    )

                    t1, t2, t3 = st.columns(3)
                    t1.metric("Subtotal" if generated_is_en else "未稅合計", f"{generated['currency']} {money(generated['quote_total'])}")
                    t2.metric("VAT" if generated_is_en else "營業稅", f"{generated['currency']} {money(generated['tax_amount'])}")
                    t3.metric("Total" if generated_is_en else "總額", f"{generated['currency']} {money(generated['grand_total'])}")

                    if generated["specification"]:
                        st.markdown("### Product Specifications" if generated_is_en else "### 產品規格")
                        for line in generated["specification"].splitlines():
                            if line.strip():
                                st.write(f"• {line.strip()}")

                    if generated["notes"]:
                        st.markdown("### Remarks" if generated_is_en else "### 備註")
                        for note in generated["notes"]:
                            st.write(f"• {note}")

                    st.markdown("**Customer Acceptance: ________________　Approved by: ________　Checked by: ________　Prepared by: ________**" if generated_is_en else "**客戶確認：________________　核准：________　覆核：________　製單：________**")

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
                            ("B1:G1", "B1", "JENN-WEI MACHINERY CO., LTD." if generated_is_en else "震唯機械股份有限公司", 18, True),
                            ("B2:G2", "B2", "No. 239, Ln. 680, Sec. 1, Xinan Rd., Wuri Dist., Taichung City, Taiwan" if generated_is_en else "台中市烏日區溪壩里溪南路一段680巷239號", 10, False),
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
                        ws["A7"] = "QUOTATION" if generated_is_en else "報  價  單"
                        ws["A7"].font = Font(size=18, bold=True)
                        ws["A7"].alignment = Alignment(horizontal="center", vertical="center")
                        ws.row_dimensions[7].height = 28

                        # =====================================================
                        # 客戶 / 商務條件
                        # =====================================================
                        ws["A9"] = "Customer" if generated_is_en else "客戶名稱"
                        ws.merge_cells("B9:D9")
                        ws["B9"] = generated["customer"]
                        ws["E9"] = "REF. NO."
                        ws.merge_cells("F9:G9")
                        ws["F9"] = generated["quote_id"]

                        ws["A10"] = "Attn"
                        ws.merge_cells("B10:D10")
                        ws["B10"] = generated["contact_person"]
                        ws["E10"] = "Date" if generated_is_en else "日期"
                        ws.merge_cells("F10:G10")
                        ws["F10"] = generated["quote_date"]

                        ws["A11"] = "Validity" if generated_is_en else "有效期限"
                        ws.merge_cells("B11:D11")
                        ws["B11"] = generated["validity"]
                        ws["E11"] = "Trade Terms" if generated_is_en else "報價條件"
                        ws.merge_cells("F11:G11")
                        ws["F11"] = generated["incoterms"]

                        ws["A12"] = "Payment Terms" if generated_is_en else "付款條件"
                        ws.merge_cells("B12:G12")
                        ws["B12"] = generated["payment_terms"]

                        ws["A13"] = "Delivery" if generated_is_en else "交貨日期"
                        ws.merge_cells("B13:G13")
                        ws["B13"] = generated["lead_time"]

                        for row_no in range(9, 14):
                            for cell in ws[row_no]:
                                cell.alignment = Alignment(vertical="center", wrap_text=True)

                        # =====================================================
                        # 報價品項
                        # =====================================================
                        headers = (["NO.", "Description", "Specification / Description", "Qty", "Unit", "Unit Price", "Amount"] if generated_is_en else ["NO.", "品名", "規格/說明", "數量", "單位", "單價", "金額"])
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
                        ws.cell(total_row, 6, "Subtotal" if generated_is_en else "合計")
                        ws.cell(total_row, 6).font = Font(bold=True)
                        ws.cell(total_row, 7, generated["quote_total"])
                        ws.cell(total_row, 7).font = Font(bold=True)
                        ws.cell(total_row, 7).number_format = '#,##0.00'

                        # 若選擇加稅，額外列出稅額與含稅總額
                        if float(generated.get("tax_amount", 0)) > 0:
                            tax_row = total_row + 1
                            grand_row = total_row + 2
                            ws.cell(tax_row, 6, "VAT" if generated_is_en else "營業稅")
                            ws.cell(tax_row, 7, generated["tax_amount"])
                            ws.cell(tax_row, 7).number_format = '#,##0.00'
                            ws.cell(grand_row, 6, "Total" if generated_is_en else "含稅總額")
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
                            ws.cell(current_row, 1, "Product Specifications" if generated_is_en else "產品規格")
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
                        ws.cell(current_row, 1, "Remarks" if generated_is_en else "備註")
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
                        ws.cell(current_row, 1, "Customer Acceptance: ________________" if generated_is_en else "客戶確認：________________")
                        ws.merge_cells(start_row=current_row, start_column=4, end_row=current_row, end_column=7)
                        ws.cell(current_row, 4, "Approved by: ________   Checked by: ________   Prepared by: ________" if generated_is_en else "核准：________   覆核：________   製單：________")

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
                            ("📊 Download Excel Quotation" if generated_is_en else "📊 下載正式 Excel 報價單"),
                            data=excel_buffer.getvalue(),
                            file_name=f"{generated['quote_id']}_{'EN' if generated_is_en else 'CN'}.xlsx",
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
                            Paragraph("JENN-WEI MACHINERY CO., LTD." if generated_is_en else "震唯機械股份有限公司", styles["company"]),
                            Paragraph(
                                "No. 239, Ln. 680, Sec. 1, Xinan Rd., Wuri Dist., Taichung City, Taiwan" if generated_is_en else "台中市烏日區溪壩里溪南路一段680巷239號",
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
                        story.append(Paragraph("QUOTATION" if generated_is_en else "報　價　單", styles["title"]))

                        # -------------------------------------------------
                        # 客戶 / 商務條件
                        # -------------------------------------------------
                        info_data = [
                            [
                                Paragraph("Customer" if generated_is_en else "客戶名稱", styles["normal"]),
                                Paragraph(str(generated["customer"]), styles["normal"]),
                                Paragraph("REF. NO.", styles["normal"]),
                                Paragraph(str(generated["quote_id"]), styles["normal"]),
                            ],
                            [
                                Paragraph("Attn", styles["normal"]),
                                Paragraph(str(generated["contact_person"]), styles["normal"]),
                                Paragraph("Date" if generated_is_en else "日期", styles["normal"]),
                                Paragraph(str(generated["quote_date"]), styles["normal"]),
                            ],
                            [
                                Paragraph("Validity" if generated_is_en else "有效期限", styles["normal"]),
                                Paragraph(str(generated["validity"]), styles["normal"]),
                                Paragraph("Trade Terms" if generated_is_en else "報價條件", styles["normal"]),
                                Paragraph(str(generated["incoterms"]), styles["normal"]),
                            ],
                            [
                                Paragraph("Payment Terms" if generated_is_en else "付款條件", styles["normal"]),
                                Paragraph(str(generated["payment_terms"]), styles["normal"]),
                                "",
                                "",
                            ],
                            [
                                Paragraph("Delivery" if generated_is_en else "交貨日期", styles["normal"]),
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
                            Paragraph("Description" if generated_is_en else "品名", styles["center"]),
                            Paragraph("Specification / Description" if generated_is_en else "規格/說明", styles["center"]),
                            Paragraph("Qty" if generated_is_en else "數量", styles["center"]),
                            Paragraph("Unit" if generated_is_en else "單位", styles["center"]),
                            Paragraph("Unit Price" if generated_is_en else "單價", styles["center"]),
                            Paragraph("Amount" if generated_is_en else "金額", styles["center"]),
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
                            Paragraph("Subtotal" if generated_is_en else "合計", styles["right"]),
                            Paragraph(f"{to_float(generated['quote_total']):,.2f}", styles["right"]),
                        ])

                        if to_float(generated.get("tax_amount", 0)) > 0:
                            pdf_item_rows.append([
                                "", "", "", "", "",
                                Paragraph("VAT" if generated_is_en else "營業稅", styles["right"]),
                                Paragraph(f"{to_float(generated['tax_amount']):,.2f}", styles["right"]),
                            ])
                            pdf_item_rows.append([
                                "", "", "", "", "",
                                Paragraph("Total" if generated_is_en else "含稅總額", styles["right"]),
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
                                Paragraph("<u>Product Specifications</u>" if generated_is_en else "<u>產品規格</u>", styles["section"])
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
                            Paragraph("<u>Remarks</u>" if generated_is_en else "<u>備註</u>", styles["section"])
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
                                Paragraph("Customer Acceptance: ________________" if generated_is_en else "客戶確認：________________", styles["normal"]),
                                Paragraph(
                                    "Approved by: ________   Checked by: ________   Prepared by: ________" if generated_is_en else "核准：________　覆核：________　製單：________",
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
                            ("📄 Download PDF Quotation" if generated_is_en else "📄 下載正式 PDF 報價單"),
                            data=pdf_buffer.getvalue(),
                            file_name=f"{generated['quote_id']}_{'EN' if generated_is_en else 'CN'}.pdf",
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
# TAB 4：KPI 儀表板
# ################################################################

with tab_kpi:
    st.markdown("## 📈 Management Dashboard")
    st.caption("從詢價、報價到成交與結案的即時管理指標，適合主管會議與成品展示。")


    def render_kpi_bubble(dataframe, x_col, y_col, size_col, label_col, title, x_title, y_title):
        """
        力學氣泡圖：
        - 氣泡面積代表 size_col
        - 氣泡彼此碰撞、不重疊
        - 中心引力讓群組自然聚合
        - 可拖曳氣泡，放開後重新進入力學模擬
        """
        if dataframe is None or dataframe.empty:
            st.info(f"{title}：目前沒有可分析資料。")
            return

        plot_df = dataframe.copy()
        for col in [x_col, y_col, size_col]:
            plot_df[col] = pd.to_numeric(plot_df[col], errors="coerce").fillna(0)

        # 限制節點數，避免瀏覽器負擔過重；KPI 通常遠低於此數量。
        plot_df = plot_df.head(80).copy()
        max_size = float(plot_df[size_col].clip(lower=0).max())
        min_size = float(plot_df[size_col].clip(lower=0).min())

        nodes = []
        for _, row in plot_df.iterrows():
            raw = max(float(row[size_col]), 0.0)
            if max_size > min_size:
                radius = 24 + ((raw - min_size) / (max_size - min_size)) ** 0.5 * 42
            else:
                radius = 38

            nodes.append({
                "label": str(row[label_col]),
                "xValue": float(row[x_col]),
                "yValue": float(row[y_col]),
                "sizeValue": raw,
                "radius": round(radius, 2),
            })

        import json
        node_json = json.dumps(nodes, ensure_ascii=False)
        title_json = json.dumps(title, ensure_ascii=False)
        x_title_json = json.dumps(x_title, ensure_ascii=False)
        y_title_json = json.dumps(y_title, ensure_ascii=False)
        size_title_json = json.dumps(size_col, ensure_ascii=False)

        html = f"""
        <div id="force-bubble-root" style="
            width:100%; font-family:Arial,'Microsoft JhengHei',sans-serif;
            color:#222; box-sizing:border-box;">
          <div style="font-size:20px;font-weight:700;margin:4px 0 2px 0;">{title}</div>
          <div style="font-size:13px;color:#cbd5e1;margin-bottom:10px;font-weight:600;">
            氣泡大小＝{size_col}｜可用滑鼠拖曳氣泡，放開後會依力學重新排列
          </div>

          <div style="
              display:grid;
              grid-template-columns:1.15fr 1fr 1fr 1fr;
              gap:1px;
              margin-bottom:10px;
              border:1px solid #d9d9d9;
              border-radius:8px;
              overflow:hidden;
              background:#d9d9d9;
              font-size:12px;">
            <div style="background:#2563eb;color:#ffffff;padding:13px 10px;font-weight:900;font-size:17px;line-height:1.35;text-align:center;display:flex;align-items:center;justify-content:center;min-height:48px;box-sizing:border-box;">
              分析項目
            </div>
            <div style="background:#2563eb;color:#ffffff;padding:13px 10px;font-weight:900;font-size:17px;line-height:1.35;text-align:center;display:flex;align-items:center;justify-content:center;min-height:48px;box-sizing:border-box;">
              X 軸
            </div>
            <div style="background:#2563eb;color:#ffffff;padding:13px 10px;font-weight:900;font-size:17px;line-height:1.35;text-align:center;display:flex;align-items:center;justify-content:center;min-height:48px;box-sizing:border-box;">
              Y 軸
            </div>
            <div style="background:#2563eb;color:#ffffff;padding:13px 10px;font-weight:900;font-size:17px;line-height:1.35;text-align:center;display:flex;align-items:center;justify-content:center;min-height:48px;box-sizing:border-box;">
              氣泡大小
            </div>

            <div style="background:#111827;color:#ffffff;padding:11px 10px;text-align:center;font-weight:700;font-size:14px;display:flex;align-items:center;justify-content:center;min-height:42px;box-sizing:border-box;">
              {title}
            </div>
            <div style="background:#111827;color:#ffffff;padding:11px 10px;text-align:center;font-weight:700;font-size:14px;display:flex;align-items:center;justify-content:center;min-height:42px;box-sizing:border-box;">
              {x_title}
            </div>
            <div style="background:#111827;color:#ffffff;padding:11px 10px;text-align:center;font-weight:700;font-size:14px;display:flex;align-items:center;justify-content:center;min-height:42px;box-sizing:border-box;">
              {y_title}
            </div>
            <div style="background:#111827;color:#ffffff;padding:11px 10px;text-align:center;font-weight:700;font-size:14px;display:flex;align-items:center;justify-content:center;min-height:42px;box-sizing:border-box;">
              {size_col}
            </div>
          </div>

          <canvas id="forceCanvas" style="
              width:100%;height:500px;border:1px solid #ddd;border-radius:10px;
              background:transparent;touch-action:none;"></canvas>
          <div id="bubbleTip" style="
              display:none;position:absolute;pointer-events:none;
              background:rgba(30,30,30,.92);color:white;padding:8px 10px;
              border-radius:6px;font-size:12px;z-index:10;"></div>
        </div>

        <script>
        (() => {{
          const root = document.getElementById("force-bubble-root");
          const canvas = document.getElementById("forceCanvas");
          const tip = document.getElementById("bubbleTip");
          const ctx = canvas.getContext("2d");
          const data = {node_json};
          const xTitle = {x_title_json};
          const yTitle = {y_title_json};
          const sizeTitle = {size_title_json};

          let W = 900, H = 500, dpr = Math.max(1, window.devicePixelRatio || 1);
          let dragged = null;
          let pointerX = 0, pointerY = 0;

          function resize() {{
            W = Math.max(320, canvas.clientWidth);
            H = 500;
            canvas.width = W * dpr;
            canvas.height = H * dpr;
            ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
          }}
          resize();

          const cx = () => W / 2;
          const cy = () => H / 2;

          const nodes = data.map((d, i) => {{
            const angle = i * 2.399963;
            const rr = Math.min(W,H) * 0.18 * Math.sqrt((i+1)/Math.max(data.length,1));
            return {{
              ...d,
              x: cx() + Math.cos(angle)*rr,
              y: cy() + Math.sin(angle)*rr,
              vx: 0, vy: 0,
              hue: (i * 47) % 360
            }};
          }});

          function physics() {{
            // Center attraction + damping
            for (const n of nodes) {{
              if (n === dragged) continue;
              n.vx += (cx() - n.x) * 0.0009;
              n.vy += (cy() - n.y) * 0.0009;
              n.vx *= 0.94;
              n.vy *= 0.94;
            }}

            // Collision / repulsion
            for (let i=0; i<nodes.length; i++) {{
              for (let j=i+1; j<nodes.length; j++) {{
                const a=nodes[i], b=nodes[j];
                let dx=b.x-a.x, dy=b.y-a.y;
                let dist=Math.sqrt(dx*dx+dy*dy) || 0.01;
                const minDist=a.radius+b.radius+5;
                if (dist < minDist) {{
                  const overlap=(minDist-dist)/dist*0.055;
                  const fx=dx*overlap, fy=dy*overlap;
                  if (a !== dragged) {{ a.vx-=fx; a.vy-=fy; }}
                  if (b !== dragged) {{ b.vx+=fx; b.vy+=fy; }}
                }}
              }}
            }}

            for (const n of nodes) {{
              if (n === dragged) {{
                n.x = pointerX; n.y = pointerY;
                n.vx = 0; n.vy = 0;
              }} else {{
                n.x += n.vx; n.y += n.vy;
              }}
              n.x = Math.max(n.radius+4, Math.min(W-n.radius-4, n.x));
              n.y = Math.max(n.radius+4, Math.min(H-n.radius-4, n.y));
            }}
          }}

          function draw() {{
            ctx.clearRect(0,0,W,H);
            for (const n of nodes) {{
              ctx.beginPath();
              ctx.arc(n.x,n.y,n.radius,0,Math.PI*2);
              // 高對比、高彩度氣泡
              const bubbleColors = [
                ["#15803d", "#4ade80"],
                ["#b45309", "#fbbf24"],
                ["#b91c1c", "#f87171"],
                ["#1d4ed8", "#60a5fa"],
                ["#6d28d9", "#c084fc"],
                ["#0e7490", "#22d3ee"],
                ["#be185d", "#f472b6"]
              ];
              const pair = bubbleColors[nodes.indexOf(n) % bubbleColors.length];

              const grad = ctx.createRadialGradient(
                n.x - n.radius * 0.30,
                n.y - n.radius * 0.32,
                n.radius * 0.06,
                n.x,
                n.y,
                n.radius
              );
              grad.addColorStop(0, pair[1]);
              grad.addColorStop(1, pair[0]);

              ctx.fillStyle = grad;
              ctx.shadowColor = pair[1];
              ctx.shadowBlur = 14;
              ctx.fill();

              ctx.shadowBlur = 0;
              ctx.strokeStyle = pair[1];
              ctx.lineWidth = 2;
              ctx.stroke();

              // 氣泡內文字：水平置中 + 垂直置中 + 白色粗體
              ctx.textAlign = "center";
              ctx.textBaseline = "middle";
              ctx.fillStyle = "#ffffff";
              ctx.shadowColor = "rgba(0,0,0,0.85)";
              ctx.shadowBlur = 4;

              const label = n.label.length > 12 ? n.label.slice(0,11)+"…" : n.label;
              const labelSize = Math.max(12, Math.min(18, n.radius / 2.7));
              const valueSize = Math.max(12, Math.min(17, n.radius / 3));

              // 名稱與數值以氣泡中心為基準上下對稱
              ctx.font = `800 ${{labelSize}}px Arial, "Microsoft JhengHei", sans-serif`;
              ctx.fillText(label, n.x, n.y - 10);

              ctx.font = `800 ${{valueSize}}px Arial, "Microsoft JhengHei", sans-serif`;
              ctx.fillText(String(n.sizeValue), n.x, n.y + 12);

              ctx.shadowBlur = 0;
            }}
          }}

          function loop() {{
            physics();
            draw();
            requestAnimationFrame(loop);
          }}

          function point(evt) {{
            const r=canvas.getBoundingClientRect();
            return {{
              x:(evt.clientX-r.left)*(W/r.width),
              y:(evt.clientY-r.top)*(H/r.height)
            }};
          }}

          function hit(x,y) {{
            for (let i=nodes.length-1;i>=0;i--) {{
              const n=nodes[i], dx=x-n.x, dy=y-n.y;
              if (dx*dx+dy*dy <= n.radius*n.radius) return n;
            }}
            return null;
          }}

          canvas.addEventListener("pointerdown", e => {{
            const p=point(e); pointerX=p.x; pointerY=p.y;
            dragged=hit(p.x,p.y);
            if (dragged) canvas.setPointerCapture(e.pointerId);
          }});
          canvas.addEventListener("pointermove", e => {{
            const p=point(e); pointerX=p.x; pointerY=p.y;
            if (dragged) return;
            const n=hit(p.x,p.y);
            if (n) {{
              const rr=root.getBoundingClientRect();
              tip.style.display="block";
              tip.style.left=(e.clientX-rr.left+12)+"px";
              tip.style.top=(e.clientY-rr.top+12)+"px";
              tip.innerHTML =
                "<b>"+n.label+"</b><br>"+
                xTitle+"："+n.xValue+"<br>"+
                yTitle+"："+n.yValue+"<br>"+
                sizeTitle+"："+n.sizeValue;
              canvas.style.cursor="grab";
            }} else {{
              tip.style.display="none";
              canvas.style.cursor="default";
            }}
          }});
          canvas.addEventListener("pointerup", e => {{
            dragged=null;
            try {{ canvas.releasePointerCapture(e.pointerId); }} catch(err) {{}}
          }});
          canvas.addEventListener("pointerleave", () => {{
            if (!dragged) tip.style.display="none";
          }});

          if (window.ResizeObserver) {{
            new ResizeObserver(() => resize()).observe(canvas);
          }}
          loop();
        }})();
        </script>
        """
        components.html(html, height=650, scrolling=False)

    st.subheader("📈 RFQ / 報價 KPI 儀表板")
    st.caption("依 Tasks、Quotes、StatusLog 即時計算；不另外輸入 KPI 數字。")

    today = pd.Timestamp(datetime.now().date())
    month_start = today.replace(day=1)

    kpi_df = df.copy()

    # ---------------------------------------------------------
    # 日期正規化
    # ---------------------------------------------------------
    if "created_time" not in kpi_df.columns:
        kpi_df["created_time"] = ""

    if "due_time" not in kpi_df.columns:
        kpi_df["due_time"] = ""

    if "status" not in kpi_df.columns:
        kpi_df["status"] = ""

    if "owner" not in kpi_df.columns:
        kpi_df["owner"] = "待確認"

    kpi_df["_created_dt"] = pd.to_datetime(
        kpi_df["created_time"],
        errors="coerce"
    )

    kpi_df["_due_dt"] = pd.to_datetime(
        kpi_df["due_time"],
        errors="coerce"
    )

    # ---------------------------------------------------------
    # Quotes
    # ---------------------------------------------------------
    try:
        kpi_quotes = conn.read(
            worksheet="Quotes",
            ttl=GSHEETS_READ_TTL
        )
    except Exception:
        kpi_quotes = pd.DataFrame()

    if kpi_quotes is None:
        kpi_quotes = pd.DataFrame()

    if not kpi_quotes.empty:
        if "quote_date" not in kpi_quotes.columns:
            kpi_quotes["quote_date"] = ""

        if "quote_total" not in kpi_quotes.columns:
            kpi_quotes["quote_total"] = 0.0

        if "RFQ_ID" not in kpi_quotes.columns:
            kpi_quotes["RFQ_ID"] = ""

        kpi_quotes["_quote_dt"] = pd.to_datetime(
            kpi_quotes["quote_date"],
            errors="coerce"
        )

        kpi_quotes["_quote_total"] = pd.to_numeric(
            kpi_quotes["quote_total"],
            errors="coerce"
        ).fillna(0.0)

    # ---------------------------------------------------------
    # StatusLog
    # ---------------------------------------------------------
    try:
        kpi_log = conn.read(
            worksheet="StatusLog",
            ttl=GSHEETS_READ_TTL
        )
    except Exception:
        kpi_log = pd.DataFrame()

    if kpi_log is None:
        kpi_log = pd.DataFrame()

    if not kpi_log.empty:
        if "change_time" not in kpi_log.columns:
            kpi_log["change_time"] = ""

        if "new_status" not in kpi_log.columns:
            kpi_log["new_status"] = ""

        if "RFQ_ID" not in kpi_log.columns:
            kpi_log["RFQ_ID"] = ""

        kpi_log["_change_dt"] = pd.to_datetime(
            kpi_log["change_time"],
            errors="coerce"
        )

    # ---------------------------------------------------------
    # KPI 計算
    # ---------------------------------------------------------
    this_month_rfq = kpi_df[
        kpi_df["_created_dt"].notna()
        & (kpi_df["_created_dt"] >= month_start)
        & (kpi_df["_created_dt"] < month_start + pd.offsets.MonthBegin(1))
    ]

    month_rfq_count = len(this_month_rfq)

    # 本月已報價：優先使用 StatusLog 的實際狀態變更時間
    if not kpi_log.empty:
        month_quoted_log = kpi_log[
            (kpi_log["new_status"].astype(str) == "已報價")
            & kpi_log["_change_dt"].notna()
            & (kpi_log["_change_dt"] >= month_start)
            & (kpi_log["_change_dt"] < month_start + pd.offsets.MonthBegin(1))
        ]

        month_quoted_count = (
            month_quoted_log["RFQ_ID"]
            .astype(str)
            .replace("", pd.NA)
            .dropna()
            .nunique()
        )
    else:
        month_quoted_count = 0

    # 如果舊資料沒有 StatusLog，使用 Quotes 當備援
    if month_quoted_count == 0 and not kpi_quotes.empty:
        month_quote_rows = kpi_quotes[
            kpi_quotes["_quote_dt"].notna()
            & (kpi_quotes["_quote_dt"] >= month_start)
            & (kpi_quotes["_quote_dt"] < month_start + pd.offsets.MonthBegin(1))
        ]
        month_quoted_count = (
            month_quote_rows["RFQ_ID"]
            .astype(str)
            .replace("", pd.NA)
            .dropna()
            .nunique()
        )

    quote_rate = (
        month_quoted_count / month_rfq_count * 100
        if month_rfq_count > 0
        else 0.0
    )

    closed_count = int(
        (kpi_df["status"].astype(str) == "結案").sum()
    )

    active_count = int(
        (kpi_df["status"].astype(str) != "結案").sum()
    )

    overdue_mask = (
        kpi_df["_due_dt"].notna()
        & (kpi_df["_due_dt"].dt.normalize() < today)
        & (kpi_df["status"].astype(str) != "結案")
    )
    overdue_count = int(overdue_mask.sum())

    pending_info_count = int(
        (kpi_df["status"].astype(str) == "待補件").sum()
    )

    pending_quote_count = int(
        (kpi_df["status"].astype(str) == "待核價").sum()
    )

    if not kpi_quotes.empty:
        quote_total_sum = float(
            kpi_quotes["_quote_total"].sum()
        )

        valid_quote_amounts = kpi_quotes[
            kpi_quotes["_quote_total"] > 0
        ]

        avg_quote_amount = (
            float(valid_quote_amounts["_quote_total"].mean())
            if not valid_quote_amounts.empty
            else 0.0
        )
    else:
        quote_total_sum = 0.0
        avg_quote_amount = 0.0

    # ---------------------------------------------------------
    # 平均報價時間：RFQ created_time → 首次進入已報價
    # ---------------------------------------------------------
    avg_quote_days = None

    if not kpi_log.empty and not kpi_df.empty:
        quoted_logs = kpi_log[
            (kpi_log["new_status"].astype(str) == "已報價")
            & kpi_log["_change_dt"].notna()
        ].copy()

        if not quoted_logs.empty:
            first_quote = (
                quoted_logs
                .sort_values("_change_dt")
                .groupby("RFQ_ID", as_index=False)
                .first()[["RFQ_ID", "_change_dt"]]
            )

            created_map = kpi_df[
                ["RFQ_ID", "_created_dt"]
            ].copy()

            quote_cycle = first_quote.merge(
                created_map,
                on="RFQ_ID",
                how="inner"
            )

            quote_cycle["days"] = (
                quote_cycle["_change_dt"]
                - quote_cycle["_created_dt"]
            ).dt.total_seconds() / 86400

            quote_cycle = quote_cycle[
                quote_cycle["days"].notna()
                & (quote_cycle["days"] >= 0)
            ]

            if not quote_cycle.empty:
                avg_quote_days = float(
                    quote_cycle["days"].mean()
                )

    # ---------------------------------------------------------
    # KPI Cards
    # ---------------------------------------------------------
    st.markdown("### 本月 KPI")

    k1, k2, k3, k4, k5 = st.columns(5)

    k1.metric(
        "本月新增 RFQ",
        f"{month_rfq_count} 件"
    )

    k2.metric(
        "本月已報價",
        f"{month_quoted_count} 件"
    )

    k3.metric(
        "本月報價率",
        f"{quote_rate:.1f}%"
    )

    k4.metric(
        "進行中案件",
        f"{active_count} 件"
    )

    k5.metric(
        "逾期案件",
        f"{overdue_count} 件"
    )

    st.markdown("### 管理 KPI")

    k6, k7, k8, k9, k10 = st.columns(5)

    k6.metric(
        "待補件",
        f"{pending_info_count} 件"
    )

    k7.metric(
        "待核價",
        f"{pending_quote_count} 件"
    )

    k8.metric(
        "結案",
        f"{closed_count} 件"
    )

    k9.metric(
        "報價總額",
        f"{money(quote_total_sum)}"
    )

    k10.metric(
        "平均報價額",
        f"{money(avg_quote_amount)}"
    )

    if avg_quote_days is None:
        st.info(
            "平均報價時間：目前沒有足夠的「RFQ 建立 → 已報價」StatusLog 資料。"
        )
    else:
        st.metric(
            "⏱️ 平均報價時間",
            f"{avg_quote_days:.1f} 天"
        )

    st.divider()

    # ---------------------------------------------------------
    # 狀態分布
    # ---------------------------------------------------------
    st.markdown("### 📊 RFQ 狀態分布")

    status_summary = (
        kpi_df["status"]
        .fillna("待確認")
        .astype(str)
        .value_counts()
        .reindex(RFQ_STATUS, fill_value=0)
        .rename_axis("狀態")
        .reset_index(name="案件數")
    )

    st.empty().bar_chart(
        status_summary.set_index("狀態")["案件數"],
        use_container_width=True
    )

    # ---------------------------------------------------------
    # 每月 RFQ / 報價趨勢
    # ---------------------------------------------------------
    st.markdown("### 📅 每月詢價 / 報價趨勢")

    rfq_monthly = (
        kpi_df[
            kpi_df["_created_dt"].notna()
        ]
        .assign(
            月份=lambda x:
                x["_created_dt"].dt.to_period("M").astype(str)
        )
        .groupby("月份")
        .size()
        .rename("RFQ")
    )

    if not kpi_quotes.empty:
        quote_monthly = (
            kpi_quotes[
                kpi_quotes["_quote_dt"].notna()
            ]
            .assign(
                月份=lambda x:
                    x["_quote_dt"].dt.to_period("M").astype(str)
            )
            .groupby("月份")
            .size()
            .rename("報價")
        )
    else:
        quote_monthly = pd.Series(
            dtype="int64",
            name="報價"
        )

    trend_df = pd.concat(
        [rfq_monthly, quote_monthly],
        axis=1
    ).fillna(0)

    if trend_df.empty:
        st.info("目前沒有足夠日期資料可顯示趨勢。")
    else:
        trend_df = trend_df.astype(int).sort_index()
        # 發表版：每月詢價與報價採並列長條圖呈現
        st.empty().bar_chart(
            trend_df,
            use_container_width=True
        )
        st.dataframe(
            trend_df.reset_index(),
            use_container_width=True,
            hide_index=True
        )

    # ---------------------------------------------------------
    # 負責人 KPI
    # ---------------------------------------------------------
    st.markdown("### 👤 負責人 KPI")

    owner_base = kpi_df.copy()

    owner_base["owner"] = (
        owner_base["owner"]
        .fillna("待確認")
        .astype(str)
        .replace("", "待確認")
    )

    owner_total = (
        owner_base
        .groupby("owner")
        .size()
        .rename("總案件")
    )

    owner_active = (
        owner_base[
            owner_base["status"].astype(str) != "結案"
        ]
        .groupby("owner")
        .size()
        .rename("進行中")
    )

    owner_closed = (
        owner_base[
            owner_base["status"].astype(str) == "結案"
        ]
        .groupby("owner")
        .size()
        .rename("結案")
    )

    owner_overdue = (
        owner_base[
            overdue_mask
        ]
        .groupby("owner")
        .size()
        .rename("逾期")
    )

    owner_kpi = pd.concat(
        [
            owner_total,
            owner_active,
            owner_closed,
            owner_overdue
        ],
        axis=1
    ).fillna(0).astype(int)

    owner_kpi["結案率"] = (
        owner_kpi["結案"]
        / owner_kpi["總案件"]
        * 100
    ).round(1)

    owner_kpi = (
        owner_kpi
        .reset_index()
        .rename(columns={"owner": "負責人"})
        .sort_values(
            ["總案件", "結案"],
            ascending=[False, False]
        )
    )

    st.dataframe(
        owner_kpi,
        use_container_width=True,
        hide_index=True,
        column_config={
            "結案率": st.column_config.NumberColumn(
                "結案率 (%)",
                format="%.1f%%"
            )
        }
    )

    # ---------------------------------------------------------
    # 逾期案件明細
    # ---------------------------------------------------------
    st.markdown("### ⚠️ 逾期案件")

    overdue_df = kpi_df[
        overdue_mask
    ].copy()

    if overdue_df.empty:
        st.success("目前沒有逾期案件。")
    else:
        overdue_columns = [
            "RFQ_ID",
            "customer",
            "title",
            "status",
            "owner",
            "due_time",
            "next_step"
        ]

        overdue_columns = [
            col
            for col in overdue_columns
            if col in overdue_df.columns
        ]

        st.dataframe(
            overdue_df[overdue_columns],
            use_container_width=True,
            hide_index=True
        )



# ################################################################
# TAB 5：RFQ 甘特圖
# ################################################################


    st.divider()
    st.markdown("### 🫧 KPI 氣泡分析")

    # RFQ status bubble: X = workflow stage, Y = share, size = case count
    if not df.empty and "status" in df.columns:
        status_bubble = (
            df["status"].fillna("未分類").astype(str)
            .value_counts().rename_axis("狀態").reset_index(name="案件數")
        )
        status_order = {s: i + 1 for i, s in enumerate(RFQ_STATUS)}
        status_bubble["流程階段"] = status_bubble["狀態"].map(status_order).fillna(0)
        status_bubble["案件占比 (%)"] = (
            status_bubble["案件數"] / status_bubble["案件數"].sum() * 100
        ).round(1)
        render_kpi_bubble(
            status_bubble, "流程階段", "案件占比 (%)", "案件數", "狀態",
            "RFQ 狀態分布氣泡圖", "流程階段", "案件占比 (%)"
        )

    # Monthly RFQ bubble: X = month order, Y = RFQ count, size = RFQ count
    if not df.empty and "created_time" in df.columns:
        mb = df.copy()
        mb["_dt"] = pd.to_datetime(mb["created_time"], errors="coerce")
        mb = mb[mb["_dt"].notna()]
        if not mb.empty:
            mb["月份"] = mb["_dt"].dt.to_period("M").astype(str)
            mb = mb.groupby("月份").size().reset_index(name="RFQ件數").sort_values("月份")
            mb["月份序號"] = range(1, len(mb) + 1)
            render_kpi_bubble(
                mb, "月份序號", "RFQ件數", "RFQ件數", "月份",
                "每月 RFQ 趨勢氣泡圖", "月份", "RFQ 件數"
            )

    # Owner bubble built directly from Tasks, independent of prior KPI variable names.
    if not df.empty and "owner" in df.columns:
        ob = df.copy()
        ob["owner"] = ob["owner"].fillna("待確認").astype(str).replace("", "待確認")
        ob["_closed"] = ob["status"].astype(str).eq("結案")
        owner_bubble = (
            ob.groupby("owner")
            .agg(總案件=("RFQ_ID", "size"), 結案件數=("_closed", "sum"))
            .reset_index()
        )
        owner_bubble["結案率 (%)"] = (
            owner_bubble["結案件數"] / owner_bubble["總案件"].replace(0, pd.NA) * 100
        ).fillna(0).round(1)
        render_kpi_bubble(
            owner_bubble, "總案件", "結案率 (%)", "結案件數", "owner",
            "業務案件量 × 結案率氣泡圖", "總案件數", "結案率 (%)"
        )

with tab_gantt:

    st.subheader("📅 RFQ 甘特圖")
    st.caption("以 RFQ 建立日期為開始日、到期日為結束日，快速查看案件時程、負責人與逾期狀況。")

    gantt_df = df.copy()

    if gantt_df.empty:
        st.info("目前沒有 RFQ 資料。")

    else:
        # ---------------------------------------------------------
        # 欄位與日期正規化
        # ---------------------------------------------------------
        for col in [
            "RFQ_ID", "customer", "title", "owner",
            "status", "created_time", "due_time", "next_step"
        ]:
            if col not in gantt_df.columns:
                gantt_df[col] = ""

        gantt_df["_start"] = pd.to_datetime(
            gantt_df["created_time"],
            errors="coerce"
        )

        gantt_df["_end"] = pd.to_datetime(
            gantt_df["due_time"],
            errors="coerce"
        )

        gantt_df["owner"] = (
            gantt_df["owner"]
            .fillna("待確認")
            .astype(str)
            .replace("", "待確認")
        )

        gantt_df["status"] = (
            gantt_df["status"]
            .fillna("待確認")
            .astype(str)
            .replace("", "待確認")
        )

        gantt_df["customer"] = gantt_df["customer"].fillna("").astype(str)
        gantt_df["title"] = gantt_df["title"].fillna("").astype(str)
        gantt_df["RFQ_ID"] = gantt_df["RFQ_ID"].fillna("").astype(str)

        # ---------------------------------------------------------
        # 篩選器
        # ---------------------------------------------------------
        f1, f2, f3 = st.columns([1, 1, 1.4])

        owner_options = sorted(
            [
                x for x in gantt_df["owner"].dropna().unique().tolist()
                if str(x).strip()
            ]
        )

        status_options = [
            s for s in RFQ_STATUS
            if s in gantt_df["status"].unique().tolist()
        ]

        with f1:
            selected_owners = st.multiselect(
                "👤 負責人",
                options=owner_options,
                default=[],
                key="gantt_owner_filter"
            )

        with f2:
            selected_statuses = st.multiselect(
                "📌 狀態",
                options=status_options,
                default=[],
                key="gantt_status_filter"
            )

        with f3:
            gantt_search = st.text_input(
                "🔎 搜尋",
                placeholder="RFQ ID / 客戶 / 品名",
                key="gantt_search"
            ).strip()

        show_closed = st.checkbox(
            "顯示已結案案件",
            value=True,
            key="gantt_show_closed"
        )

        # ---------------------------------------------------------
        # 甘特圖顯示期間
        # ---------------------------------------------------------
        valid_start_dates = gantt_df["_start"].dropna()
        valid_end_dates = gantt_df["_end"].dropna()

        default_period_start = (
            valid_start_dates.min().date()
            if not valid_start_dates.empty
            else datetime.now().date()
        )

        default_period_end = (
            valid_end_dates.max().date()
            if not valid_end_dates.empty
            else datetime.now().date() + timedelta(days=30)
        )

        if default_period_end < default_period_start:
            default_period_end = default_period_start + timedelta(days=30)

        st.markdown("#### 🗓️ 甘特圖顯示期間")

        period_col1, period_col2 = st.columns(2)

        with period_col1:
            gantt_period_start = st.date_input(
                "起始日期",
                value=default_period_start,
                key="gantt_period_start"
            )

        with period_col2:
            gantt_period_end = st.date_input(
                "結束日期",
                value=default_period_end,
                key="gantt_period_end"
            )

        if gantt_period_end < gantt_period_start:
            st.error("❌ 甘特圖結束日期不可早於起始日期。")
            st.stop()

        period_start_ts = pd.Timestamp(gantt_period_start)
        # 結束日包含整天
        period_end_ts = pd.Timestamp(gantt_period_end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

        filtered_gantt = gantt_df.copy()

        if selected_owners:
            filtered_gantt = filtered_gantt[
                filtered_gantt["owner"].isin(selected_owners)
            ]

        if selected_statuses:
            filtered_gantt = filtered_gantt[
                filtered_gantt["status"].isin(selected_statuses)
            ]

        if not show_closed:
            filtered_gantt = filtered_gantt[
                filtered_gantt["status"] != "結案"
            ]

        if gantt_search:
            search_mask = (
                filtered_gantt["RFQ_ID"].str.contains(
                    gantt_search, case=False, na=False
                )
                | filtered_gantt["customer"].str.contains(
                    gantt_search, case=False, na=False
                )
                | filtered_gantt["title"].str.contains(
                    gantt_search, case=False, na=False
                )
            )
            filtered_gantt = filtered_gantt[search_mask]

        # 顯示與指定期間有交集的案件：
        # RFQ開始日 <= 顯示期間結束日 且 RFQ結束日 >= 顯示期間起始日
        period_overlap_mask = (
            filtered_gantt["_start"].notna()
            & filtered_gantt["_end"].notna()
            & (filtered_gantt["_start"] <= period_end_ts)
            & (filtered_gantt["_end"] >= period_start_ts)
        )

        missing_period_date_mask = (
            filtered_gantt["_start"].isna()
            | filtered_gantt["_end"].isna()
        )

        filtered_gantt = filtered_gantt[
            period_overlap_mask | missing_period_date_mask
        ]

        # 甘特圖需要開始日；沒有到期日則無法畫完整時程
        valid_gantt = filtered_gantt[
            filtered_gantt["_start"].notna()
            & filtered_gantt["_end"].notna()
        ].copy()

        # 若 due_time 比 created_time 早，先排除並在下方提醒
        invalid_date_mask = (
            valid_gantt["_end"] < valid_gantt["_start"]
        )
        invalid_date_count = int(invalid_date_mask.sum())

        valid_gantt = valid_gantt[
            ~invalid_date_mask
        ].copy()

        today_gantt = pd.Timestamp(datetime.now().date())

        valid_gantt["逾期"] = (
            (valid_gantt["_end"].dt.normalize() < today_gantt)
            & (valid_gantt["status"] != "結案")
        )

        valid_gantt["工期天數"] = (
            valid_gantt["_end"].dt.normalize()
            - valid_gantt["_start"].dt.normalize()
        ).dt.days + 1

        # ---------------------------------------------------------
        # 摘要
        # ---------------------------------------------------------
        g1, g2, g3, g4 = st.columns(4)

        g1.metric(
            "顯示案件",
            f"{len(valid_gantt)} 件"
        )

        g2.metric(
            "逾期",
            f"{int(valid_gantt['逾期'].sum()) if not valid_gantt.empty else 0} 件"
        )

        g3.metric(
            "無完整日期",
            f"{len(filtered_gantt) - len(valid_gantt) - invalid_date_count} 件"
        )

        g4.metric(
            "日期異常",
            f"{invalid_date_count} 件"
        )

        st.divider()

        if valid_gantt.empty:
            st.info("目前篩選條件下沒有可繪製甘特圖的 RFQ。請確認 created_time 與 due_time 都有日期。")

        else:
            # -----------------------------------------------------
            # 使用 Plotly 畫真正的甘特圖
            # -----------------------------------------------------
            st.caption(
                f"顯示期間：{gantt_period_start.strftime('%Y-%m-%d')} ～ "
                f"{gantt_period_end.strftime('%Y-%m-%d')}"
            )

            try:
                import plotly.express as px

                plot_df = valid_gantt.copy()

                plot_df["案件"] = (
                    plot_df["RFQ_ID"]
                    + "｜"
                    + plot_df["customer"]
                    + "｜"
                    + plot_df["title"]
                )

                # 同一狀態仍用 Plotly 自動配色；逾期另加文字提示
                plot_df["顯示狀態"] = plot_df["status"]
                plot_df.loc[
                    plot_df["逾期"],
                    "顯示狀態"
                ] = plot_df.loc[
                    plot_df["逾期"],
                    "status"
                ] + "｜逾期"

                plot_df = plot_df.sort_values(
                    ["_start", "_end"],
                    ascending=[True, True]
                )

                fig = px.timeline(
                    plot_df,
                    x_start="_start",
                    x_end="_end",
                    y="案件",
                    color="顯示狀態",
                    hover_name="RFQ_ID",
                    hover_data={
                        "customer": True,
                        "title": True,
                        "owner": True,
                        "status": True,
                        "_start": "|%Y-%m-%d",
                        "_end": "|%Y-%m-%d",
                        "工期天數": True,
                        "逾期": True,
                        "案件": False,
                        "顯示狀態": False,
                    },
                    labels={
                        "_start": "開始日",
                        "_end": "到期日",
                        "customer": "客戶",
                        "title": "品名",
                        "owner": "負責人",
                        "status": "狀態",
                        "工期天數": "工期",
                        "逾期": "是否逾期",
                        "顯示狀態": "狀態",
                    },
                )

                fig.update_yaxes(
                    autorange="reversed",
                    title=""
                )

                fig.update_xaxes(
                    title="日期",
                    showgrid=True,
                    range=[
                        pd.Timestamp(gantt_period_start),
                        pd.Timestamp(gantt_period_end) + pd.Timedelta(days=1)
                    ]
                )

                fig.update_layout(
                    height=max(
                        430,
                        min(1100, 170 + len(plot_df) * 38)
                    ),
                    margin=dict(
                        l=10,
                        r=10,
                        t=35,
                        b=20
                    ),
                    legend_title_text="狀態",
                    hoverlabel=dict(
                        namelength=-1
                    ),
                )

                # 今天參考線
                fig.add_vline(
                    x=today_gantt.timestamp() * 1000,
                    line_dash="dash",
                    annotation_text="今天",
                    annotation_position="top"
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True,
                    key="rfq_gantt_chart"
                )

            except ImportError:
                st.warning(
                    "甘特圖需要 plotly。請在 requirements.txt 加入 plotly 後重新部署。"
                )

            except Exception as e:
                st.warning(f"甘特圖產生失敗：{e}")

            # -----------------------------------------------------
            # 甘特明細
            # -----------------------------------------------------
            st.markdown("### 📋 甘特圖明細")

            gantt_detail = valid_gantt.copy()

            gantt_detail["開始日"] = (
                gantt_detail["_start"]
                .dt.strftime("%Y-%m-%d")
            )

            gantt_detail["到期日"] = (
                gantt_detail["_end"]
                .dt.strftime("%Y-%m-%d")
            )

            gantt_detail["逾期狀態"] = gantt_detail["逾期"].map(
                {
                    True: "⚠️ 逾期",
                    False: ""
                }
            )

            gantt_detail = gantt_detail[
                [
                    "RFQ_ID",
                    "customer",
                    "title",
                    "owner",
                    "status",
                    "開始日",
                    "到期日",
                    "工期天數",
                    "逾期狀態",
                    "next_step",
                ]
            ].rename(
                columns={
                    "customer": "客戶",
                    "title": "品名",
                    "owner": "負責人",
                    "status": "狀態",
                    "next_step": "下一步",
                }
            )

            st.dataframe(
                gantt_detail,
                use_container_width=True,
                hide_index=True
            )

        # ---------------------------------------------------------
        # 日期不完整案件
        # ---------------------------------------------------------
        missing_date_df = filtered_gantt[
            filtered_gantt["_start"].isna()
            | filtered_gantt["_end"].isna()
        ].copy()

        if not missing_date_df.empty:
            with st.expander(
                f"⚠️ 尚未顯示於甘特圖：缺少日期 {len(missing_date_df)} 件"
            ):
                missing_cols = [
                    "RFQ_ID",
                    "customer",
                    "title",
                    "owner",
                    "status",
                    "created_time",
                    "due_time",
                ]

                st.dataframe(
                    missing_date_df[missing_cols].rename(
                        columns={
                            "customer": "客戶",
                            "title": "品名",
                            "owner": "負責人",
                            "status": "狀態",
                            "created_time": "建立日期",
                            "due_time": "到期日",
                        }
                    ),
                    use_container_width=True,
                    hide_index=True
                )


# ################################################################
# TAB 6：產品販售分析 / 業務銷售分析
# ################################################################

with tab_sales:
    st.subheader("📊 產品販售分析 / 業務銷售分析")
    st.caption("分析來源：Tasks + Quotes。成交分析僅使用明確標記為「已成交」的案件，不把「結案」自動視為成交。")

    st.markdown("### ✍️ 成交資料登錄")
    if df.empty:
        st.info("目前沒有 RFQ 可登錄成交資料。")
    else:
        sales_edit_options = (
            df["RFQ_ID"].fillna("").astype(str)
            + "｜"
            + df["customer"].fillna("").astype(str)
            + "｜"
            + df["title"].fillna("").astype(str)
        ).tolist()

        selected_sales_case = st.selectbox(
            "選擇 RFQ",
            options=sales_edit_options,
            key="sales_edit_rfq"
        )
        selected_sales_idx = sales_edit_options.index(selected_sales_case)
        selected_sales_row = df.iloc[selected_sales_idx]

        ec1, ec2, ec3 = st.columns(3)
        with ec1:
            edit_result = st.selectbox(
                "成交狀態",
                ["進行中", "已成交", "未成交"],
                index=(
                    ["進行中", "已成交", "未成交"].index(str(selected_sales_row.get("sales_result", "進行中")))
                    if str(selected_sales_row.get("sales_result", "進行中")) in ["進行中", "已成交", "未成交"]
                    else 0
                ),
                key="sales_edit_result"
            )
        with ec2:
            existing_deal_date = pd.to_datetime(selected_sales_row.get("deal_date", ""), errors="coerce")
            edit_deal_date = st.date_input(
                "成交日期",
                value=(existing_deal_date.date() if pd.notna(existing_deal_date) else datetime.now().date()),
                key="sales_edit_date"
            )
        with ec3:
            edit_deal_amount = st.number_input(
                "成交金額",
                min_value=0.0,
                value=float(pd.to_numeric(selected_sales_row.get("deal_amount", 0), errors="coerce") or 0),
                step=1000.0,
                key="sales_edit_amount"
            )

        lost_options = ["", "價格", "交期", "規格", "客戶取消", "競爭對手", "其他"]
        current_lost = str(selected_sales_row.get("lost_reason", "") or "")
        edit_lost_reason = st.selectbox(
            "未成交原因",
            lost_options,
            index=(lost_options.index(current_lost) if current_lost in lost_options else 0),
            disabled=(edit_result != "未成交"),
            key="sales_edit_lost_reason"
        )

        if st.button("💾 儲存成交資料", type="primary", key="save_sales_result"):
            real_idx = df.index[selected_sales_idx]
            df.loc[real_idx, "sales_result"] = edit_result
            df.loc[real_idx, "deal_date"] = (
                edit_deal_date.strftime("%Y-%m-%d")
                if edit_result == "已成交"
                else ""
            )
            df.loc[real_idx, "deal_amount"] = (
                float(edit_deal_amount)
                if edit_result == "已成交"
                else 0.0
            )
            df.loc[real_idx, "lost_reason"] = (
                edit_lost_reason
                if edit_result == "未成交"
                else ""
            )
            conn.update(worksheet="Tasks", data=df)
            clear_gsheets_cache()
            st.success("成交資料已儲存。")
            st.rerun()

    st.divider()

    sales_df = df.copy()

    # 相容舊資料：缺少的新欄位自動補空值
    sales_defaults = {
        "product_category": "",
        "product_model": "",
        "sales_result": "進行中",
        "deal_date": "",
        "deal_amount": 0.0,
        "lost_reason": "",
        "owner": "待確認",
        "customer": "",
        "created_time": "",
        "status": "",
        "RFQ_ID": "",
        "title": "",
        "continent": "",
        "country": "",
    }
    for col, default in sales_defaults.items():
        if col not in sales_df.columns:
            sales_df[col] = default

    sales_df["product_category"] = sales_df["product_category"].fillna("").astype(str)
    sales_df["product_model"] = sales_df["product_model"].fillna("").astype(str)
    sales_df["owner"] = sales_df["owner"].fillna("待確認").astype(str).replace("", "待確認")
    sales_df["customer"] = sales_df["customer"].fillna("").astype(str)
    sales_df["sales_result"] = sales_df["sales_result"].fillna("進行中").astype(str).replace("", "進行中")
    sales_df["continent"] = sales_df["continent"].fillna("").astype(str).replace("", "未分類")
    sales_df["country"] = sales_df["country"].fillna("").astype(str).replace("", "未分類")
    sales_df["_created_dt"] = pd.to_datetime(sales_df["created_time"], errors="coerce")
    sales_df["_deal_dt"] = pd.to_datetime(sales_df["deal_date"], errors="coerce")
    sales_df["_deal_amount"] = pd.to_numeric(sales_df["deal_amount"], errors="coerce").fillna(0.0)

    # 舊案件若只有 title，可暫時以 title 當型號顯示，避免分析完全空白。
    sales_df["_analysis_model"] = sales_df["product_model"].where(
        sales_df["product_model"].str.strip() != "",
        sales_df["title"].fillna("").astype(str)
    )
    sales_df["_analysis_category"] = sales_df["product_category"].where(
        sales_df["product_category"].str.strip() != "",
        "未分類"
    )

    try:
        sales_quotes = conn.read(worksheet="Quotes", ttl=GSHEETS_READ_TTL)
    except Exception:
        sales_quotes = pd.DataFrame()

    if sales_quotes is None:
        sales_quotes = pd.DataFrame()

    if not sales_quotes.empty:
        if "RFQ_ID" not in sales_quotes.columns:
            sales_quotes["RFQ_ID"] = ""
        if "quote_total" not in sales_quotes.columns:
            sales_quotes["quote_total"] = 0.0
        if "quote_date" not in sales_quotes.columns:
            sales_quotes["quote_date"] = ""

        sales_quotes["_quote_total"] = pd.to_numeric(
            sales_quotes["quote_total"], errors="coerce"
        ).fillna(0.0)
        sales_quotes["_quote_dt"] = pd.to_datetime(
            sales_quotes["quote_date"], errors="coerce"
        )

        quote_by_rfq = (
            sales_quotes.groupby("RFQ_ID", as_index=False)
            .agg(
                報價次數=("RFQ_ID", "size"),
                報價金額=("_quote_total", "sum"),
                首次報價日=("_quote_dt", "min"),
            )
        )
        sales_df = sales_df.merge(quote_by_rfq, on="RFQ_ID", how="left")
    else:
        sales_df["報價次數"] = 0
        sales_df["報價金額"] = 0.0
        sales_df["首次報價日"] = pd.NaT

    sales_df["報價次數"] = pd.to_numeric(sales_df["報價次數"], errors="coerce").fillna(0).astype(int)
    sales_df["報價金額"] = pd.to_numeric(sales_df["報價金額"], errors="coerce").fillna(0.0)
    sales_df["_quoted"] = sales_df["報價次數"] > 0
    sales_df["_won"] = sales_df["sales_result"] == "已成交"

    # 分析期間
    valid_dates = pd.concat([
        sales_df["_created_dt"].dropna(),
        sales_df["_deal_dt"].dropna()
    ])
    default_start = valid_dates.min().date() if not valid_dates.empty else datetime.now().date().replace(day=1)
    default_end = valid_dates.max().date() if not valid_dates.empty else datetime.now().date()

    a1, a2, a3 = st.columns([1, 1, 1.2])
    with a1:
        analysis_start = st.date_input("分析起始日", value=default_start, key="sales_analysis_start")
    with a2:
        analysis_end = st.date_input("分析結束日", value=default_end, key="sales_analysis_end")
    with a3:
        analysis_basis = st.selectbox(
            "期間依據",
            ["詢價建立日", "成交日"],
            key="sales_analysis_basis"
        )

    if analysis_end < analysis_start:
        st.error("分析結束日不可早於起始日。")
    else:
        s = pd.Timestamp(analysis_start)
        e = pd.Timestamp(analysis_end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        basis_col = "_created_dt" if analysis_basis == "詢價建立日" else "_deal_dt"
        analysis_df = sales_df[
            sales_df[basis_col].notna()
            & (sales_df[basis_col] >= s)
            & (sales_df[basis_col] <= e)
        ].copy()

        total_rfq = len(analysis_df)
        quoted_count = int(analysis_df["_quoted"].sum())
        won_count = int(analysis_df["_won"].sum())
        quote_amount = float(analysis_df["報價金額"].sum())
        deal_amount = float(analysis_df.loc[analysis_df["_won"], "_deal_amount"].sum())
        win_rate = won_count / quoted_count * 100 if quoted_count else 0.0

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("詢價件數", f"{total_rfq} 件")
        m2.metric("已報價", f"{quoted_count} 件")
        m3.metric("已成交", f"{won_count} 件")
        m4.metric("成交率", f"{win_rate:.1f}%")
        m5.metric("成交金額", money(deal_amount))

        st.caption(f"期間報價總額：{money(quote_amount)}")

        product_tab, owner_tab, continent_tab, country_tab, country_purchase_tab = st.tabs(
            [
                "📦 產品販售分析",
                "👤 業務銷售分析",
                "🌍 洲別分析",
                "🌐 國家分析",
                "🛒 國家採購分析",
            ]
        )

        with product_tab:
            category_summary = (
                analysis_df.groupby("_analysis_category", dropna=False)
                .agg(
                    詢價件數=("RFQ_ID", "size"),
                    報價件數=("_quoted", "sum"),
                    報價金額=("報價金額", "sum"),
                    成交件數=("_won", "sum"),
                    成交金額=("_deal_amount", lambda x: x[analysis_df.loc[x.index, "_won"]].sum()),
                )
                .reset_index()
                .rename(columns={"_analysis_category": "產品分類"})
            )
            if not category_summary.empty:
                category_summary["成交率 (%)"] = (
                    category_summary["成交件數"]
                    / category_summary["報價件數"].replace(0, pd.NA)
                    * 100
                ).fillna(0).round(1)
                category_summary = category_summary.sort_values("成交金額", ascending=False)
                st.markdown("#### 產品分類績效")
                st.dataframe(category_summary, use_container_width=True, hide_index=True)

            model_summary = (
                analysis_df.groupby("_analysis_model", dropna=False)
                .agg(
                    詢價件數=("RFQ_ID", "size"),
                    報價件數=("_quoted", "sum"),
                    報價金額=("報價金額", "sum"),
                    成交件數=("_won", "sum"),
                    成交金額=("_deal_amount", lambda x: x[analysis_df.loc[x.index, "_won"]].sum()),
                )
                .reset_index()
                .rename(columns={"_analysis_model": "產品型號"})
            )
            model_summary = model_summary[model_summary["產品型號"].astype(str).str.strip() != ""]
            if model_summary.empty:
                st.info("目前期間內沒有產品型號資料。")
            else:
                model_summary["成交率 (%)"] = (
                    model_summary["成交件數"]
                    / model_summary["報價件數"].replace(0, pd.NA)
                    * 100
                ).fillna(0).round(1)
                model_summary = model_summary.sort_values(
                    ["成交金額", "成交件數", "詢價件數"], ascending=False
                )
                st.markdown("#### 產品型號績效")
                st.dataframe(model_summary, use_container_width=True, hide_index=True)

                top_models = model_summary[model_summary["成交金額"] > 0].head(10).copy()
                st.markdown("#### Top 10 型號－成交金額")

                if top_models.empty:
                    st.info("目前尚無已成交金額資料；有成交紀錄後將自動顯示 Top 10 型號排行。")
                else:
                    # 發表版：橫式型號文字、金額標籤、由高至低
                    top_models = top_models.sort_values("成交金額", ascending=False)
                    fig_top_models = px.bar(
                        top_models,
                        x="產品型號",
                        y="成交金額",
                        text="成交金額",
                        labels={
                            "產品型號": "產品型號",
                            "成交金額": "成交金額",
                        },
                    )
                    fig_top_models.update_xaxes(
                        tickangle=0,
                        automargin=True,
                    )
                    fig_top_models.update_traces(
                        texttemplate="%{text:,.0f}",
                        textposition="outside",
                        cliponaxis=False,
                    )
                    fig_top_models.update_layout(
                        xaxis_title="產品型號",
                        yaxis_title="成交金額",
                        showlegend=False,
                        margin=dict(b=90, t=20),
                    )
                    st.plotly_chart(
                        fig_top_models,
                        use_container_width=True,
                        key="top10_model_sales_bar",
                    )

        with owner_tab:
            owner_summary = (
                analysis_df.groupby("owner", dropna=False)
                .agg(
                    詢價件數=("RFQ_ID", "size"),
                    客戶數=("customer", lambda x: x[x.astype(str).str.strip() != ""].nunique()),
                    報價件數=("_quoted", "sum"),
                    報價金額=("報價金額", "sum"),
                    成交件數=("_won", "sum"),
                    成交金額=("_deal_amount", lambda x: x[analysis_df.loc[x.index, "_won"]].sum()),
                )
                .reset_index()
                .rename(columns={"owner": "業務"})
            )

            if owner_summary.empty:
                st.info("目前期間內沒有業務資料。")
            else:
                owner_summary["成交率 (%)"] = (
                    owner_summary["成交件數"]
                    / owner_summary["報價件數"].replace(0, pd.NA)
                    * 100
                ).fillna(0).round(1)

                # 平均報價時間
                cycle = analysis_df[
                    analysis_df["_created_dt"].notna()
                    & pd.to_datetime(analysis_df["首次報價日"], errors="coerce").notna()
                ].copy()
                if not cycle.empty:
                    cycle["_first_quote"] = pd.to_datetime(cycle["首次報價日"], errors="coerce")
                    cycle["_quote_days"] = (
                        cycle["_first_quote"] - cycle["_created_dt"]
                    ).dt.total_seconds() / 86400
                    cycle = cycle[cycle["_quote_days"] >= 0]
                    owner_cycle = cycle.groupby("owner")["_quote_days"].mean().round(1)
                    owner_summary["平均報價天數"] = owner_summary["業務"].map(owner_cycle)
                else:
                    owner_summary["平均報價天數"] = pd.NA

                owner_summary = owner_summary.sort_values(
                    ["成交金額", "成交件數"], ascending=False
                )
                st.markdown("#### 業務績效排名")
                st.dataframe(owner_summary, use_container_width=True, hide_index=True)

                st.markdown("#### 業務成交金額")
                st.bar_chart(
                    owner_summary.set_index("業務")["成交金額"],
                    use_container_width=True
                )

        with continent_tab:
            continent_summary = (
                analysis_df.groupby("continent", dropna=False)
                .agg(
                    詢價件數=("RFQ_ID", "size"),
                    客戶數=("customer", lambda x: x[x.astype(str).str.strip() != ""].nunique()),
                    報價件數=("_quoted", "sum"),
                    報價金額=("報價金額", "sum"),
                    成交件數=("_won", "sum"),
                    成交金額=("_deal_amount", lambda x: x[analysis_df.loc[x.index, "_won"]].sum()),
                )
                .reset_index()
                .rename(columns={"continent": "洲別"})
            )

            if continent_summary.empty:
                st.info("目前期間內沒有洲別資料。")
            else:
                continent_summary["成交率 (%)"] = (
                    continent_summary["成交件數"]
                    / continent_summary["報價件數"].replace(0, pd.NA)
                    * 100
                ).fillna(0).round(1)

                continent_summary = continent_summary.sort_values(
                    ["成交金額", "成交件數", "詢價件數"],
                    ascending=False
                )

                st.markdown("#### 洲別銷售績效")
                st.dataframe(
                    continent_summary,
                    use_container_width=True,
                    hide_index=True
                )

                st.markdown("#### 各洲成交金額")
                st.bar_chart(
                    continent_summary.set_index("洲別")["成交金額"],
                    use_container_width=True
                )

                st.markdown("#### 各洲詢價件數")
                st.bar_chart(
                    continent_summary.set_index("洲別")["詢價件數"],
                    use_container_width=True
                )

        with country_tab:
            country_summary = (
                analysis_df.groupby(["continent", "country"], dropna=False)
                .agg(
                    詢價件數=("RFQ_ID", "size"),
                    客戶數=("customer", lambda x: x[x.astype(str).str.strip() != ""].nunique()),
                    報價件數=("_quoted", "sum"),
                    報價金額=("報價金額", "sum"),
                    成交件數=("_won", "sum"),
                    成交金額=("_deal_amount", lambda x: x[analysis_df.loc[x.index, "_won"]].sum()),
                )
                .reset_index()
                .rename(columns={"continent": "洲別", "country": "國家"})
            )

            if country_summary.empty:
                st.info("目前期間內沒有國家資料。")
            else:
                country_summary["成交率 (%)"] = (
                    country_summary["成交件數"]
                    / country_summary["報價件數"].replace(0, pd.NA)
                    * 100
                ).fillna(0).round(1)

                country_summary = country_summary.sort_values(
                    ["成交金額", "成交件數", "詢價件數"],
                    ascending=False
                )

                st.markdown("#### 國家銷售績效")
                st.dataframe(
                    country_summary,
                    use_container_width=True,
                    hide_index=True
                )

                top_countries = country_summary[
                    country_summary["國家"] != "未分類"
                ].head(15)

                if not top_countries.empty:
                    st.markdown("#### Top 15 國家－成交金額")
                    st.bar_chart(
                        top_countries.set_index("國家")["成交金額"],
                        use_container_width=True
                    )

                    st.markdown("#### Top 15 國家－詢價件數")
                    st.bar_chart(
                        top_countries.set_index("國家")["詢價件數"],
                        use_container_width=True
                    )

        with country_purchase_tab:
            st.markdown("#### 🛒 國家採購分析")
            st.caption(
                "以「已成交」案件代表實際採購；可查看各國採購產品、採購次數、採購金額、產品組合與主要市場。"
            )

            purchase_df = analysis_df[
                (analysis_df["_won"])
                & (analysis_df["country"].astype(str).str.strip() != "")
                & (analysis_df["country"] != "未分類")
            ].copy()

            if purchase_df.empty:
                st.info("目前分析期間內尚無已成交且已填國家的採購資料。")
            else:
                # Country filter
                purchase_countries = sorted(
                    purchase_df["country"].dropna().astype(str).unique().tolist()
                )
                selected_purchase_countries = st.multiselect(
                    "篩選國家",
                    purchase_countries,
                    default=purchase_countries,
                    key="country_purchase_filter"
                )

                filtered_purchase = purchase_df[
                    purchase_df["country"].isin(selected_purchase_countries)
                ].copy()

                # Country-level purchasing summary
                country_purchase_summary = (
                    filtered_purchase.groupby(["continent", "country"], dropna=False)
                    .agg(
                        採購案件數=("RFQ_ID", "size"),
                        採購客戶數=("customer", lambda x: x[x.astype(str).str.strip() != ""].nunique()),
                        採購金額=("_deal_amount", "sum"),
                        採購產品種類=("_analysis_model", lambda x: x[x.astype(str).str.strip() != ""].nunique()),
                    )
                    .reset_index()
                    .rename(columns={"continent": "洲別", "country": "國家"})
                )

                country_purchase_summary["平均每案採購金額"] = (
                    country_purchase_summary["採購金額"]
                    / country_purchase_summary["採購案件數"].replace(0, pd.NA)
                ).fillna(0).round(0)

                country_purchase_summary = country_purchase_summary.sort_values(
                    ["採購金額", "採購案件數"], ascending=False
                )

                cp1, cp2, cp3, cp4 = st.columns(4)
                cp1.metric("採購國家數", f"{filtered_purchase['country'].nunique()} 國")
                cp2.metric("採購案件數", f"{len(filtered_purchase)} 件")
                cp3.metric("採購客戶數", f"{filtered_purchase['customer'].replace('', pd.NA).dropna().nunique()} 家")
                cp4.metric("採購總金額", money(float(filtered_purchase["_deal_amount"].sum())))

                st.markdown("##### 各國採購總覽")
                st.dataframe(
                    country_purchase_summary,
                    use_container_width=True,
                    hide_index=True
                )

                st.markdown("##### 各國採購金額排名")
                st.bar_chart(
                    country_purchase_summary.set_index("國家")["採購金額"],
                    use_container_width=True
                )

                # Country x model purchasing matrix/detail
                country_model_purchase = (
                    filtered_purchase.groupby(
                        ["continent", "country", "_analysis_category", "_analysis_model"],
                        dropna=False
                    )
                    .agg(
                        採購次數=("RFQ_ID", "size"),
                        採購金額=("_deal_amount", "sum"),
                        客戶數=("customer", lambda x: x[x.astype(str).str.strip() != ""].nunique()),
                    )
                    .reset_index()
                    .rename(
                        columns={
                            "continent": "洲別",
                            "country": "國家",
                            "_analysis_category": "產品分類",
                            "_analysis_model": "產品型號",
                        }
                    )
                )

                country_model_purchase = country_model_purchase[
                    country_model_purchase["產品型號"].astype(str).str.strip() != ""
                ].sort_values(
                    ["採購金額", "採購次數"], ascending=False
                )

                st.markdown("##### 國家 × 產品採購明細")
                st.dataframe(
                    country_model_purchase,
                    use_container_width=True,
                    hide_index=True
                )

                # Top product per country
                if not country_model_purchase.empty:
                    top_product_by_country = (
                        country_model_purchase.sort_values(
                            ["國家", "採購金額", "採購次數"],
                            ascending=[True, False, False]
                        )
                        .groupby("國家", as_index=False)
                        .first()
                    )

                    top_product_by_country = top_product_by_country[
                        ["國家", "產品分類", "產品型號", "採購次數", "採購金額"]
                    ]

                    st.markdown("##### 各國主要採購產品")
                    st.dataframe(
                        top_product_by_country,
                        use_container_width=True,
                        hide_index=True
                    )

                # Monthly purchasing trend by country
                trend_df = filtered_purchase[
                    filtered_purchase["_deal_dt"].notna()
                ].copy()

                if not trend_df.empty:
                    trend_df["月份"] = trend_df["_deal_dt"].dt.to_period("M").astype(str)
                    monthly_country_purchase = (
                        trend_df.groupby(["月份", "country"])["_deal_amount"]
                        .sum()
                        .reset_index()
                        .rename(columns={"country": "國家", "_deal_amount": "採購金額"})
                    )

                    pivot_purchase = monthly_country_purchase.pivot(
                        index="月份",
                        columns="國家",
                        values="採購金額"
                    ).fillna(0)

                    st.markdown("##### 各國每月採購趨勢")
                    st.line_chart(
                        pivot_purchase,
                        use_container_width=True
                    )

        # 未成交原因
        lost_df = analysis_df[analysis_df["sales_result"] == "未成交"].copy()
        if not lost_df.empty:
            st.divider()
            st.markdown("### ❌ 未成交原因分析")
            lost_reason = (
                lost_df["lost_reason"]
                .fillna("未填寫")
                .astype(str)
                .replace("", "未填寫")
                .value_counts()
                .rename_axis("未成交原因")
                .reset_index(name="案件數")
            )
            st.dataframe(lost_reason, use_container_width=True, hide_index=True)
            st.bar_chart(
                lost_reason.set_index("未成交原因")["案件數"],
                use_container_width=True
            )


# ################################################################
# ################################################################
# TAB 7：SALES INTELLIGENCE 銷售戰情室
# ################################################################

with tab_intelligence:

    # -------------------------------------------------------------
    # Dashboard styling
    # -------------------------------------------------------------
    st.markdown(
        """
        <style>
        .si-title {
            background: linear-gradient(90deg,#3730a3,#4f46e5);
            color:white; text-align:center; font-size:30px; font-weight:900;
            padding:15px 18px; border-radius:8px; margin-bottom:14px;
            letter-spacing:1px;
        }
        .si-card {
            background:linear-gradient(135deg,#ffffff,#f8fafc);
            border:1px solid #e2e8f0; border-radius:14px;
            padding:15px 16px; min-height:118px;
            box-shadow:0 2px 8px rgba(15,23,42,.08);
        }
        .si-card-title {font-size:15px;font-weight:800;color:#334155;}
        .si-card-value {font-size:28px;font-weight:900;color:#3730a3;margin-top:6px;}
        .si-card-sub {font-size:12px;color:#64748b;margin-top:5px;}
        .si-panel-title {
            background:linear-gradient(90deg,#3730a3,#4338ca);
            color:white; padding:9px 14px; border-radius:7px 7px 0 0;
            font-size:18px; font-weight:900; margin-top:6px;
        }
        </style>
        <div class="si-title">🎯 Sales Intelligence｜智慧銷售戰情室</div>
        """,
        unsafe_allow_html=True
    )

    intel = df.copy()
    intel_defaults = {
        "RFQ_ID": "", "customer": "", "title": "", "status": "",
        "owner": "", "created_time": "", "due_time": "",
        "first_followup_due": "", "product_category": "",
        "product_model": "", "continent": "", "country": "",
        "sales_result": "進行中", "deal_date": "",
        "deal_amount": 0.0, "lost_reason": ""
    }
    for col, default in intel_defaults.items():
        if col not in intel.columns:
            intel[col] = default

    intel["_created"] = pd.to_datetime(intel["created_time"], errors="coerce")
    intel["_due"] = pd.to_datetime(intel["due_time"], errors="coerce")
    intel["_followup"] = pd.to_datetime(intel["first_followup_due"], errors="coerce")
    intel["_deal_date"] = pd.to_datetime(intel["deal_date"], errors="coerce")
    intel["_deal_amount"] = pd.to_numeric(intel["deal_amount"], errors="coerce").fillna(0)
    intel["sales_result"] = intel["sales_result"].fillna("進行中").astype(str).replace("", "進行中")
    intel["country"] = intel["country"].fillna("").astype(str).replace("", "未分類")
    intel["continent"] = intel["continent"].fillna("").astype(str).replace("", "未分類")
    intel["product_model"] = intel["product_model"].fillna("").astype(str)
    intel["_product"] = intel["product_model"].where(
        intel["product_model"].str.strip() != "",
        intel["title"].fillna("").astype(str)
    )

    today_intel = pd.Timestamp.now().normalize()
    valid_created = intel["_created"].dropna()
    default_start = valid_created.min().date() if not valid_created.empty else today_intel.date()

    # -------------------------------------------------------------
    # Filters
    # -------------------------------------------------------------
    f1, f2, f3, f4 = st.columns([1,1,1.4,1.4])
    with f1:
        intel_start = st.date_input("起始日期", default_start, key="si_start")
    with f2:
        intel_end = st.date_input("結束日期", today_intel.date(), key="si_end")
    with f3:
        owner_opts = sorted([x for x in intel["owner"].dropna().astype(str).unique() if x.strip()])
        intel_owners = st.multiselect("業務", owner_opts, default=owner_opts, key="si_owner")
    with f4:
        # 國家篩選：改為單選下拉，避免 multiselect 預設全選時難以操作
        country_opts = sorted(
            {
                str(x).strip()
                for x in intel["country"].dropna().astype(str).tolist()
                if str(x).strip() and str(x).strip() != "未分類"
            }
        )
        intel_country = st.selectbox(
            "國家",
            ["全部國家"] + country_opts,
            index=0,
            key="si_country_filter"
        )

    if intel_start > intel_end:
        st.error("起始日期不可晚於結束日期。")
        view = intel.iloc[0:0].copy()
    else:
        s = pd.Timestamp(intel_start)
        e = pd.Timestamp(intel_end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        view = intel[intel["_created"].between(s, e, inclusive="both")].copy()
        if intel_owners:
            view = view[view["owner"].astype(str).isin(intel_owners)]
        if intel_country != "全部國家":
            view = view[
                view["country"].fillna("").astype(str).str.strip()
                == intel_country
            ].copy()

    view["_quoted"] = view["status"].astype(str).isin(["已報價", "追蹤中", "成交生產", "交貨", "收款", "結案"])
    view["_won"] = view["sales_result"].eq("已成交")
    view["_lost"] = view["sales_result"].eq("未成交")
    view["_active"] = ~view["sales_result"].isin(["已成交", "未成交"])
    view["_won_amount"] = view["_deal_amount"].where(view["_won"], 0)

    total = len(view)
    quoted = int(view["_quoted"].sum())
    won = int(view["_won"].sum())
    active = int(view["_active"].sum())
    lost_count = int(view["_lost"].sum())
    deal_amount = float(view["_won_amount"].sum())
    quote_rate = quoted / total * 100 if total else 0
    conversion = won / quoted * 100 if quoted else 0

    # -------------------------------------------------------------
    # KPI cards - reference-image style
    # -------------------------------------------------------------
    cards = [
        ("📥", "詢價案件", f"{total:,} 件", "期間內 RFQ 總量"),
        ("💰", "已報價案件", f"{quoted:,} 件", f"報價率 {quote_rate:.1f}%"),
        ("🔄", "進行中案件", f"{active:,} 件", "目前 Pipeline"),
        ("🏆", "已成交案件", f"{won:,} 件", f"成交率 {conversion:.1f}%"),
        ("💵", "成交金額", money(deal_amount), "已成交案件金額"),
        ("❌", "未成交案件", f"{lost_count:,} 件", "Lost Order"),
    ]
    card_cols = st.columns(6)
    for c, (icon, title, value, sub) in zip(card_cols, cards):
        with c:
            st.markdown(
                f"""
                <div class="si-card">
                    <div class="si-card-title">{icon} {title}</div>
                    <div class="si-card-value">{value}</div>
                    <div class="si-card-sub">{sub}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.write("")

    # -------------------------------------------------------------
    # First dashboard row: Funnel / Aging / detail
    # -------------------------------------------------------------
    col_left, col_mid, col_right = st.columns([1.05,1.25,1.7])

    with col_left:
        st.markdown('<div class="si-panel-title">銷售漏斗與流程分布</div>', unsafe_allow_html=True)

        funnel = pd.DataFrame([
            {"階段": status_name, "案件數": int((view["status"].astype(str) == status_name).sum())}
            for status_name in RFQ_STATUS
        ])
        funnel = funnel[funnel["案件數"] > 0]

        if funnel.empty:
            st.info("目前沒有流程資料。")
        else:
            fig_funnel = px.bar(
                funnel,
                x="案件數",
                y="階段",
                orientation="h",
                text="案件數",
            )
            fig_funnel.update_layout(
                height=300,
                margin=dict(l=10,r=10,t=20,b=10),
                showlegend=False,
                yaxis=dict(categoryorder="array", categoryarray=list(reversed(RFQ_STATUS))),
            )
            fig_funnel.update_traces(textposition="outside")
            st.plotly_chart(fig_funnel, use_container_width=True)

        # Pipeline composition
        comp = pd.DataFrame({
            "狀態": ["進行中","已成交","未成交"],
            "案件數": [active, won, lost_count]
        })
        fig_comp = px.pie(comp, names="狀態", values="案件數", hole=0.35)
        fig_comp.update_layout(height=270, margin=dict(l=5,r=5,t=20,b=5), legend_orientation="h")
        st.plotly_chart(fig_comp, use_container_width=True)

    with col_mid:
        st.markdown('<div class="si-panel-title">RFQ Aging 與業務績效</div>', unsafe_allow_html=True)

        aging = view[view["_active"] & view["_created"].notna()].copy()
        aging["_days"] = (today_intel - aging["_created"].dt.normalize()).dt.days.clip(lower=0)

        def _aging_bucket(d):
            if d <= 7: return "0–7 天"
            if d <= 14: return "8–14 天"
            if d <= 30: return "15–30 天"
            if d <= 60: return "31–60 天"
            return "60+ 天"

        if not aging.empty:
            aging["區間"] = aging["_days"].apply(_aging_bucket)
            order = ["0–7 天","8–14 天","15–30 天","31–60 天","60+ 天"]
            aging_sum = (
                aging["區間"].value_counts()
                .reindex(order, fill_value=0)
                .rename_axis("案件老化")
                .reset_index(name="案件數")
            )
            fig_age = px.bar(aging_sum, x="案件老化", y="案件數", text="案件數")
            fig_age.update_layout(height=275, margin=dict(l=10,r=10,t=20,b=10), showlegend=False)
            st.plotly_chart(fig_age, use_container_width=True)
        else:
            st.info("目前沒有進行中 Aging 資料。")

        owners = view.copy()
        owners["owner"] = owners["owner"].fillna("待確認").astype(str).replace("", "待確認")
        if not owners.empty:
            owner_sum = owners.groupby("owner").agg(
                詢價=("RFQ_ID","size"),
                報價=("_quoted","sum"),
                成交=("_won","sum"),
                成交金額=("_won_amount","sum")
            ).reset_index().rename(columns={"owner":"業務"})
            owner_sum["成交率"] = (
                owner_sum["成交"] / owner_sum["報價"].replace(0,pd.NA) * 100
            ).fillna(0).round(1)

            fig_owner = go.Figure()
            fig_owner.add_bar(
                x=owner_sum["業務"],
                y=owner_sum["詢價"],
                name="詢價件數"
            )
            fig_owner.add_trace(go.Scatter(
                x=owner_sum["業務"],
                y=owner_sum["成交率"],
                name="成交率 %",
                mode="lines+markers",
                yaxis="y2"
            ))
            fig_owner.update_layout(
                height=290,
                margin=dict(l=10,r=10,t=20,b=10),
                yaxis=dict(title="案件數"),
                yaxis2=dict(title="成交率 %", overlaying="y", side="right", range=[0,100]),
                legend=dict(orientation="h")
            )
            st.plotly_chart(fig_owner, use_container_width=True)

    with col_right:
        st.markdown('<div class="si-panel-title">銷售戰情明細</div>', unsafe_allow_html=True)

        detail = view.copy()
        detail["建立日"] = detail["_created"].dt.strftime("%Y-%m-%d")
        detail["成交金額"] = detail["_deal_amount"].round(0)
        detail["報價"] = detail["_quoted"].map({True:"是", False:"否"})
        detail["成交"] = detail["_won"].map({True:"是", False:"否"})
        detail_show = detail[
            ["RFQ_ID","customer","_product","country","owner","status",
             "報價","成交","成交金額","建立日"]
        ].copy()
        detail_show.columns = [
            "RFQ ID","客戶","產品","國家","業務","狀態",
            "已報價","已成交","成交金額","建立日"
        ]
        st.dataframe(
            detail_show.sort_values("建立日", ascending=False),
            use_container_width=True,
            hide_index=True,
            height=590
        )

    # -------------------------------------------------------------
    # Second row: country/product opportunity + overdue detail
    # -------------------------------------------------------------
    market_col, alert_col = st.columns([1.45,1])

    with market_col:
        st.markdown('<div class="si-panel-title">國家 × 產品市場機會</div>', unsafe_allow_html=True)

        market = view[
            (view["country"] != "未分類") &
            (view["_product"].astype(str).str.strip() != "")
        ].copy()

        if market.empty:
            st.info("尚無足夠的國家 / 產品資料。")
        else:
            ms = market.groupby(["country","_product"]).agg(
                詢價件數=("RFQ_ID","size"),
                報價件數=("_quoted","sum"),
                成交件數=("_won","sum"),
                成交金額=("_won_amount","sum")
            ).reset_index().rename(columns={"country":"國家","_product":"產品"})
            ms["成交率 (%)"] = (
                ms["成交件數"] / ms["報價件數"].replace(0,pd.NA) * 100
            ).fillna(0).round(1)
            ms = ms.sort_values(["成交金額","詢價件數"], ascending=False)

            top_ms = ms.head(12)
            fig_market = px.bar(
                top_ms,
                x="產品",
                y="詢價件數",
                color="國家",
                text="詢價件數",
                barmode="group"
            )
            fig_market.update_layout(
                height=330,
                margin=dict(l=10,r=10,t=20,b=10),
                legend=dict(orientation="h")
            )
            st.plotly_chart(fig_market, use_container_width=True)
            st.dataframe(ms.head(20), use_container_width=True, hide_index=True)

    with alert_col:
        st.markdown('<div class="si-panel-title">需要立即處理</div>', unsafe_allow_html=True)

        attention = view[view["_active"]].copy()
        attention["_alert"] = attention["_followup"].fillna(attention["_due"])
        attention["_overdue"] = (
            today_intel - attention["_alert"].dt.normalize()
        ).dt.days
        attention = attention[
            attention["_alert"].notna() & (attention["_overdue"] > 0)
        ].sort_values("_overdue", ascending=False)

        if attention.empty:
            st.success("目前沒有逾期追蹤案件。")
        else:
            at = attention[
                ["RFQ_ID","customer","_product","owner","status","_alert","_overdue"]
            ].copy()
            at.columns = ["RFQ ID","客戶","產品","業務","狀態","應追蹤日","逾期天數"]
            at["應追蹤日"] = at["應追蹤日"].dt.strftime("%Y-%m-%d")
            st.metric("逾期案件", f"{len(at)} 件")
            st.dataframe(at, use_container_width=True, hide_index=True, height=430)

    # -------------------------------------------------------------
    # Lost order
    # -------------------------------------------------------------
    st.markdown('<div class="si-panel-title">Lost Order 失單情報</div>', unsafe_allow_html=True)
    lost = view[view["_lost"]].copy()

    if lost.empty:
        st.info("目前期間沒有未成交案件。")
    else:
        lost["lost_reason"] = lost["lost_reason"].fillna("未填寫").astype(str).replace("", "未填寫")
        lost_sum = (
            lost["lost_reason"].value_counts()
            .rename_axis("未成交原因")
            .reset_index(name="案件數")
        )
        lc1, lc2 = st.columns([1,2])
        with lc1:
            st.dataframe(lost_sum, use_container_width=True, hide_index=True)
        with lc2:
            fig_lost = px.bar(
                lost_sum,
                x="未成交原因",
                y="案件數",
                text="案件數"
            )
            fig_lost.update_layout(height=300, margin=dict(l=10,r=10,t=20,b=10), showlegend=False)
            st.plotly_chart(fig_lost, use_container_width=True)


# ################################################################
# TAB 7：STATUS LOG
# ################################################################

with tab_log:

    st.subheader(
        "📜 RFQ 狀態異動紀錄"
    )


    try:

        log_df = conn.read(

            worksheet="StatusLog",

            ttl=GSHEETS_READ_TTL

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
