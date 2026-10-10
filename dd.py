"""JENN-WEI CNC / special-purpose machine shipping management (LAN pilot)."""
import os
import io
import csv
import uuid
from datetime import date, datetime, timezone
from decimal import Decimal

import pandas as pd
import streamlit as st
from sqlalchemy import create_engine, text
from sqlalchemy.exc import IntegrityError

st.set_page_config(page_title="JENN-WEI 出貨管理", page_icon="🚚", layout="wide")

STAGES = ["建立出貨單", "出貨準備", "品質檢驗", "待核准", "待出貨", "運送中", "客戶簽收", "完成結案", "取消"]
TYPES = ["CNC 車床", "CNC 車銑複合機", "CNC 多軸加工機", "專用機", "其他"]
CHECKS = [
    ("inspection", "出廠檢驗與試車完成"),
    ("accuracy", "精度／功能測試紀錄確認"),
    ("accessories", "附件、工具及備品核對"),
    ("manual", "操作手冊／電路圖／相關文件齊全"),
    ("photos", "出貨前機台照片存檔"),
    ("packing", "防鏽、固定與包裝完成"),
    ("lifting", "吊裝／運輸安全確認"),
    ("payment", "出貨付款條件確認"),
]

@st.cache_resource
def engine(url):
    return create_engine(url, pool_pre_ping=True, pool_size=5, max_overflow=10)


def db():
    url = os.getenv("DATABASE_URL", "")
    if not url:
        try:
            url = str(st.secrets.get("DATABASE_URL", ""))
        except Exception:
            pass
    if not url:
        st.error("請先設定 DATABASE_URL（PostgreSQL 連線字串）。")
        st.stop()
    return engine(url)


def rows(sql, params=None):
    with db().connect() as c:
        return [dict(r) for r in c.execute(text(sql), params or {}).mappings()]


def execute(sql, params=None):
    with db().begin() as c:
        return c.execute(text(sql), params or {})


def audit(shipment_id, event, detail=""):
    execute("INSERT INTO shipment_audit (shipment_id, event, detail, actor) VALUES (:id,:event,:detail,:actor)",
            {"id": shipment_id, "event": event, "detail": detail, "actor": st.session_state.get("operator", "未指定")})


def read_master(sheet):
    try:
        data = rows("SELECT rows FROM worksheet_store WHERE sheet_name=:sheet", {"sheet": sheet})
        return data[0]["rows"] if data else []
    except Exception:
        return []


def customer_names():
    return sorted({str(x.get("name") or x.get("customer") or "").strip() for x in read_master("Customers") if x.get("name") or x.get("customer")})


def employee_names():
    return sorted({str(x.get("name") or "").strip() for x in read_master("Employees") if x.get("name") and str(x.get("status") or "在職").strip() not in ("離職", "留職停薪", "停用")})


def pick(label, options, key, value=""):
    opts = ["✍️ 手動輸入"] + [x for x in options if x]
    if value and value not in opts:
        opts.append(value)
    choice = st.selectbox(label, opts, index=opts.index(value) if value in opts else 0, key=f"{key}_pick")
    return st.text_input(label + "（手動輸入）", value=value, key=f"{key}_text") if choice == "✍️ 手動輸入" else choice


def dstr(value):
    return value.isoformat() if isinstance(value, (date, datetime)) else (str(value) if value is not None else "")


def money(value):
    return f"{float(value or 0):,.0f}"

st.title("🚚 JENN-WEI CNC 機械與專用機出貨管理")
st.caption("內網 PostgreSQL 版本 · 出貨單、檢查表、流程管制、物流與簽收紀錄")
st.sidebar.text_input("操作人員（測試用）", key="operator", value="出貨承辦")
st.sidebar.warning("此欄位不是身分驗證。正式上線須接入登入與權限系統。")

try:
    shipments = rows("SELECT * FROM shipments ORDER BY created_at DESC LIMIT 1000")
except Exception as exc:
    st.error(f"無法連線出貨資料表：{exc}。請先執行 sql/02_shipments.sql")
    st.stop()

open_count = sum(s["status"] not in ("完成結案", "取消") for s in shipments)
c1,c2,c3,c4 = st.columns(4)
c1.metric("出貨單總數", len(shipments))
c2.metric("進行中", open_count)
c3.metric("待出貨", sum(s["status"] == "待出貨" for s in shipments))
c4.metric("已完成", sum(s["status"] == "完成結案" for s in shipments))

tab1, tab2, tab3, tab4 = st.tabs(["📋 出貨總覽", "➕ 新增出貨單", "🛠️ 出貨作業", "📈 紀錄與匯出"])

with tab1:
    q = st.text_input("搜尋：出貨單號／客戶／機型／機台序號／訂單號")
    status_filter = st.multiselect("篩選狀態", STAGES)
    filtered = [s for s in shipments if (not status_filter or s["status"] in status_filter) and
                (not q or any(q.lower() in str(s.get(k) or "").lower() for k in ("shipment_no","customer_name","machine_model","serial_number","sales_order_no")))]
    cols = ["shipment_no","status","machine_type","customer_name","machine_model","serial_number","quantity","planned_ship_date","actual_ship_date","sales_order_no","owner"]
    if filtered:
        st.dataframe(pd.DataFrame(filtered)[cols].rename(columns={"shipment_no":"出貨單號","status":"狀態","machine_type":"機種","customer_name":"客戶","machine_model":"機型","serial_number":"機台序號","quantity":"數量","planned_ship_date":"預計出貨","actual_ship_date":"實際出貨","sales_order_no":"訂單號","owner":"負責人"}), use_container_width=True, hide_index=True)
    else:
        st.info("沒有符合條件的出貨單。")

with tab2:
    st.subheader("建立出貨單")
    with st.form("new_shipment", clear_on_submit=False):
        left,right = st.columns(2)
        with left:
            customer = st.selectbox("客戶主檔", ["✍️ 手動輸入"] + customer_names())
            customer_manual = st.text_input("客戶名稱（手動輸入）") if customer == "✍️ 手動輸入" else ""
            machine_type = st.selectbox("機械類別", TYPES)
            model = st.text_input("機型／設備名稱 *")
            serial = st.text_input("機台序號 *")
            qty = st.number_input("數量", min_value=1, value=1, step=1)
            order = st.text_input("訂單／合約編號")
        with right:
            owner = st.selectbox("出貨負責人", ["✍️ 手動輸入"] + employee_names())
            owner_manual = st.text_input("負責人（手動輸入）") if owner == "✍️ 手動輸入" else ""
            planned = st.date_input("預計出貨日期", value=date.today())
            destination = st.text_input("交貨地址／目的地")
            incoterms = st.text_input("交易條件（例如 EXW、FOB）")
            carrier = st.text_input("物流／承運商")
            notes = st.text_area("備註／特殊出貨需求")
        submitted = st.form_submit_button("建立出貨單", type="primary")
    if submitted:
        cname = (customer_manual if customer == "✍️ 手動輸入" else customer).strip()
        oname = (owner_manual if owner == "✍️ 手動輸入" else owner).strip()
        if not cname or not model.strip() or not serial.strip():
            st.error("客戶、機型、機台序號為必填。")
        else:
            new_id = str(uuid.uuid4())
            number = "SHP-" + datetime.now().strftime("%Y%m%d") + "-" + uuid.uuid4().hex[:6].upper()
            try:
                execute("""INSERT INTO shipments (id,shipment_no,customer_name,machine_type,machine_model,serial_number,quantity,sales_order_no,owner,planned_ship_date,destination,incoterms,carrier,notes)
                        VALUES (:id,:number,:customer,:type,:model,:serial,:qty,:ord,:owner,:planned,:dest,:terms,:carrier,:notes)""",
                        {"id":new_id,"number":number,"customer":cname,"type":machine_type,"model":model.strip(),"serial":serial.strip(),"qty":int(qty),"ord":order,"owner":oname,"planned":planned,"dest":destination,"terms":incoterms,"carrier":carrier,"notes":notes})
                audit(new_id,"建立出貨單",number)
                st.success(f"已建立 {number}")
                st.rerun()
            except Exception as exc:
                st.error(f"建立失敗：{exc}")

with tab3:
    if not shipments:
        st.info("請先建立出貨單。")
    else:
        by_number = {s["shipment_no"]:s for s in shipments}
        chosen = st.selectbox("選擇出貨單", list(by_number))
        current = by_number[chosen]
        sid = current["id"]
        st.write(f"**客戶：** {current['customer_name']}　**機型：** {current['machine_model']}　**機台序號：** {current['serial_number']}")
        st.write(f"**目前狀態：** {current['status']}　**資料版本：** {current['version']}")
        checks = rows("SELECT item_key,checked,checked_by,checked_at FROM shipment_checks WHERE shipment_id=:id", {"id":sid})
        checked = {x["item_key"]:x for x in checks}
        with st.form(f"checks_{sid}"):
            st.markdown("#### 出貨檢查表")
            values = {key:st.checkbox(label,value=bool(checked.get(key,{}).get("checked",False)),key=f"chk_{sid}_{key}") for key,label in CHECKS}
            if st.form_submit_button("儲存檢查表"):
                with db().begin() as conn:
                    for key,val in values.items():
                        conn.execute(text("""INSERT INTO shipment_checks (shipment_id,item_key,checked,checked_by,checked_at) VALUES (:id,:key,:checked,:by,CASE WHEN :checked THEN NOW() ELSE NULL END)
                            ON CONFLICT (shipment_id,item_key) DO UPDATE SET checked=EXCLUDED.checked,checked_by=EXCLUDED.checked_by,checked_at=EXCLUDED.checked_at"""),
                            {"id":sid,"key":key,"checked":val,"by":st.session_state.operator})
                audit(sid,"更新出貨檢查表",f"完成 {sum(values.values())}/{len(CHECKS)} 項")
                st.success("檢查表已儲存")
                st.rerun()
        st.progress(sum(bool(checked.get(key,{}).get("checked",False)) for key,_ in CHECKS)/len(CHECKS),text=f"出貨檢查完成 {sum(bool(checked.get(key,{}).get('checked',False)) for key,_ in CHECKS)}/{len(CHECKS)}")
        with st.form(f"workflow_{sid}"):
            st.markdown("#### 出貨狀態與物流資訊")
            status = st.selectbox("更新狀態", STAGES, index=STAGES.index(current["status"]))
            left,right = st.columns(2)
            with left:
                carrier2 = st.text_input("承運商",value=current["carrier"] or "")
                tracking = st.text_input("貨運單號",value=current["tracking_no"] or "")
                actual = st.date_input("實際出貨日期",value=current["actual_ship_date"] or date.today())
            with right:
                signed = st.text_input("客戶簽收人",value=current["received_by"] or "")
                received = st.date_input("簽收日期",value=current["received_date"] or date.today())
                remarks = st.text_area("處理備註",value=current["notes"] or "")
            save = st.form_submit_button("儲存出貨進度",type="primary")
        if save:
            done_checks = {k for k,v in checked.items() if v["checked"]}
            all_done = all(k in done_checks for k,_ in CHECKS)
            if status in ("待出貨","運送中","客戶簽收","完成結案") and not all_done:
                st.error("檢查表尚未全數完成，禁止進入待出貨及後續階段。請先儲存所有檢查項目。")
            elif status in ("客戶簽收","完成結案") and not signed.strip():
                st.error("客戶簽收或結案前，必須填寫簽收人。")
            elif status == "完成結案" and not tracking.strip():
                st.error("結案前請填寫貨運單號（自運案件可填寫內部運送單號）。")
            else:
                result = execute("""UPDATE shipments SET status=:status,carrier=:carrier,tracking_no=:tracking,
                     actual_ship_date=:actual,received_by=:signed,received_date=:received,notes=:notes,
                     version=version+1,updated_at=NOW() WHERE id=:id AND version=:version""",
                     {"status":status,"carrier":carrier2,"tracking":tracking,"actual":actual if status in ("運送中","客戶簽收","完成結案") else current["actual_ship_date"],
                      "signed":signed,"received":received if status in ("客戶簽收","完成結案") else current["received_date"],"notes":remarks,"id":sid,"version":current["version"]})
                if result.rowcount != 1:
                    st.error("資料已被其他使用者修改，請重新整理後再操作。")
                else:
                    audit(sid,"更新出貨進度",f"{current['status']} → {status}")
                    st.success("已儲存")
                    st.rerun()

with tab4:
    if shipments:
        report = pd.DataFrame(shipments)
        st.bar_chart(report.groupby("status").size().reindex(STAGES,fill_value=0),horizontal=True)
        export = report.drop(columns=["id"],errors="ignore").copy()
        for c in export.columns:
            export[c] = export[c].map(dstr)
        st.download_button("下載出貨清單 CSV (UTF-8 BOM)", export.to_csv(index=False).encode("utf-8-sig"),"JENNWEI_shipments.csv","text/csv")
        selected_log = st.selectbox("檢視操作紀錄",[s["shipment_no"] for s in shipments],key="log_select")
        sid_log = next(s["id"] for s in shipments if s["shipment_no"] == selected_log)
        history = rows("SELECT created_at,actor,event,detail FROM shipment_audit WHERE shipment_id=:id ORDER BY created_at DESC",{"id":sid_log})
        st.dataframe(pd.DataFrame(history),use_container_width=True,hide_index=True)
    else:
        st.info("尚無出貨資料。")

