"""JENN-WEI 出貨管理：與 AA／BB／CC 共用 Google Sheets。
部署：保留現有 requirements.txt，加入 gspread>=5.8.0,<6。
複製 BB 的 [connections.gsheets] service_account Secrets 及 SERVICE_SPREADSHEET_URL（若有）。
自動建立 Shipments、ShipmentChecks、ShipmentAudit、WarrantyRecords；既有 PostgreSQL 資料不會自動遷移。
客戶優先沿用 shared_contacts.load_master；沒有該模組則讀 Customers.company_name。
Sheets 版本檢查不具跨部署原子鎖，請避免同時修改同一單據或手動排序資料分頁。
"""
import threading
from zoneinfo import ZoneInfo
import uuid
from datetime import date, datetime, timezone

import pandas as pd
import streamlit as st
import gspread
from gspread.exceptions import WorksheetNotFound

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

SHIP_COLS = ["id", "shipment_no", "customer_name", "machine_type", "machine_model", "serial_number", "quantity", "sales_order_no", "owner", "planned_ship_date", "destination", "incoterms", "carrier", "notes", "status", "tracking_no", "actual_ship_date", "received_by", "received_date", "version", "created_at", "updated_at"]
CHECK_COLS = ["shipment_id", "item_key", "checked", "checked_by", "checked_at"]
AUDIT_COLS = ["id", "shipment_id", "event", "detail", "actor", "created_at"]
WARRANTY_COLS = ["id", "shipment_id", "shipment_no", "customer_name", "machine_type", "machine_model", "serial_number", "sales_order_no", "owner", "actual_ship_date", "received_date", "start_basis", "start_date", "end_date", "terms", "notes", "version", "created_at", "updated_at"]
SCHEMAS = {"WarrantyRecords": WARRANTY_COLS, "Shipments": SHIP_COLS, "ShipmentChecks": CHECK_COLS, "ShipmentAudit": AUDIT_COLS}


def now():
    return datetime.now(ZoneInfo("Asia/Taipei")).isoformat(timespec="seconds")


def today():
    return datetime.now(ZoneInfo("Asia/Taipei")).date()


@st.cache_resource
def connect_sheet():
    config = dict(st.secrets["connections"]["gsheets"])
    if config.get("type") != "service_account":
        raise ValueError("請複製 AA／BB／CC 的 service_account Google Sheets Secrets；公開唯讀連線不能儲存。")
    target = str(st.secrets.get("SERVICE_SPREADSHEET_URL", "") or config.get("spreadsheet", "")).strip()
    config.pop("spreadsheet", None)
    if not target:
        raise ValueError("connections.gsheets.spreadsheet 尚未設定。")
    config.pop("worksheet", None)
    allowed = {"type", "project_id", "private_key_id", "private_key", "client_email", "client_id", "auth_uri", "token_uri", "auth_provider_x509_cert_url", "client_x509_cert_url", "universe_domain"}
    credentials = {k: v for k, v in config.items() if k in allowed}
    client = gspread.service_account_from_dict(credentials)
    if target.startswith("https://"):
        book = client.open_by_url(target)
    elif " " not in target and len(target) >= 25:
        # Try an ID first, then a title if the configured value is a long title.
        try:
            book = client.open_by_key(target)
        except gspread.exceptions.SpreadsheetNotFound:
            book = client.open(target)
    else:
        book = client.open(target)
    return book, threading.RLock()


def worksheet(name):
    book, _ = connect_sheet()
    try:
        sheet = book.worksheet(name)
    except WorksheetNotFound:
        if name not in SCHEMAS:
            raise
        sheet = book.add_worksheet(title=name, rows=1000, cols=len(SCHEMAS[name]))
    if name in SCHEMAS:
        header = sheet.row_values(1)
        if not header:
            sheet.append_row(SCHEMAS[name], value_input_option="RAW")
        elif header != SCHEMAS[name]:
            raise ValueError(f"{name} 欄位不符，請核對 dd.py 的 SCHEMAS 欄位定義；程式不會覆蓋既有欄位。")
    return sheet


def records(name):
    sheet = worksheet(name)
    values = sheet.get_all_values()
    if not values:
        return []
    header = values[0]
    result = []
    for rownum, cells in enumerate(values[1:], start=2):
        if not any(cells):
            continue
        row = dict(zip(header, cells + [""] * max(0, len(header) - len(cells))))
        row["_row"] = rownum
        if name == "Shipments":
            row["quantity"] = int(float(row.get("quantity") or 1))
            row["version"] = int(float(row.get("version") or 1))
            for key in ("planned_ship_date", "actual_ship_date", "received_date"):
                row[key] = date.fromisoformat(row[key][:10]) if row.get(key) else None
        if name == "ShipmentChecks":
            row["checked"] = str(row.get("checked", "")).lower() in ("true", "1", "yes")
        result.append(row)
    return result


def encode(row, columns):
    return [dstr(row.get(k, "")) for k in columns]


def audit(shipment_id, event, detail=""):
    row = {"id": str(uuid.uuid4()), "shipment_id": shipment_id, "event": event, "detail": detail,
           "actor": st.session_state.get("operator", "未指定"), "created_at": now()}
    try:
        worksheet("ShipmentAudit").append_row(encode(row, AUDIT_COLS), value_input_option="RAW")
    except Exception:
        st.warning("主要資料已儲存，但操作紀錄寫入失敗，請核對 ShipmentAudit 分頁與權限。")


def add_shipment(row):
    row.update(status="建立出貨單", version=1, created_at=now(), updated_at=now())
    worksheet("Shipments").append_row(encode(row, SHIP_COLS), value_input_option="RAW")


def save_checks(sid, values):
    _, lock = connect_sheet()
    with lock:
        existing = {r["item_key"]: r for r in records("ShipmentChecks") if r["shipment_id"] == sid}
        sheet = worksheet("ShipmentChecks")
        updates, additions = [], []
        for key, val in values.items():
            row = {"shipment_id": sid, "item_key": key, "checked": val,
                   "checked_by": st.session_state.operator, "checked_at": now() if val else ""}
            cells = encode(row, CHECK_COLS)
            if key in existing:
                n = existing[key]["_row"]
                updates.append({"range": f"A{n}:E{n}", "values": [cells]})
            else:
                additions.append(cells)
        if updates:
            sheet.batch_update(updates, value_input_option="RAW")
        if additions:
            sheet.append_rows(additions, value_input_option="RAW")


def update_shipment(sid, version, changes):
    _, lock = connect_sheet()
    with lock:
        latest = next((r for r in records("Shipments") if r["id"] == sid), None)
        if latest is None or latest["version"] != version:
            return False
        latest.update(changes, version=version + 1, updated_at=now())
        n = latest["_row"]
        worksheet("Shipments").update(range_name=f"A{n}:V{n}", values=[encode(latest, SHIP_COLS)], value_input_option="RAW")
        return True


SHIPPED_STAGES = ("運送中", "客戶簽收", "完成結案")


def ensure_warranties(shipments):
    """Append missing records; never replace confirmed warranty terms on retries."""
    _, lock = connect_sheet()
    with lock:
        existing = {r["shipment_id"] for r in records("WarrantyRecords")}
        additions = []
        for shipment in shipments:
            if shipment["status"] not in SHIPPED_STAGES or not shipment.get("actual_ship_date") or shipment["id"] in existing:
                continue
            row = {k: shipment.get(k, "") for k in ("shipment_no", "customer_name", "machine_type", "machine_model", "serial_number", "sales_order_no", "owner", "actual_ship_date", "received_date")}
            row.update(id="WAR-" + shipment["id"], shipment_id=shipment["id"], start_basis="待確認", version=1, created_at=now(), updated_at=now())
            additions.append(encode(row, WARRANTY_COLS))
            existing.add(shipment["id"])
        if additions:
            worksheet("WarrantyRecords").append_rows(additions, value_input_option="RAW")
        return len(additions)


def update_warranty(wid, version, changes):
    _, lock = connect_sheet()
    with lock:
        current = next((r for r in records("WarrantyRecords") if r["id"] == wid), None)
        if current is None or int(current.get("version") or 1) != version:
            return False
        current.update(changes, version=version + 1, updated_at=now())
        n = current["_row"]
        worksheet("WarrantyRecords").update(range_name=f"A{n}:S{n}", values=[encode(current, WARRANTY_COLS)], value_input_option="RAW")
        return True


def warranty_state(row, shipment_status):
    if shipment_status == "取消":
        return "出貨已取消／待人工確認"
    if shipment_status not in SHIPPED_STAGES:
        return "出貨狀態異動／待人工確認"
    if row.get("start_basis") == "待確認" or not row.get("start_date") or not row.get("end_date"):
        return "待確認保固條件"
    try:
        start, end = date.fromisoformat(row["start_date"]), date.fromisoformat(row["end_date"])
    except ValueError:
        return "日期格式錯誤"
    if end < start:
        return "日期範圍錯誤"
    if today() < start:
        return "尚未起算"
    return "保固中" if today() <= end else "已到期"


@st.cache_data(ttl=60)
def read_master(sheet):
    try:
        return records(sheet)
    except WorksheetNotFound:
        return []
    except Exception:
        st.warning(f"{sheet} 主檔讀取失敗，暫時可手動輸入；請檢查分頁及權限。")
        return []


def customer_names():
    # Use the same shared_contacts module as BB when it is deployed alongside DD.
    try:
        from shared_contacts import load_master
    except ModuleNotFoundError as exc:
        if exc.name != "shared_contacts":
            raise
        data = read_master("Customers")
    else:
        master = load_master("customer")
        data = [] if master is None else master.fillna("").to_dict("records")
    return sorted({str(x.get("company_name") or x.get("name") or x.get("customer") or x.get("customer_name") or x.get("客戶名稱") or "").strip() for x in data if str(x.get("status") or "").strip() != "停用"} - {""})


def employee_names():
    return sorted({str(x.get("name") or x.get("owner") or x.get("姓名") or "").strip() for x in read_master("Employees") if str(x.get("status") or "在職").strip() not in ("離職", "留職停薪", "停用")} - {""})


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
st.caption("共用 Google Sheets 版本 · 出貨單、檢查表、流程管制、物流與簽收紀錄")
st.sidebar.caption("多人操作請避免同時編輯同一張出貨單；Google Sheets 不提供資料庫交易鎖。")
st.sidebar.text_input("操作人員（測試用）", key="operator", value="出貨承辦")
st.sidebar.warning("此欄位不是身分驗證。正式上線須接入登入與權限系統。")

try:
    shipments = sorted(records("Shipments"), key=lambda x: x.get("created_at", ""), reverse=True)
    worksheet("ShipmentChecks")
    worksheet("ShipmentAudit")
except Exception as exc:
    st.error("無法讀取 Google Sheets。請確認 DD 的 connections.gsheets 與 AA／BB／CC 相同，且服務帳號具編輯權限。")
    st.info("請核對 Shipments、ShipmentChecks、ShipmentAudit 欄位；完整錯誤請查看 Streamlit Cloud 日誌。")
    import logging
    logging.exception("DD Google Sheets initialization failed")
    st.stop()

open_count = sum(s["status"] not in ("完成結案", "取消") for s in shipments)
c1,c2,c3,c4 = st.columns(4)
c1.metric("出貨單總數", len(shipments))
c2.metric("進行中", open_count)
c3.metric("待出貨", sum(s["status"] == "待出貨" for s in shipments))
c4.metric("已完成", sum(s["status"] == "完成結案" for s in shipments))

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📋 出貨總覽", "➕ 新增出貨單", "🛠️ 出貨作業", "📈 紀錄與匯出", "🛡️ 保固紀錄"])

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
            customer_manual = st.text_input("客戶名稱（選手動輸入時填寫）")
            machine_type = st.selectbox("機械類別", TYPES)
            model = st.text_input("機型／設備名稱 *")
            serial = st.text_input("機台序號 *")
            qty = st.number_input("數量", min_value=1, value=1, step=1)
            order = st.text_input("訂單／合約編號")
        with right:
            owner = st.selectbox("出貨負責人", ["✍️ 手動輸入"] + employee_names())
            owner_manual = st.text_input("負責人（選手動輸入時填寫）")
            planned = st.date_input("預計出貨日期", value=today())
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
            number = "SHP-" + datetime.now(ZoneInfo("Asia/Taipei")).strftime("%Y%m%d") + "-" + uuid.uuid4().hex[:6].upper()
            try:
                add_shipment({"id":new_id,"shipment_no":number,"customer_name":cname,"machine_type":machine_type,"machine_model":model.strip(),"serial_number":serial.strip(),"quantity":int(qty),"sales_order_no":order,"owner":oname,"planned_ship_date":planned,"destination":destination,"incoterms":incoterms,"carrier":carrier,"notes":notes})
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
        checks = [r for r in records("ShipmentChecks") if r["shipment_id"] == sid]
        checked = {x["item_key"]:x for x in checks}
        with st.form(f"checks_{sid}"):
            st.markdown("#### 出貨檢查表")
            values = {key:st.checkbox(label,value=bool(checked.get(key,{}).get("checked",False)),key=f"chk_{sid}_{key}") for key,label in CHECKS}
            if st.form_submit_button("儲存檢查表"):
                try:
                    save_checks(sid, values)
                except Exception:
                    st.error("檢查表儲存失敗，請重新整理並核對資料後重試。")
                    st.stop()
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
                actual = st.date_input("實際出貨日期",value=current["actual_ship_date"] or today())
            with right:
                signed = st.text_input("客戶簽收人",value=current["received_by"] or "")
                received = st.date_input("簽收日期",value=current["received_date"] or today())
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
                try:
                    fresh_checks = {r["item_key"] for r in records("ShipmentChecks") if r["shipment_id"] == sid and r["checked"]}
                    if status in ("待出貨", "運送中", "客戶簽收", "完成結案") and not all(k in fresh_checks for k, _ in CHECKS):
                        st.error("檢查表已變更，請重新確認所有檢查項目。")
                        st.stop()
                    result = update_shipment(sid, current["version"], {"status":status,"carrier":carrier2,"tracking_no":tracking,"actual_ship_date":actual if status in ("運送中","客戶簽收","完成結案") else current["actual_ship_date"],"received_by":signed,"received_date":received if status in ("客戶簽收","完成結案") else current["received_date"],"notes":remarks})
                except Exception:
                    st.error("進度儲存失敗，請重新整理並核對資料後重試。")
                    st.stop()
                if not result:
                    st.error("資料已被其他使用者修改，請重新整理後再操作。")
                else:
                    audit(sid,"更新出貨進度",f"{current['status']} → {status}")
                    try:
                        latest_shipments = records("Shipments")
                        ensure_warranties(latest_shipments)
                    except Exception:
                        st.warning("出貨進度已儲存，但保固紀錄建立失敗。請到保固紀錄頁重新整理補建。")
                        st.stop()
                    st.success("已儲存；已出貨單據已建立保固紀錄。")
                    st.rerun()

with tab4:
    if shipments:
        report = pd.DataFrame(shipments)
        st.bar_chart(report.groupby("status").size().reindex(STAGES,fill_value=0),horizontal=True)
        export = report.drop(columns=["id", "_row"],errors="ignore").copy()
        for c in export.columns:
            export[c] = export[c].map(dstr)
        st.download_button("下載出貨清單 CSV (UTF-8 BOM)", export.to_csv(index=False).encode("utf-8-sig"),"JENNWEI_shipments.csv","text/csv")
        selected_log = st.selectbox("檢視操作紀錄",[s["shipment_no"] for s in shipments],key="log_select")
        sid_log = next(s["id"] for s in shipments if s["shipment_no"] == selected_log)
        history = sorted([{k:r[k] for k in ("created_at","actor","event","detail")} for r in records("ShipmentAudit") if r["shipment_id"] == sid_log], key=lambda r:r["created_at"], reverse=True)
        st.dataframe(pd.DataFrame(history),use_container_width=True,hide_index=True)
    else:
        st.info("尚無出貨資料。")


with tab5:
    st.subheader("🛡️ 出貨後保固紀錄")
    st.caption("已出貨單據自動建檔；保固起算依據及起訖日期須依合約確認。到期日包含當日。每張出貨單以其機台序號建一筆紀錄，多台不同序號請分單建立。")
    if st.button("重新整理／補建已出貨保固紀錄", key="warranty_refresh"):
        st.rerun()
    try:
        ensure_warranties(shipments)
        warranties = records("WarrantyRecords")
    except Exception:
        warranties = None
        st.error("保固紀錄讀寫失敗；出貨資料不受影響。請核對 WarrantyRecords 分頁、標題及編輯權限後重試。")
    if warranties is not None:
        shipment_statuses = {r["id"]: r["status"] for r in shipments}
        for w in warranties:
            w["warranty_status"] = warranty_state(w, shipment_statuses.get(w["shipment_id"], ""))
        keyword = st.text_input("搜尋保固：客戶／機型／序號／出貨單號", key="warranty_search").strip().lower()
        filtered_w = [w for w in warranties if not keyword or any(keyword in str(w.get(k, "")).lower() for k in ("customer_name", "machine_model", "serial_number", "shipment_no"))]
        labels = {"shipment_no":"出貨單號", "customer_name":"客戶", "machine_model":"機型", "serial_number":"機台序號", "actual_ship_date":"出貨日期", "start_basis":"起算依據", "start_date":"保固起日", "end_date":"保固迄日", "warranty_status":"保固狀態", "owner":"負責人"}
        if filtered_w:
            st.dataframe(pd.DataFrame(filtered_w)[list(labels)].rename(columns=labels), use_container_width=True, hide_index=True)
        if warranties:
            st.download_button("下載保固紀錄 CSV", pd.DataFrame(warranties).drop(columns=["_row"]).to_csv(index=False).encode("utf-8-sig"), "JENNWEI_warranties.csv", "text/csv")
            options = {w["id"]: w for w in warranties}
            wid = st.selectbox("選擇保固紀錄", list(options), format_func=lambda x: f"{options[x]['shipment_no']}｜{options[x]['serial_number']}｜{options[x]['customer_name']}")
            w = options[wid]
            with st.form("warranty_" + wid):
                bases = ["待確認", "出貨日", "簽收日", "驗收日", "合約指定日"]
                basis = st.selectbox("保固起算依據", bases, index=bases.index(w["start_basis"]) if w["start_basis"] in bases else 0)
                start_text = st.text_input("保固起日（YYYY-MM-DD；待確認可留空）", value=w["start_date"])
                end_text = st.text_input("保固迄日（YYYY-MM-DD；待確認可留空）", value=w["end_date"])
                terms = st.text_area("保固範圍／合約條件", value=w["terms"])
                warranty_notes = st.text_area("保固備註", value=w["notes"])
                submitted_w = st.form_submit_button("儲存保固條件", type="primary")
            if submitted_w:
                try:
                    start_date = date.fromisoformat(start_text.strip()) if start_text.strip() else None
                    end_date = date.fromisoformat(end_text.strip()) if end_text.strip() else None
                    if bool(start_date) != bool(end_date):
                        raise ValueError("保固起日與迄日需一起填寫，或一起留空。")
                    if start_date and end_date < start_date:
                        raise ValueError("保固迄日不可早於起日。")
                    if start_date and basis == "待確認":
                        raise ValueError("填寫日期時，請確認保固起算依據。")
                    source = next((r for r in shipments if r["id"] == w["shipment_id"]), {})
                    expected_start = source.get("actual_ship_date") if basis == "出貨日" else source.get("received_date") if basis == "簽收日" else None
                    if start_date and basis in ("出貨日", "簽收日") and start_date != expected_start:
                        raise ValueError("起日需與出貨單上的對應日期一致；若尚無日期，請先補齊出貨單。")
                except ValueError as exc:
                    st.error(f"日期或條件有誤：{exc}")
                else:
                    try:
                        saved_w = update_warranty(wid, int(w.get("version") or 1), {"start_basis":basis, "start_date":start_text.strip(), "end_date":end_text.strip(), "terms":terms, "notes":warranty_notes})
                    except Exception:
                        st.error("保固儲存失敗，請重新整理並確認資料後重試。")
                    else:
                        if saved_w:
                            audit(w["shipment_id"], "更新保固條件", f"{basis}：{start_text.strip()} ～ {end_text.strip()}")
                            st.rerun()
                        else:
                            st.error("保固資料已異動，請重新整理後再修改。")
        else:
            st.info("尚無已出貨單據；實際出貨日期及出貨狀態儲存後，會自動建立保固紀錄。")
