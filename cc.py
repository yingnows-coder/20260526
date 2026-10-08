import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime, date
import uuid
import io
import base64
from PIL import Image, ImageDraw
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image as PDFImage, KeepTogether, Flowable
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from xml.sax.saxutils import escape

try:
    from streamlit_drawable_canvas import st_canvas
except ImportError:
    st_canvas = None

SIGNATURE_COLUMNS = ['customer_signature', 'signed_by', 'signed_at']
FORM_COLUMNS = ['address', 'manufacture_date', 'arrival_time', 'departure_time', 'is_completed', 'next_service_date', 'next_service_hours', 'service_fee', 'tax_amount', 'service_supervisor', 'sales_supervisor', 'part_1_name', 'part_1_price', 'part_1_qty', 'part_2_name', 'part_2_price', 'part_2_qty', 'part_3_name', 'part_3_price', 'part_3_qty', 'part_4_name', 'part_4_price', 'part_4_qty', 'part_5_name', 'part_5_price', 'part_5_qty', 'part_6_name', 'part_6_price', 'part_6_qty', 'part_7_name', 'part_7_price', 'part_7_qty', 'part_8_name', 'part_8_price', 'part_8_qty', 'part_9_name', 'part_9_price', 'part_9_qty', 'part_10_name', 'part_10_price', 'part_10_qty']

st.set_page_config(page_title='JENN-WEI 售服維修管理系統', page_icon='🛠️', layout='wide')
st.title('🛠️ JENN-WEI 售服維修管理系統')
st.caption('Service Management | 報修、派工、維修、驗收與售服績效')
# aa.py 獨立部署網址
AA_APP_URL = "https://aazzyyii.streamlit.app/"
BB_APP_URL = "https://aazzyybb.streamlit.app/"
nav_aa, nav_bb, nav_space = st.columns([1.5, 1.5, 4])
with nav_aa:
    st.link_button("📋 返回 RFQ 詢價管理", AA_APP_URL or "https://share.streamlit.io/", disabled=not AA_APP_URL, use_container_width=True)
with nav_bb:
    st.link_button("📊 返回 Trello 任務管理", BB_APP_URL, use_container_width=True)

SHEET = 'ServiceTickets'
COLUMNS = ['ticket_id','created_at','customer','contact','phone','machine_model','serial_number','warranty','issue','priority','status','engineer','scheduled_date','diagnosis','repair_action','parts','parts_cost','labor_hours','labor_rate','total_cost','resolution','closed_at','updated_at'] + SIGNATURE_COLUMNS + FORM_COLUMNS
STATUSES = ['新報修','待確認','待派工','待零件','維修中','待驗收','已結案']
PRIORITIES = ['一般','急件','緊急']
# cc.py 是獨立部署的 App，必須在此 App 的 Secrets 指定試算表。
# 建議在 Streamlit Cloud > 售服 App > Settings > Secrets 設定：
# SERVICE_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/你的ID/edit"
def get_spreadsheet_url():
    try:
        direct = st.secrets.get("SERVICE_SPREADSHEET_URL", "")
        if direct:
            return str(direct).strip()
        settings = st.secrets.get("connections", {})
        gsheets = settings.get("gsheets", {})
        return str(gsheets.get("spreadsheet", "") or gsheets.get("spreadsheet_url", "")).strip()
    except Exception:
        return ""

SPREADSHEET_URL = get_spreadsheet_url()
if not SPREADSHEET_URL:
    st.error("尚未設定售服系統的 Google Spreadsheet 網址。")
    st.info('請到售服 App 的 Streamlit Cloud → Settings → Secrets，新增 SERVICE_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/你的試算表ID/edit"。並保留 gsheets 連線的服務帳號憑證。')
    st.stop()

conn = st.connection('gsheets', type=GSheetsConnection)

def read_tickets():
    try:
        data = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=SHEET, ttl=300)
        if data is None:
            data = pd.DataFrame()
        for c in COLUMNS:
            if c not in data.columns:
                data[c] = ''
        # Google Sheets 讀取時 pandas 可能把空白欄推斷成 float64。
        # 後續需要同欄存文字/日期/數字，先轉為 object 避免 LossySetitemError。
        return data[COLUMNS].copy().astype(object)
    except Exception as exc:
        st.error(f'無法讀取 {SHEET} 工作表。請檢查試算表網址、ServiceTickets 分頁及服務帳號的共用權限。')
        st.exception(exc)
        st.stop()

def save_tickets(data):
    try:
        data = data[COLUMNS].copy().astype(object)
        conn.update(spreadsheet=SPREADSHEET_URL, worksheet=SHEET, data=data.fillna(''))
        st.cache_data.clear()
        st.success('資料已儲存')
        st.rerun()
    except Exception as exc:
        st.error('Google Sheets 寫入失敗')
        st.exception(exc)

def number(value):
    try:
        result = pd.to_numeric(value, errors='coerce')
        return 0.0 if pd.isna(result) else float(result)
    except (TypeError, ValueError):
        return 0.0

def safe_str(value):
    return '' if pd.isna(value) else str(value)

def signature_to_base64(canvas_data):
    """縮小並轉為 PNG；避免超過 Google Sheets 單格 50,000 字元上限。"""
    if canvas_data is None:
        return ''
    image = Image.fromarray(canvas_data.astype('uint8'), 'RGBA')
    image.thumbnail((550, 190))
    background = Image.new('RGB', image.size, 'white')
    background.paste(image, mask=image.getchannel('A'))
    background = background.convert('L').point(lambda x: 255 if x > 195 else 0, '1')
    output = io.BytesIO()
    background.save(output, format='PNG', optimize=True)
    encoded = base64.b64encode(output.getvalue()).decode('ascii')
    if len(encoded) > 48000:
        raise ValueError('簽名圖片超過 Google Sheets 單格限制，請清除後簡化簽名。')
    return encoded


def signature_from_drawing(drawing, width=340, height=170):
    """Render Fabric.js strokes with path offsets and transforms, including mobile output."""
    import json
    if isinstance(drawing, str):
        drawing = json.loads(drawing)
    if not isinstance(drawing, dict):
        raise ValueError('簽名資料格式異常，請重新整理後再簽名。')
    render_width, render_height = width * 4, height * 4
    image = Image.new('RGB', (render_width, render_height), 'white')
    pen = ImageDraw.Draw(image)
    count = 0
    for obj in drawing.get('objects', []):
        if not isinstance(obj, dict):
            continue
        commands = obj.get('path', [])
        if isinstance(commands, str):
            try:
                commands = json.loads(commands)
            except ValueError:
                continue
        if not commands:
            continue
        offset = obj.get('pathOffset') or {}
        ox = float(offset.get('x', 0)) if isinstance(offset, dict) else 0
        oy = float(offset.get('y', 0)) if isinstance(offset, dict) else 0
        left = float(obj.get('left') or 0)
        top = float(obj.get('top') or 0)
        sx = float(obj.get('scaleX') or 1)
        sy = float(obj.get('scaleY') or 1)
        width_obj = float(obj.get('width') or 0)
        height_obj = float(obj.get('height') or 0)
        # Fabric's left/top are object bounds; pathOffset is in path coordinates.
        if not offset:
            ox, oy = width_obj / 2, height_obj / 2
        stroke_width = max(2, round(float(obj.get('strokeWidth') or 3)))
        def pos(x, y):
            return (left + (float(x) - ox + width_obj / 2) * sx,
                    top + (float(y) - oy + height_obj / 2) * sy)
        previous = None
        for command in commands:
            if not isinstance(command, (list, tuple)) or not command:
                continue
            kind = str(command[0]).upper()
            args = command[1:]
            if kind == 'M' and len(args) >= 2:
                previous = pos(args[0], args[1])
                pen.ellipse((previous[0]-2,previous[1]-2,previous[0]+2,previous[1]+2),fill='black')
                count += 1
            elif kind == 'L' and len(args) >= 2 and previous is not None:
                point = pos(args[0], args[1]); pen.line([previous,point],fill='black',width=stroke_width)
                previous=point; count += 1
            elif kind in ('Q','C') and previous is not None:
                required = 4 if kind=='Q' else 6
                if len(args) < required: continue
                pts = [pos(args[i], args[i+1]) for i in range(0, required, 2)]
                segment = []
                for i in range(31):
                    t=i/30; u=1-t
                    if kind=='Q':
                        x=u*u*previous[0]+2*u*t*pts[0][0]+t*t*pts[1][0]
                        y=u*u*previous[1]+2*u*t*pts[0][1]+t*t*pts[1][1]
                    else:
                        x=u**3*previous[0]+3*u*u*t*pts[0][0]+3*u*t*t*pts[1][0]+t**3*pts[2][0]
                        y=u**3*previous[1]+3*u*u*t*pts[0][1]+3*u*t*t*pts[1][1]+t**3*pts[2][1]
                    segment.append((x,y))
                pen.line(segment,fill='black',width=stroke_width)
                previous=pts[-1]; count += 1
        # support simple line objects
        if obj.get('type')=='line':
            pen.line([(float(obj.get('x1',0)),float(obj.get('y1',0))),
                      (float(obj.get('x2',0)),float(obj.get('y2',0)))],fill='black',width=stroke_width)
            count += 1
    if count == 0:
        raise ValueError('簽名元件未提供可用筆跡。可改用下方上傳簽名照片功能。')
    # Normalize actual ink bounds into the full signature pad, avoiding
    # device-pixel-ratio coordinate mismatches and the top-left-quarter crop.
    ink = Image.eval(image.convert('L'), lambda v: 255-v)
    bounds = ink.getbbox()
    if bounds is None:
        raise ValueError('簽名沒有可辨識的筆跡，請清除後重新簽名。')
    cropped = image.crop((max(0,bounds[0]-5),max(0,bounds[1]-5),
                          min(render_width,bounds[2]+5),min(render_height,bounds[3]+5)))
    cropped.thumbnail((width-24,height-24), Image.Resampling.LANCZOS)
    normalized = Image.new('RGB',(width,height),'white')
    normalized.paste(cropped,((width-cropped.width)//2,(height-cropped.height)//2))
    output=io.BytesIO(); normalized.save(output,format='PNG',optimize=True)
    encoded=base64.b64encode(output.getvalue()).decode('ascii')
    if len(encoded)>48000:
        raise ValueError('簽名資料超過 Google Sheets 限制。')
    return encoded


def signature_from_upload(uploaded):
    """Reliable fallback for phones: image upload or camera snapshot."""
    image = Image.open(uploaded).convert('RGB')
    image.thumbnail((550,190))
    white=Image.new('RGB',(550,190),'white')
    white.paste(image,((550-image.width)//2,(190-image.height)//2))
    output=io.BytesIO(); white.save(output,format='PNG',optimize=True)
    encoded=base64.b64encode(output.getvalue()).decode('ascii')
    if len(encoded)>48000:
        # threshold and reduce palette for ink signatures
        mono=white.convert('L').point(lambda x: 255 if x>190 else 0,'1')
        output=io.BytesIO(); mono.save(output,format='PNG',optimize=True)
        encoded=base64.b64encode(output.getvalue()).decode('ascii')
    if len(encoded)>48000:
        raise ValueError('簽名圖片過大，請使用簡單白底黑字的簽名圖片。')
    return encoded


def make_ticket_pdf(ticket):
    """JENN WEI 售後服務單：參照紙本表格、固定 A4 一頁、簽名框內置中。"""
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.platypus import Paragraph
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib import colors
    from xml.sax.saxutils import escape as xml_escape

    font = 'MSung-Light'
    pdfmetrics.registerFont(UnicodeCIDFont(font))
    out = io.BytesIO()
    c = canvas.Canvas(out, pagesize=A4)
    c.setTitle('JENN WEI MACHINERY - 售後服務單')
    pw, ph = A4
    x0, x1 = 15*mm, pw-15*mm
    W = x1-x0
    ink = colors.HexColor('#242424')
    red = colors.HexColor('#9B352E')
    c.setStrokeColor(ink)
    c.setLineWidth(.65)

    def txt(x,y,value,size=9,color=ink):
        c.setFillColor(color)
        c.setFont(font,size)
        c.drawString(x,y,str(value))

    def line(xa,ya,xb,yb):
        c.line(xa,ya,xb,yb)

    def box(x,y,w,h):
        c.rect(x,y,w,h,fill=0,stroke=1)

    def value(key):
        return safe_str(ticket.get(key,''))

    def cell(x,y,w,h,label,content='',label_width=None,fontsize=9):
        box(x,y,w,h)
        txt(x+4,y+h-13,label,9,red)
        if content:
            label_width = label_width if label_width is not None else min(w*.36, 31*mm)
            para = Paragraph(xml_escape(str(content)).replace('\n','<br/>'),
                ParagraphStyle('v',fontName=font,fontSize=fontsize,leading=fontsize+2,wordWrap='CJK'))
            avail_w = max(12,w-label_width-7)
            _,h2=para.wrap(avail_w, 1000)
            # 文字過多時限制在欄位內；全文仍保存在 Sheets
            if h2 <= h-4:
                para.drawOn(c,x+label_width,y+h-3-h2)
            else:
                txt(x+label_width,y+h-13,str(content)[:18]+'…',7)

    def multiline_box(x,y,w,h,label,content):
        box(x,y,w,h)
        txt(x+5,y+h-14,label,9,red)
        raw = str(content or '')
        for sz in [9,8,7,6,5]:
            style=ParagraphStyle('m',fontName=font,fontSize=sz,leading=sz+2,wordWrap='CJK')
            para=Paragraph(xml_escape(raw).replace('\n','<br/>') or ' ',style)
            _,ht=para.wrap(w-13,10000)
            if ht<=h-25:
                para.drawOn(c,x+6,y+h-22-ht)
                return
        # 极端長文字以縮放繪製，不會新增第二頁
        c.saveState()
        scale=min(1.,(h-25)/max(ht,1))
        c.translate(x+6,y+4)
        c.scale(scale,scale)
        para.wrap((w-13)/scale,10000)
        para.drawOn(c,0,0)
        c.restoreState()

    # 公司抬頭，依使用者提供的售後服務單格式
    logo_path = __import__('os').path.join(__import__('os').path.dirname(__file__),'jenn_wei_logo.png')
    if __import__('os').path.exists(logo_path):
        try:
            c.drawImage(logo_path,x0,ph-48*mm,width=36*mm,height=32*mm,preserveAspectRatio=True,anchor='c',mask='auto')
        except Exception:
            pass
    txt(x0+41*mm,ph-20*mm,'震唯機械股份有限公司',15)
    c.setFont('Helvetica-Bold',10)
    c.drawString(x0+41*mm,ph-25*mm,'JENN WEI MACHINERY CO., LTD.')
    c.setFont('Helvetica',7.5)
    c.drawString(x0+41*mm,ph-30*mm,'E-mail: L3352368@ms49.hinet.net')
    c.drawString(x0+41*mm,ph-34*mm,'http://www.jennwei.com.tw/')
    c.drawString(x0+41*mm,ph-38*mm,'TEL: 886-4-23352368   FAX: 886-4-23353880')
    c.setFillColor(red)
    c.setFont(font,14)
    c.drawCentredString(pw/2,ph-53*mm,'售後服務單')
    c.setFillColor(ink)
    c.setFont('Helvetica',8)
    c.drawRightString(x1,ph-53*mm,'No. '+value('ticket_id'))

    # 表格自上而下。A4 固定 1 頁
    top = ph-59*mm
    row=11*mm
    y=top-row
    cell(x0,y,W*.46,row,'客戶名稱：',value('customer'))
    cell(x0+W*.46,y,W*.31,row,'電話：',value('phone'))
    cell(x0+W*.77,y,W*.23,row,'日期：',value('created_at')[:10],18*mm,8)
    y-=row
    cell(x0,y,W,row,'住址：',value('address'))
    y-=row
    cell(x0,y,W*.35,row,'機型：',value('machine_model'))
    cell(x0+W*.35,y,W*.35,row,'機號：',value('serial_number'))
    cell(x0+W*.70,y,W*.30,row,'出廠日：',value('manufacture_date'))

    fault_h=29*mm
    y-=fault_h
    multiline_box(x0,y,W,fault_h,'故障原因：',value('diagnosis') or value('issue'))
    action_h=30*mm
    y-=action_h
    multiline_box(x0,y,W,action_h,'處理方式：',value('repair_action') or value('resolution'))

    y-=10*mm
    cell(x0,y,W*.30,10*mm,'到達時間：',value('arrival_time'))
    cell(x0+W*.30,y,W*.30,10*mm,'離開時間：',value('departure_time'))
    cell(x0+W*.60,y,W*.40,10*mm,'本案是否完成：', '是' if (value('is_completed') == '是' or (not value('is_completed') and value('status')=='已結案')) else '待確認',33*mm)
    y-=10*mm
    cell(x0,y,W*.76,10*mm,'未處理完成、再處理日期：',value('next_service_date') or value('scheduled_date'))
    cell(x0+W*.76,y,W*.24,10*mm,'時 間：',value('next_service_hours'),17*mm)
    # 零件費用明細，與原始紙本的 10 列一致
    parts_h=6.4*mm
    y-=parts_h
    fractions=[.50,.14,.14,.22]
    heads=['零件名稱編號','單價','數量','小計']
    cursor=x0
    for frac,head in zip(fractions,heads):
        box(cursor,y,W*frac,parts_h)
        c.setFont(font,8)
        c.drawCentredString(cursor+W*frac/2,y+2*mm,head)
        cursor+=W*frac
    legacy_parts=[p.strip() for p in value('parts').replace('；','\\n').replace(';','\\n').splitlines() if p.strip()]
    parts_total=0.0
    active_parts=[]
    for i in range(1,11):
        name=value(f'part_{i}_name') or (legacy_parts[i-1] if i<=len(legacy_parts) else '')
        if name.strip():
            active_parts.append((name,number(ticket.get(f'part_{i}_price','')),number(ticket.get(f'part_{i}_qty',''))))
    for display_n,(name,price,qty) in enumerate(active_parts,1):
        y-=parts_h
        cursor=x0
        subtotal=price*qty
        parts_total+=subtotal
        values=[f'{display_n}. {name}', f'{price:,.2f}' if price else '',
                f'{qty:g}' if qty else '',
                f'{subtotal:,.2f}' if price and qty else '']
        for frac,entry in zip(fractions,values):
            box(cursor,y,W*frac,parts_h)
            if entry:
                txt(cursor+3,y+2*mm,entry[:46],7)
            cursor+=W*frac

    y-=10*mm
    service_fee=number(ticket.get('service_fee',''))
    tax=number(ticket.get('tax_amount',''))
    total=parts_total+service_fee+tax
    cell(x0,y,W*.34,10*mm,'服務費：',f'{service_fee:,.2f}' if service_fee else '0')
    cell(x0+W*.34,y,W*.30,10*mm,'5% 稅金：',f'{tax:,.2f}' if tax else '0')
    cell(x0+W*.64,y,W*.36,10*mm,'費用總計：',f'{total:,.2f}')

    # 保留簽名框在同一頁；姓名/時間與筆跡不重疊
    y-=11*mm
    cell(x0,y,W*.33,11*mm,'服務人員：',value('engineer'))
    cell(x0+W*.33,y,W*.33,11*mm,'售服主管：',value('service_supervisor'))
    cell(x0+W*.66,y,W*.34,11*mm,'業務主管：',value('sales_supervisor'))
    sign_h=24*mm
    y-=sign_h
    box(x0,y,W,sign_h)
    txt(x0+4,y+sign_h-11,'客戶確認／簽名：',9,red)
    signature=value('customer_signature')
    if signature:
        try:
            source=Image.open(io.BytesIO(base64.b64decode(signature))).convert('RGB')
            gray=source.convert('L')
            bbox=gray.point(lambda px:255 if px<205 else 0).getbbox()
            if bbox:
                source=source.crop(bbox)
                # 簽名採固定上限，禁止因原始筆跡較小而被放大到整個簽名框。
                # 預留左側欄位標題空間，簽名在右側區域等比例置中。
                max_w = 72*mm
                max_h = 15*mm
                natural_scale = 72.0 / 110.0  # 原始畫布約 110 DPI，避免過度放大
                factor = min(max_w/source.width, max_h/source.height, natural_scale)
                iw, ih = source.width*factor, source.height*factor
                region_left = x0 + 42*mm
                region_width = W - 47*mm
                draw_x = region_left + (region_width-iw)/2
                draw_y = y + (sign_h-ih)/2 - 1*mm
                c.drawImage(ImageReader(source), draw_x, draw_y, width=iw, height=ih)
        except Exception:
            txt(x0+45*mm,y+10*mm,'簽名圖片無法載入',8)
    txt(x0+3,y-10,'簽收人：'+value('signed_by')+'   簽收時間：'+value('signed_at'),8)
    txt(x0+3,y-23,'本單據記錄維修及簽收資訊；簽收不代表另行承諾保固或費用條件。',7)
    c.showPage()
    c.save()
    return out.getvalue()


df = read_tickets()

new_tab, manage_tab, form_tab, dashboard_tab, history_tab = st.tabs(['📝 建立報修單','🔧 維修工單管理','📄 售後服務單填寫','📊 售服 KPI','📚 維修紀錄'])
with new_tab:
    st.subheader('建立客戶報修工單')
    with st.form('new_service_ticket', clear_on_submit=True):
        a,b,c = st.columns(3)
        with a:
            customer = st.text_input('客戶名稱 *')
            contact = st.text_input('聯絡人')
            phone = st.text_input('電話 / 聯絡方式')
        with b:
            model = st.text_input('機台型號 *')
            serial = st.text_input('機台序號')
            warranty = st.selectbox('保固狀態', ['待確認','保固內','保固外'])
        with c:
            priority = st.selectbox('緊急程度', PRIORITIES)
            engineer = st.text_input('預定負責工程師')
            schedule = st.date_input('預計處理日期', value=date.today())
        st.markdown('#### 售後服務單基本資料')
        address = st.text_input('客戶住址')
        manufacture_date = st.text_input('機台出廠日（YYYY-MM-DD，可留空）')
        issue = st.text_area('故障現象 / 客戶反映 *')
        submitted = st.form_submit_button('➕ 建立維修工單', type='primary')
    if submitted:
        if not customer.strip() or not model.strip() or not issue.strip():
            st.error('請填寫客戶名稱、機台型號及故障現象。')
        else:
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            ticket = {c:'' for c in COLUMNS}
            ticket.update(ticket_id=f'SRV-{datetime.now():%Y%m%d}-{uuid.uuid4().hex[:4].upper()}',created_at=now,customer=customer.strip(),contact=contact.strip(),phone=phone.strip(),machine_model=model.strip(),serial_number=serial.strip(),warranty=warranty,issue=issue.strip(),priority=priority,status='新報修',engineer=engineer.strip(),scheduled_date=str(schedule),parts_cost=0.0,labor_hours=0.0,labor_rate=0.0,total_cost=0.0,updated_at=now,address=address.strip(),manufacture_date=manufacture_date.strip())
            save_tickets(pd.concat([df,pd.DataFrame([ticket])],ignore_index=True))

with manage_tab:
    st.subheader('維修工單管理')
    f1,f2,f3 = st.columns(3)
    with f1: status_filter = st.selectbox('維修狀態', ['全部']+STATUSES)
    with f2: engineer_filter = st.selectbox('負責工程師', ['全部']+sorted(x for x in df['engineer'].dropna().astype(str).unique() if x.strip()))
    with f3: keyword = st.text_input('搜尋工單 / 客戶 / 型號 / 序號')
    view = df.copy()
    if status_filter != '全部': view = view[view['status']==status_filter]
    if engineer_filter != '全部': view = view[view['engineer']==engineer_filter]
    if keyword: view = view[view.astype(str).apply(lambda s:s.str.contains(keyword,case=False,regex=False,na=False)).any(axis=1)]
    st.caption(f'符合條件工單：{len(view)} 筆')
    if view.empty:
        st.info('目前沒有符合條件的維修工單。')
    else:
        st.dataframe(view[['ticket_id','customer','machine_model','serial_number','priority','status','engineer','scheduled_date','updated_at']],hide_index=True,use_container_width=True)
        ticket_id = st.selectbox('選擇要更新的工單',view['ticket_id'].astype(str).tolist())
        matches = df.index[df['ticket_id'].astype(str)==ticket_id]
        if len(matches):
            idx=matches[0]; item=df.loc[idx]
            with st.form('edit_service_ticket'):
                st.markdown(f'**{ticket_id}｜{safe_str(item["customer"])}｜{safe_str(item["machine_model"])}**')
                c1,c2,c3=st.columns(3)
                with c1:
                    current=safe_str(item['status']); status=st.selectbox('維修狀態',STATUSES,index=STATUSES.index(current) if current in STATUSES else 0)
                    priority=safe_str(item['priority']); priority_new=st.selectbox('優先程度',PRIORITIES,index=PRIORITIES.index(priority) if priority in PRIORITIES else 0)
                with c2:
                    engineer_new=st.text_input('負責工程師',value=safe_str(item['engineer']))
                    warranty=safe_str(item['warranty']); warranty_new=st.selectbox('保固', ['待確認','保固內','保固外'],index=['待確認','保固內','保固外'].index(warranty) if warranty in ['待確認','保固內','保固外'] else 0)
                with c3:
                    parts_cost=st.number_input('零件費用',min_value=0.0,value=number(item['parts_cost']))
                    labor_hours=st.number_input('維修工時（小時）',min_value=0.0,value=number(item['labor_hours']))
                    labor_rate=st.number_input('每小時工資',min_value=0.0,value=number(item['labor_rate']))
                diagnosis=st.text_area('故障診斷',value=safe_str(item['diagnosis']))
                action=st.text_area('維修處置',value=safe_str(item['repair_action']))
                parts=st.text_area('更換零件',value=safe_str(item['parts']))
                resolution=st.text_area('測試 / 客戶驗收紀錄',value=safe_str(item['resolution']))
                st.metric('本次維修費用合計',f'{parts_cost+labor_hours*labor_rate:,.0f}')
                saved=st.form_submit_button('💾 儲存維修紀錄',type='primary')
            if saved:
                updates=dict(status=status,priority=priority_new,engineer=engineer_new.strip(),warranty=warranty_new,parts_cost=parts_cost,labor_hours=labor_hours,labor_rate=labor_rate,total_cost=parts_cost+labor_hours*labor_rate,diagnosis=diagnosis,repair_action=action,parts=parts,resolution=resolution,updated_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
                if status=='已結案' and not safe_str(item['closed_at']): updates['closed_at']=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                if status!='已結案': updates['closed_at']=''
                for k,v in updates.items(): df.at[idx,k]=v
                save_tickets(df)

with form_tab:
    st.subheader('📄 售後服務單 — 對應紙本欄位填寫')
    st.caption('依紙本售後服務單順序填寫。儲存後可至「客戶簽名／PDF」下載 A4 單頁表單。')
    if df.empty:
        st.info('請先在「建立報修單」建立工單。')
    else:
        form_ticket_id = st.selectbox('選擇維修單', df['ticket_id'].astype(str).tolist(), key='form_ticket_id')
        form_idx = df.index[df['ticket_id'].astype(str)==form_ticket_id][0]
        item = df.loc[form_idx]
        def old(k):
            return safe_str(item.get(k,''))
        # 零件列數由 session_state 管理；新增按鈕放在表單外，才能立即顯示下一列。
        count_key = f'parts_visible_count_{form_ticket_id}'
        existing_slots = [n for n in range(1, 11) if old(f'part_{n}_name').strip()]
        if count_key not in st.session_state:
            st.session_state[count_key] = max(1, max(existing_slots, default=0))
        st.caption(f'零件明細：目前顯示 {st.session_state[count_key]} 筆，最多 10 筆。')
        if st.button('➕ 追加一筆零件', key=f'add_part_{form_ticket_id}',
                     disabled=st.session_state[count_key] >= 10):
            st.session_state[count_key] += 1
            st.rerun()
        with st.form('paper_service_form'):
            st.markdown('#### 1. 客戶及機台資料')
            a,b,c = st.columns(3)
            with a:
                customer_f=st.text_input('客戶名稱',value=old('customer'))
                address_f=st.text_input('客戶住址',value=old('address'))
                machine_f=st.text_input('機型',value=old('machine_model'))
            with b:
                phone_f=st.text_input('電話',value=old('phone'))
                serial_f=st.text_input('機號',value=old('serial_number'))
                manufactured_f=st.text_input('出廠日',value=old('manufacture_date'),placeholder='YYYY-MM-DD')
            with c:
                created_f=st.text_input('服務單日期',value=old('created_at')[:10],placeholder='YYYY-MM-DD')
                engineer_f=st.text_input('服務人員',value=old('engineer'))
            st.markdown('#### 2. 故障原因與處理方式')
            diagnosis_f=st.text_area('故障原因',value=old('diagnosis') or old('issue'),height=110)
            repair_f=st.text_area('處理方式',value=old('repair_action'),height=110)
            st.markdown('#### 3. 服務時間與完工狀態')
            a,b,c=st.columns(3)
            with a:
                arrival_f=st.text_input('到達時間',value=old('arrival_time'),placeholder='YYYY-MM-DD HH:MM')
            with b:
                departure_f=st.text_input('離開時間',value=old('departure_time'),placeholder='YYYY-MM-DD HH:MM')
            with c:
                completed_default=old('is_completed') or ('是' if old('status')=='已結案' else '否')
                completed_f=st.radio('本案是否完成',options=['是','否'],index=0 if completed_default=='是' else 1,horizontal=True)
            a,b=st.columns(2)
            with a:
                next_date_f=st.text_input('未完成，再處理日期',value=old('next_service_date'),placeholder='YYYY-MM-DD')
            with b:
                next_hours_f=st.text_input('再處理時間',value=old('next_service_hours'),placeholder='HH:MM')
            st.markdown('#### 4. 零件名稱、單價、數量（按需追加）')
            parts_input=[]
            for n in range(1, st.session_state[count_key] + 1):
                a,b,c=st.columns([5,2,2])
                with a:
                    name_f=st.text_input(f'{n}. 零件名稱／編號',value=old(f'part_{n}_name'),key=f'partname_{form_ticket_id}_{n}')
                with b:
                    price_f=st.number_input(f'{n}. 單價',min_value=0.0,value=number(item.get(f'part_{n}_price','')),step=1.0,key=f'partprice_{form_ticket_id}_{n}')
                with c:
                    qty_f=st.number_input(f'{n}. 數量',min_value=0.0,value=number(item.get(f'part_{n}_qty','')),step=1.0,key=f'partqty_{form_ticket_id}_{n}')
                parts_input.append((name_f,price_f,qty_f))
            st.markdown('#### 5. 費用與主管確認')
            a,b,c=st.columns(3)
            with a:
                service_fee_f=st.number_input('服務費',min_value=0.0,value=number(item.get('service_fee','')),step=100.0)
            with b:
                tax_default=number(item.get('tax_amount',''))
                tax_f=st.number_input('5% 稅金（手動填寫）',min_value=0.0,value=tax_default,step=1.0)
            with c:
                subtotal_f=sum(price*qty for name,price,qty in parts_input if name.strip())
                st.metric('費用總計（零件＋服務費＋稅金）',f'{subtotal_f+service_fee_f+tax_f:,.2f}')
            a,b=st.columns(2)
            with a:
                supervisor_f=st.text_input('售服主管',value=old('service_supervisor'))
            with b:
                sales_f=st.text_input('業務主管',value=old('sales_supervisor'))
            saved_form=st.form_submit_button('💾 儲存售後服務單資料',type='primary',use_container_width=True)
        if saved_form:
            updates={
                'customer':customer_f.strip(),'address':address_f.strip(),'phone':phone_f.strip(),
                'machine_model':machine_f.strip(),'serial_number':serial_f.strip(),
                'manufacture_date':manufactured_f.strip(),'engineer':engineer_f.strip(),
                'diagnosis':diagnosis_f.strip(),'repair_action':repair_f.strip(),
                'arrival_time':arrival_f.strip(),'departure_time':departure_f.strip(),
                'is_completed':completed_f,'next_service_date':next_date_f.strip(),
                'next_service_hours':next_hours_f.strip(),'service_fee':service_fee_f,
                'tax_amount':tax_f,'service_supervisor':supervisor_f.strip(),
                'sales_supervisor':sales_f.strip(),'parts_cost':subtotal_f,
                'total_cost':subtotal_f+service_fee_f+tax_f,
                'updated_at':datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            if created_f.strip():
                original_time=old('created_at')
                updates['created_at']=created_f.strip() + (' '+original_time.split(' ',1)[1] if ' ' in original_time else '')
            for n,(name,price,qty) in enumerate(parts_input,1):
                updates[f'part_{n}_name']=name.strip()
                updates[f'part_{n}_price']=price
                updates[f'part_{n}_qty']=qty
            # 同時處理舊資料的數值欄與新增文字欄，避免 pandas 3 嚴格型別寫入失敗。
            for key in updates:
                if key not in df.columns:
                    df[key] = pd.Series('', index=df.index, dtype=object)
                elif df[key].dtype != object:
                    df[key] = df[key].astype(object)
            for key,val in updates.items():
                df.at[form_idx,key]=val
            save_tickets(df)

    st.divider()
    st.markdown('### ✍️ 客戶簽名／PDF 匯出')
    st.subheader('✍️ 客戶手機簽名與維修單 PDF')
    st.caption('手機可用手指直接在簽名區書寫。簽名會儲存至同一份 Google Sheets 的 ServiceTickets 分頁。')
    if df.empty:
        st.info('請先建立維修工單。')
    else:
        selected_id = st.selectbox('選擇維修單號', df['ticket_id'].astype(str).tolist(), index=df['ticket_id'].astype(str).tolist().index(form_ticket_id), key='sign_ticket')
        selected_idx = df.index[df['ticket_id'].astype(str) == selected_id][0]
        ticket = df.loc[selected_idx]
        st.write(f'**客戶：** {safe_str(ticket["customer"])}　**機台：** {safe_str(ticket["machine_model"])}')
        signed = bool(safe_str(ticket.get('customer_signature', '')))
        if signed:
            st.success(f'已簽收：{safe_str(ticket.get("signed_by", ""))}　{safe_str(ticket.get("signed_at", ""))}')
            try:
                st.image(base64.b64decode(safe_str(ticket['customer_signature'])), width=340)
            except Exception:
                st.warning('已存簽名資料，但無法顯示預覽。')
        else:
            st.info('本工單尚未簽名。')
        signer = st.text_input('客戶簽收人姓名', value=safe_str(ticket.get('signed_by', '')), key=f'signer_{selected_id}')
        if st_canvas is None:
            st.error('尚未安裝手機簽名元件。請在 requirements.txt 加入 streamlit-drawable-canvas。')
        else:
            st.caption('請在白色區域完整簽名。系統將以手寫軌跡重建完整簽名，不再使用可能只擷取左上角的圖片。')
            # Fixed canvas pixels: keep within a typical mobile viewport.
            SIGN_PAD_WIDTH, SIGN_PAD_HEIGHT = 340, 170
            st.caption('簽名區：340 × 170 px（手機友善尺寸）')
            canvas = st_canvas(
                fill_color='rgba(255,255,255,0)',
                stroke_width=3,
                stroke_color='#17263C',
                background_color='#FFFFFF',
                update_streamlit=True,
                height=SIGN_PAD_HEIGHT,
                width=SIGN_PAD_WIDTH,
                drawing_mode='freedraw',
                key=f'canvas_mobile_v2_{selected_id}',
            )
            st.caption('如畫布簽名無法儲存，可在手機記事本／繪圖 App 簽名後，截圖並上傳。')
            uploaded_signature = st.file_uploader(
                '備用：上傳簽名圖片（PNG／JPG）', type=['png','jpg','jpeg'],
                key=f'upload_signature_{selected_id}'
            )
            if st.button('💾 儲存客戶簽名', type='primary', key=f'save_signature_{selected_id}'):
                if not signer.strip():
                    st.error('請先輸入簽收人姓名。')
                else:
                    try:
                        if uploaded_signature is not None:
                            encoded = signature_from_upload(uploaded_signature)
                        else:
                            drawing = canvas.json_data or {}
                            # IMPORTANT: canvas.image_data may be only the top-left quarter
                            # on high-DPI mobile displays (devicePixelRatio=2). Rebuild
                            # from Fabric vector strokes, not from its clipped raster.
                            if drawing.get('objects'):
                                encoded = signature_from_drawing(
                                    drawing, width=SIGN_PAD_WIDTH, height=SIGN_PAD_HEIGHT
                                )
                            else:
                                st.error('尚未偵測到簽名筆跡，請重新簽名。')
                                st.stop()
                        df.at[selected_idx, 'customer_signature'] = encoded
                        df.at[selected_idx, 'signed_by'] = signer.strip()
                        df.at[selected_idx, 'signed_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        df.at[selected_idx, 'updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        save_tickets(df)
                    except Exception as exc:
                        st.error(f'簽名儲存失敗：{exc}')
        st.divider()
        st.download_button(
            '📄 下載維修服務單 PDF（含已儲存簽名）',
            data=make_ticket_pdf(ticket),
            file_name=f'{selected_id}_service_report.pdf',
            mime='application/pdf',
            use_container_width=True,
        )
        st.caption('先儲存簽名，頁面重新載入後再下載 PDF，才會包含最新簽名。')


with dashboard_tab:
    st.subheader('售服管理 KPI')
    total=len(df); closed=int((df['status']=='已結案').sum()); active=total-closed; urgent=int(df['priority'].isin(['急件','緊急']).sum())
    a,b,c,d=st.columns(4)
    a.metric('累計報修',total); b.metric('處理中',active); c.metric('已結案',closed); d.metric('急件 / 緊急',urgent)
    st.markdown('#### 維修狀態分布')
    st.bar_chart(df['status'].value_counts().reindex(STATUSES,fill_value=0))
    st.markdown('#### 工程師工單數')
    if not df.empty:
        st.bar_chart(df['engineer'].replace('', '未指派').fillna('未指派').value_counts())
    st.metric('累計維修費用',f'{pd.to_numeric(df["total_cost"],errors="coerce").fillna(0).sum():,.0f}')

with history_tab:
    st.subheader('維修紀錄查詢 / 匯出')
    st.dataframe(df.sort_values('created_at',ascending=False),hide_index=True,use_container_width=True)
    st.download_button('⬇️ 匯出維修工單 CSV',data=df.to_csv(index=False).encode('utf-8-sig'),file_name='service_tickets.csv',mime='text/csv')

with st.expander('⚙️ Google Sheets 工作表設定'):
    st.write(f'請在與 aa.py 相同的 Google Spreadsheet 中新增工作表 **{SHEET}**，第一列依序建立以下欄位：')
    st.code(','.join(COLUMNS))
    st.info('新增的 3 欄為 customer_signature、signed_by、signed_at，請加在 ServiceTickets 工作表第一列原有欄位後方。')
    st.caption('讀取快取 TTL 為 300 秒；寫入後清除 Streamlit 資料快取。')
