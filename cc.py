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
