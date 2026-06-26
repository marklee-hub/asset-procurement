"""
行政後勤 ── 採購・資產整合平台 (Streamlit + Google Sheets)
此版本：外觀比照 React 原型（墨綠/琥珀配色、深色側欄、卡片式介面），
後端邏輯與資料結構不變。
"""

import os
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import date, datetime

# ============================ 設定 ============================
st.set_page_config(page_title="採購・資產整合平台", page_icon="📦", layout="wide")

# ---- 配色（對齊 React 版 tokens）----
INK, SUB, FAINT = "#161A22", "#5A6472", "#8A93A2"
LINE, LINE_SOFT = "#E4E8EE", "#EEF1F5"
JADE, JADE_SOFT = "#0F766E", "#E2F1EF"
AMBER, AMBER_SOFT = "#B45309", "#FBEEDC"
INDIGO, INDIGO_SOFT = "#4338CA", "#E7E7FA"
DANGER = "#B42318"

CATEGORIES = ["3C設備", "辦公家具", "音響設備", "文宣品", "文具耗材", "其他"]
DEFAULT_UNITS = ["傳道部", "行銷部", "影音部", "行政部", "神學院", "財務部"]
UNITS = list(DEFAULT_UNITS)   # 啟動時會從 Google Sheet 的 units 分頁覆蓋
LEDGER_THRESHOLD = 80000      # ≥ 此金額預設為「列帳資產」

# 資產分類（採購／手動可選三種；其中「一般耗材」不列管、不給編號、不進資產清單）
ASSET_CLASS_OPTS = ["列管資產", "列帳資產", "一般耗材"]
ASSET_RECORD_CLASSES = ["列管資產", "列帳資產"]          # 會給財產編號並進入資產清單
CLASS_PREFIX = {"列管資產": "A26", "列帳資產": "B26"}     # 列管 A26-xxxx／列帳 B26-xxxx

STATUS_OPTS = ["草稿", "待驗收", "已驗收"]
ASSET_STATUS_OPTS = ["使用中", "維修中", "已報廢"]

STATUS_STYLE = {"草稿": (LINE_SOFT, SUB), "待驗收": (AMBER_SOFT, AMBER), "已驗收": (JADE_SOFT, JADE), "已作廢": ("#F3EDED", DANGER)}
ASTATUS_STYLE = {"使用中": (JADE_SOFT, JADE), "維修中": (AMBER_SOFT, AMBER), "已報廢": ("#F3EDED", DANGER), "已作廢": ("#F3EDED", DANGER)}
ATYPE_STYLE = {"列帳資產": (INDIGO_SOFT, INDIGO), "列管資產": (JADE_SOFT, JADE), "一般耗材": (LINE_SOFT, SUB)}

SCHEMAS = {
    "suppliers":       ["id", "name", "tax_id", "contact", "phone", "note"],
    "purchase_orders": ["id", "supplier_id", "date", "status", "purpose", "buyer", "note"],
    "po_items":        ["po_id", "name", "category", "qty", "price"],
    "assets":          ["id", "name", "category", "value", "source_po",
                        "acquired", "asset_type", "unit", "status"],
    "units":           ["name"],
    "void_requests":   ["req_id", "target_type", "target_id", "reason",
                        "requested_by", "requested_at", "status"],
}
NUMERIC = {"po_items": ["qty", "price"], "assets": ["value"]}


# ============================ 樣式 ============================
def inject_css():
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"], .stApp {{
        font-family: 'Manrope','PingFang TC','Noto Sans TC','Microsoft JhengHei',sans-serif;
    }}
    .stApp {{ background: #F6F8F9; }}
    #MainMenu, footer {{ visibility: hidden; }}
    [data-testid="stHeader"] {{ background: transparent; }}
    .block-container {{ padding-top: 2.2rem; padding-bottom: 3rem; max-width: 1140px; }}

    /* 側欄深色漸層 */
    section[data-testid="stSidebar"] {{ background: linear-gradient(180deg,#16221F 0%,#0F1B19 100%); }}
    section[data-testid="stSidebar"] * {{ color: #C9D6D2; }}
    section[data-testid="stSidebar"] .brand-title {{ color:#fff; font-weight:800; font-size:1rem; }}

    /* 主按鈕：圓角＋陰影 */
    .stButton > button {{ border-radius:12px; font-weight:600; border:1px solid {LINE}; transition:all .15s ease; }}
    .stButton > button:hover {{ transform:translateY(-1px); }}
    .stButton > button[kind="primary"] {{ background:{JADE}; border-color:{JADE}; box-shadow:0 6px 16px rgba(15,118,110,.25); }}
    .stButton > button[kind="primary"]:hover {{ background:#0c655e; border-color:#0c655e; box-shadow:0 8px 20px rgba(15,118,110,.32); }}

    /* 側欄導覽：乾淨無框、左對齊、選中highlight＋左側強調條 */
    section[data-testid="stSidebar"] .stButton > button {{
        border:none !important; background:transparent !important; color:#C9D6D2 !important;
        justify-content:flex-start !important; text-align:left !important;
        font-weight:600; border-radius:11px; padding:10px 14px; box-shadow:none !important; transform:none !important;
    }}
    section[data-testid="stSidebar"] .stButton > button p {{ text-align:left; width:100%; }}
    section[data-testid="stSidebar"] .stButton > button:hover {{
        background:rgba(255,255,255,.07) !important; color:#fff !important;
    }}
    section[data-testid="stSidebar"] .stButton > button[kind="primary"] {{
        background:rgba(15,118,110,.22) !important; color:#fff !important;
        box-shadow:inset 3px 0 0 {JADE} !important;
    }}

    /* 輸入元件圓角 */
    [data-baseweb="select"] > div, .stTextInput input, .stNumberInput input, .stTextArea textarea {{
        border-radius:12px !important;
    }}
    .stTextInput input:focus, .stNumberInput input:focus, .stTextArea textarea:focus {{
        border-color:{JADE} !important; box-shadow:0 0 0 2px rgba(15,118,110,.12) !important;
    }}

    /* 標題 */
    h1 {{ font-weight:800 !important; letter-spacing:-.015em; color:{INK}; }}
    h2, h3 {{ font-weight:700 !important; color:{INK}; }}

    /* 卡片系統：柔和陰影＋hover 微浮 */
    .card {{ background:#fff; border:1px solid {LINE}; border-radius:18px; padding:18px;
             box-shadow:0 1px 3px rgba(16,24,40,.04), 0 1px 2px rgba(16,24,40,.03); }}
    .grid4 {{ display:grid; grid-template-columns:repeat(4,1fr); gap:14px; }}
    .grid2 {{ display:grid; grid-template-columns:repeat(2,1fr); gap:14px; }}
    @media (max-width:760px){{ .grid4{{grid-template-columns:repeat(2,1fr)}} .grid2{{grid-template-columns:1fr}} }}
    .stat-label {{ color:{FAINT}; font-size:12px; font-weight:700; letter-spacing:.04em; text-transform:uppercase; }}
    .stat-value {{ font-size:27px; font-weight:800; line-height:1.1; margin-top:5px; font-variant-numeric:tabular-nums; }}
    .stat-sub {{ color:{SUB}; font-size:12px; margin-top:3px; }}
    .pill {{ display:inline-flex; align-items:center; padding:2px 10px; border-radius:999px; font-size:12px; font-weight:700; white-space:nowrap; }}
    .flow {{ display:flex; align-items:center; gap:10px; flex-wrap:wrap; }}
    .flowcard {{ border-radius:14px; padding:13px 16px; flex:1; min-width:150px; }}
    .flowcard .fl {{ font-size:12px; font-weight:700; }}
    .flowcard .fv {{ font-weight:800; color:{INK}; font-variant-numeric:tabular-nums; font-size:17px; }}
    .sect {{ color:{FAINT}; font-size:12px; font-weight:700; letter-spacing:.04em; text-transform:uppercase; margin-bottom:10px; }}
    table.t {{ width:100%; border-collapse:collapse; font-size:14px; }}
    table.t th {{ text-align:left; color:{FAINT}; font-size:11px; text-transform:uppercase; font-weight:700; padding:10px 14px; border-bottom:1px solid {LINE}; }}
    table.t td {{ padding:11px 14px; border-bottom:1px solid {LINE_SOFT}; }}
    table.t tr:hover td {{ background:{JADE_SOFT}33; }}
    .acard {{ background:#fff; border:1px solid {LINE}; border-radius:16px; padding:14px; margin-bottom:6px;
              box-shadow:0 1px 2px rgba(16,24,40,.03); transition:box-shadow .15s ease; }}
    .acard:hover {{ box-shadow:0 4px 14px rgba(16,24,40,.07); }}
    .acard .aid {{ color:{FAINT}; font-size:12px; font-weight:700; font-variant-numeric:tabular-nums; }}
    .acard .anm {{ font-weight:800; }}
    .acard .acat {{ color:{SUB}; font-size:12px; }}
    .acard .aval {{ font-weight:800; color:{INK}; font-variant-numeric:tabular-nums; }}
    .unit-head {{ font-weight:800; display:flex; align-items:center; justify-content:space-between; margin-bottom:2px; }}

    /* 分頁籤 */
    .stTabs [data-baseweb="tab-list"] {{ gap:4px; }}
    .stTabs [data-baseweb="tab"] {{ font-weight:700; }}
    .stTabs [aria-selected="true"] {{ color:{JADE} !important; }}

    /* 採購單／資產：欄位對齊的可點列（扁平、相連） */
    div[data-testid="stVerticalBlock"]:has(#po-rows) .stButton > button,
    div[data-testid="stVerticalBlock"]:has(#as-rows) .stButton > button {{
        border:none !important; border-radius:0 !important;
        background:transparent !important; box-shadow:none !important; transform:none !important;
        text-align:left !important; justify-content:flex-start !important;
        padding:6px 4px !important; font-weight:700 !important; color:{INK} !important;
        font-variant-numeric:tabular-nums;
    }}
    div[data-testid="stVerticalBlock"]:has(#po-rows) .stButton > button:hover,
    div[data-testid="stVerticalBlock"]:has(#as-rows) .stButton > button:hover {{
        background:{JADE_SOFT}66 !important; color:{JADE} !important;
    }}
    div[data-testid="stVerticalBlock"]:has(#po-rows) .stButton > button[kind="primary"],
    div[data-testid="stVerticalBlock"]:has(#as-rows) .stButton > button[kind="primary"] {{
        background:transparent !important; color:{JADE} !important;
        box-shadow:inset 3px 0 0 {JADE} !important;
    }}
    /* 列分隔線 */
    div[data-testid="stVerticalBlock"]:has(#po-rows) div[data-testid="stHorizontalBlock"],
    div[data-testid="stVerticalBlock"]:has(#as-rows) div[data-testid="stHorizontalBlock"] {{
        border-bottom:1px solid {LINE_SOFT}; align-items:center; padding:2px 0;
    }}
    .po-th {{ color:{FAINT}; font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:.04em; padding:8px 4px; }}
    .po-td {{ color:{SUB}; font-size:14px; padding:6px 4px; }}
    </style>
    """, unsafe_allow_html=True)


def pill(text, style_map):
    bg, fg = style_map.get(text, (LINE_SOFT, SUB))
    return f"<span class='pill' style='background:{bg};color:{fg}'>{text}</span>"


def flash(msg):
    """暫存提示訊息，rerun 後以 toast 顯示。"""
    st.session_state["_flash"] = msg


# ============================ 登入 ============================
def require_login() -> bool:
    if st.session_state.get("authed"):
        return True
    st.markdown("<style>input[type=password]{-webkit-text-security:disc}</style>", unsafe_allow_html=True)
    _, mid, _ = st.columns([1, 1.6, 1])
    with mid:
        st.markdown("<div style='height:9vh'></div>", unsafe_allow_html=True)
        st.markdown(
            f"<div class='card' style='text-align:center;padding:40px 32px'>"
            f"<div style='width:54px;height:54px;border-radius:15px;background:{JADE};margin:0 auto 16px;"
            f"display:flex;align-items:center;justify-content:center;font-size:26px'>📦</div>"
            f"<div style='font-size:23px;font-weight:800;white-space:nowrap;letter-spacing:.01em;color:{INK}'>採購・資產整合平台</div>"
            f"<div style='color:{SUB};font-size:14px;margin-top:8px'>請輸入密碼登入</div></div>",
            unsafe_allow_html=True)
        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
        pw = st.text_input("密碼", type="password", label_visibility="collapsed", placeholder="請輸入密碼")
        if st.button("登入", type="primary", use_container_width=True):
            admin_pw = st.secrets.get("admin_password", "")
            user_pw = st.secrets.get("app_password", "")
            if admin_pw and pw == admin_pw:
                st.session_state.authed = True
                st.session_state.role = "admin"
                st.rerun()
            elif user_pw and pw == user_pw:
                st.session_state.authed = True
                st.session_state.role = "user"
                st.rerun()
            else:
                st.error("密碼錯誤")
    return False


def is_admin() -> bool:
    return st.session_state.get("role") == "admin"



# ============================ Google Sheets ============================
@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    key_path = os.path.join(os.path.dirname(__file__), ".streamlit", "service_account.json")
    if os.path.exists(key_path):                              # 1) 本機：讀 JSON 檔
        creds = Credentials.from_service_account_file(key_path, scopes=scopes)
    elif "service_account_json" in st.secrets:                # 2) 雲端：整段 JSON 字串（最省事）
        import json
        info = json.loads(st.secrets["service_account_json"])
        creds = Credentials.from_service_account_info(info, scopes=scopes)
    else:                                                     # 3) 雲端：[gcp_service_account] 分段
        creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=scopes)
    return gspread.authorize(creds).open_by_url(st.secrets["sheet"]["url"])


def _read_ws(name: str) -> pd.DataFrame:
    ss = get_spreadsheet()
    cols = SCHEMAS[name]
    try:
        records = ss.worksheet(name).get_all_records()
    except gspread.WorksheetNotFound:
        records = []
    df = pd.DataFrame(records)
    for col in cols:
        if col not in df.columns:
            df[col] = pd.Series(dtype="object")
    df = df[cols]
    for col in cols:
        if col in NUMERIC.get(name, []):
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = df[col].fillna("")
    return df


@st.cache_data(ttl=60, show_spinner=False)
def load_all() -> dict:
    return {name: _read_ws(name) for name in SCHEMAS}


def save(name: str, df: pd.DataFrame):
    ss = get_spreadsheet()
    df = df[SCHEMAS[name]]
    try:
        ws = ss.worksheet(name)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=name, rows=300, cols=max(8, len(SCHEMAS[name])))
    ws.clear()
    values = [df.columns.tolist()] + df.fillna("").astype(object).values.tolist()
    ws.append_rows(values, value_input_option="USER_ENTERED")
    load_all.clear()


def init_sheets():
    ss = get_spreadsheet()
    existing = [w.title for w in ss.worksheets()]
    for name, cols in SCHEMAS.items():
        if name not in existing:
            ss.add_worksheet(title=name, rows=300, cols=max(8, len(cols)))
    save("suppliers", _seed_suppliers())
    save("purchase_orders", _seed_pos())
    save("po_items", _seed_items())
    save("assets", _seed_assets())
    save("units", pd.DataFrame({"name": DEFAULT_UNITS}))


# ============================ 種子資料 ============================
def _seed_suppliers():
    return pd.DataFrame([
        ["S1", "全美電腦資訊", "12345675", "王經理", "02-2222-3333", "配合多年，可月結"],
        ["S2", "永興辦公家具", "23456781", "林小姐", "02-2555-6666", ""],
        ["S3", "聲學音響工程", "34567892", "陳先生", "03-3333-4444", "報價含安裝"],
        ["S4", "印刷大師文宣", "45678903", "李主任", "02-2888-9999", ""],
    ], columns=SCHEMAS["suppliers"])


def _seed_pos():
    return pd.DataFrame([
        ["PO-2026-001", "S1", "2026-04-12", "已驗收", "辦公室設備汰換", "王小明", ""],
        ["PO-2026-002", "S2", "2026-05-03", "已驗收", "主日場地桌椅", "李美華", ""],
        ["PO-2026-003", "S3", "2026-06-01", "待驗收", "敬拜團音響升級", "陳大同", "含現場安裝"],
        ["PO-2026-004", "S4", "2026-06-06", "草稿", "主日文宣印製", "王小明", ""],
    ], columns=SCHEMAS["purchase_orders"])


def _seed_items():
    return pd.DataFrame([
        ["PO-2026-001", "筆記型電腦", "3C設備", 2, 28000],
        ["PO-2026-001", "短焦投影機", "3C設備", 2, 24000],
        ["PO-2026-001", "雷射印表機", "3C設備", 1, 15000],
        ["PO-2026-002", "折疊長桌", "辦公家具", 4, 1800],
        ["PO-2026-002", "堆疊摺疊椅", "辦公家具", 8, 500],
        ["PO-2026-003", "無線麥克風", "音響設備", 4, 3500],
        ["PO-2026-003", "數位混音器", "音響設備", 1, 18000],
        ["PO-2026-004", "主日文宣海報", "文宣品", 500, 12],
    ], columns=SCHEMAS["po_items"])


def _seed_assets():
    rows = []
    seq = {"A26": 1, "B26": 1}

    def mk(name, cat, val, po, d, aclass, unit, status="使用中"):
        prefix = CLASS_PREFIX[aclass]
        rows.append([f"{prefix}-{seq[prefix]:04d}", name, cat, val, po, d, aclass, unit, status])
        seq[prefix] += 1

    mk("伺服器主機", "3C設備", 95000, "PO-2026-001", "2026-04-18", "列帳資產", "行政部")
    mk("筆記型電腦", "3C設備", 28000, "PO-2026-001", "2026-04-18", "列管資產", "行政部")
    mk("筆記型電腦", "3C設備", 28000, "PO-2026-001", "2026-04-18", "列管資產", "影音部", "維修中")
    mk("短焦投影機", "3C設備", 24000, "PO-2026-001", "2026-04-18", "列管資產", "傳道部")
    mk("專業混音器", "音響設備", 88000, "PO-2026-001", "2026-04-18", "列帳資產", "影音部")
    mk("雷射印表機", "3C設備", 15000, "PO-2026-001", "2026-04-18", "列管資產", "行政部")
    for i in range(4):
        mk("折疊長桌", "辦公家具", 1800, "PO-2026-002", "2026-05-09", "列管資產", UNITS[i % len(UNITS)])
    for i in range(6):
        mk("堆疊摺疊椅", "辦公家具", 500, "PO-2026-002", "2026-05-09", "列管資產", UNITS[i % len(UNITS)],
           "已報廢" if i == 5 else "使用中")
    return pd.DataFrame(rows, columns=SCHEMAS["assets"])


# ============================ 工具 ============================
def nt(v) -> str:
    try:
        return "NT$" + format(int(round(float(v))), ",")
    except (TypeError, ValueError):
        return "NT$0"


def next_po_id(pos):
    yr = date.today().year
    seq = len(pos[pos["id"].str.startswith(f"PO-{yr}-", na=False)]) + 1
    return f"PO-{yr}-{seq:03d}"


def has_pending_void(data, target_type, target_id):
    """該採購單／資產是否已有待審的作廢申請。"""
    vr = data.get("void_requests")
    if vr is None or vr.empty:
        return False
    m = ((vr["target_type"] == target_type) & (vr["target_id"].astype(str) == str(target_id))
         & (vr["status"] == "待審核"))
    return bool(m.any())


def submit_void_request(data, target_type, target_id, reason, by="前台"):
    vr = data["void_requests"]
    rid = f"VR-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    row = pd.DataFrame([[rid, target_type, str(target_id), reason.strip(), by,
                         datetime.now().strftime("%Y-%m-%d %H:%M"), "待審核"]],
                       columns=SCHEMAS["void_requests"])
    save("void_requests", pd.concat([vr, row], ignore_index=True))


def next_asset_seq(assets, prefix="A26"):
    if assets.empty:
        return 1
    nums = assets["id"].astype(str).str.extract(rf"{prefix}-(\d+)")[0].dropna().astype(int)
    return (nums.max() + 1) if len(nums) else 1


def new_asset_id(assets, asset_class, used=None):
    """依分類取下一個編號：列管 A26-xxxx／列帳 B26-xxxx。used 為本批已配發的編號集合。"""
    prefix = CLASS_PREFIX.get(asset_class, "A26")
    seq = next_asset_seq(assets, prefix)
    used = used or set()
    while f"{prefix}-{seq:04d}" in used:
        seq += 1
    return f"{prefix}-{seq:04d}", prefix


def po_total(items, po_id):
    rows = items[items["po_id"] == po_id]
    return float((rows["qty"] * rows["price"]).sum())


def sup_name(suppliers, sid):
    m = suppliers[suppliers["id"] == sid]
    return m.iloc[0]["name"] if len(m) else "—"


# ============================ 儀表板 ============================
def page_dashboard(data):
    pos, items, assets, sups = data["purchase_orders"], data["po_items"], data["assets"], data["suppliers"]
    vr = data.get("void_requests")
    pending = pos[pos["status"] == "待驗收"]["id"].tolist()
    pending_amt = sum(po_total(items, p) for p in pending)
    pending_void = int((vr["status"] == "待審核").sum()) if (vr is not None and not vr.empty) else 0

    live = assets[~assets["status"].isin(["已報廢", "已作廢"])]
    ledger = live[live["asset_type"] == "列帳資產"]
    managed = live[live["asset_type"] == "列管資產"]
    ledger_val, managed_val = ledger["value"].sum(), managed["value"].sum()

    # 本年度採購金額（已驗收＋待驗收，依採購日期判斷今年）
    yr = str(date.today().year)
    year_pos = pos[(pos["status"].isin(["已驗收", "待驗收"]))
                   & (pos["date"].astype(str).str.startswith(yr))]
    year_ids = year_pos["id"].tolist()
    year_spend = sum(po_total(items, p) for p in year_ids)
    year_items = items[items["po_id"].isin(year_ids)]

    st.markdown("<h1>儀表板</h1><p style='color:#5A6472;margin-top:-8px'>資產總覽與待辦一覽</p>", unsafe_allow_html=True)

    def stat(label, value, accent, sub=""):
        return (f"<div class='card'><div class='stat-label'>{label}</div>"
                f"<div class='stat-value' style='color:{accent}'>{value}</div>"
                f"<div class='stat-sub'>{sub}</div></div>")

    # 第一排：資產總覽（筆數＋總值）
    st.markdown(
        "<div class='grid4'>"
        + stat("列帳資產", nt(ledger_val), INDIGO, f"{len(ledger)} 筆 ・ B26 ・ ≥ 8 萬")
        + stat("列管資產", nt(managed_val), JADE, f"{len(managed)} 筆 ・ A26")
        + stat("資產總值", nt(ledger_val + managed_val), INK, f"{len(live)} 筆有效資產")
        + stat("資產總數", f"{len(live)}", SUB, "排除已報廢／已作廢")
        + "</div>", unsafe_allow_html=True)

    # 本年度採購金額（預算分佈）— 僅管理者可見
    if is_admin():
        st.markdown(f"<div style='height:14px'></div><div class='sect'>{yr} 年度採購金額（已驗收＋待驗收）</div>", unsafe_allow_html=True)
        bl, br = st.columns([2, 3])
        with bl:
            st.markdown(
                f"<div class='card' style='background:linear-gradient(135deg,{JADE} 0%,#0c655e 100%);border:none'>"
                f"<div style='color:#CFE8E4;font-size:12px;font-weight:700;letter-spacing:.04em'>本年度採購總額</div>"
                f"<div style='color:#fff;font-size:32px;font-weight:800;margin-top:6px;font-variant-numeric:tabular-nums'>{nt(year_spend)}</div>"
                f"<div style='color:#CFE8E4;font-size:12px;margin-top:4px'>{len(year_ids)} 張採購單 ・ 含已下單未驗收</div></div>",
                unsafe_allow_html=True)
        with br:
            with st.container(border=True):
                st.markdown("<div class='sect'>各類別花費（看預算分佈）</div>", unsafe_allow_html=True)
                if year_items.empty:
                    st.caption(f"{yr} 年度尚無採購紀錄")
                else:
                    cat_spend = []
                    for cat in CATEGORIES:
                        sub = year_items[year_items["category"] == cat]
                        amt = int((sub["qty"] * sub["price"]).sum())
                        if amt > 0:
                            cat_spend.append((cat, amt))
                    cat_spend.sort(key=lambda x: x[1], reverse=True)
                    total = sum(a for _, a in cat_spend) or 1
                    bars = ""
                    for cat, amt in cat_spend:
                        pct = amt / total * 100
                        bars += (f"<div style='margin-bottom:10px'>"
                                 f"<div style='display:flex;justify-content:space-between;font-size:13px;margin-bottom:3px'>"
                                 f"<span style='font-weight:600'>{cat}</span>"
                                 f"<span style='color:{SUB};font-variant-numeric:tabular-nums'>{nt(amt)}　({pct:.0f}%)</span></div>"
                                 f"<div style='background:{LINE_SOFT};border-radius:6px;height:8px;overflow:hidden'>"
                                 f"<div style='background:{JADE};width:{pct:.1f}%;height:100%'></div></div></div>")
                    st.markdown(bars, unsafe_allow_html=True)

    # 第二排：待辦事項（待驗收、待審作廢）＋ 資產狀態
    using = int((live["status"] == "使用中").sum())
    repair = int((assets["status"] == "維修中").sum())
    scrap = int((assets["status"] == "已報廢").sum())

    def todo(label, n, accent, sub):
        badge = f"<span style='background:{accent};color:#fff;border-radius:999px;padding:1px 9px;font-size:13px;font-weight:800'>{n}</span>" if n else f"<span style='color:{FAINT};font-weight:800'>0</span>"
        return (f"<div class='card' style='display:flex;justify-content:space-between;align-items:center'>"
                f"<div><div style='font-weight:700'>{label}</div><div class='stat-sub'>{sub}</div></div>{badge}</div>")

    st.markdown("<div style='height:14px'></div><div class='sect'>待辦事項</div>", unsafe_allow_html=True)
    if is_admin():
        st.markdown(
            "<div class='grid2'>"
            + todo("待驗收採購單", len(pending), AMBER, f"金額合計 {nt(pending_amt)}")
            + todo("待審核作廢申請", pending_void, DANGER, "前台送出、等你核准")
            + "</div>", unsafe_allow_html=True)
    else:
        st.markdown("<div class='grid2'>"
                    + todo("待驗收採購單", len(pending), AMBER, f"金額合計 {nt(pending_amt)}")
                    + "</div>", unsafe_allow_html=True)

    st.markdown("<div style='height:14px'></div><div class='sect'>資產狀態</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='grid4'>"
        + stat("使用中", f"{using}", JADE, "有效資產")
        + stat("維修中", f"{repair}", AMBER, "需追蹤")
        + stat("已報廢", f"{scrap}", DANGER, "")
        + stat("供應商", f"{len(sups)}", SUB, "合作廠商數")
        + "</div>", unsafe_allow_html=True)

    # 第三排：各單位資產值長條圖 ＋ 待驗收清單
    left, right = st.columns([3, 2])
    with left:
        with st.container(border=True):
            st.markdown("<div class='sect'>各單位資產值（列帳＋列管）</div>", unsafe_allow_html=True)
            rows = []
            for u in UNITS:
                sub = live[live["unit"] == u]
                rows.append({"單位": u,
                             "列帳資產": sub[sub["asset_type"] == "列帳資產"]["value"].sum(),
                             "列管資產": sub[sub["asset_type"] == "列管資產"]["value"].sum()})
            df = pd.DataFrame(rows).set_index("單位")
            if df.values.sum() == 0:
                st.caption("尚無資產資料")
            else:
                st.bar_chart(df, color=[INDIGO, JADE], height=260)
    with right:
        with st.container(border=True):
            st.markdown("<div class='sect'>待驗收採購單</div>", unsafe_allow_html=True)
            if not pending:
                st.caption("沒有待驗收的採購單")
            for p in pending:
                row = pos[pos["id"] == p].iloc[0]
                st.markdown(
                    f"<div style='border:1px solid {LINE};border-radius:14px;padding:12px;margin-bottom:8px;"
                    f"display:flex;justify-content:space-between;align-items:center'>"
                    f"<div><div style='font-weight:700'>{p}</div>"
                    f"<div style='color:{SUB};font-size:12px'>{sup_name(sups, row['supplier_id'])}</div></div>"
                    f"<div style='text-align:right'><div style='font-weight:800;font-variant-numeric:tabular-nums'>{nt(po_total(items, p))}</div>"
                    f"{pill('待驗收', STATUS_STYLE)}</div></div>", unsafe_allow_html=True)


# ============================ 採購 ============================
def _po_detail(data, sel, pos, items, assets, sups):
    po = pos[pos["id"] == sel].iloc[0]
    its = items[items["po_id"] == sel].copy()
    meta = f"{sup_name(sups, po['supplier_id'])} ・ 採購日期 {po['date']}"
    if po.get("buyer"):
        meta += f" ・ 採購人員 {po['buyer']}"
    purpose_html = f"<div style='margin:.2rem 0'><b>採購用途：</b>{po.get('purpose') or '—'}</div>"
    note_html = f"<div style='color:{SUB}'><b>備註：</b>{po.get('note')}</div>" if po.get("note") else ""
    st.markdown(f"<p style='color:{SUB};margin:.1rem 0 .3rem'>{meta}</p>{purpose_html}{note_html}",
                unsafe_allow_html=True)
    rowhtml = ""
    for it in its.itertuples():
        rowhtml += (f"<tr><td style='font-weight:600'>{it.name}</td><td style='color:{SUB}'>{it.category}</td>"
                    f"<td style='text-align:right;font-variant-numeric:tabular-nums'>{int(it.qty)}</td>"
                    f"<td style='text-align:right;font-variant-numeric:tabular-nums'>{nt(it.price)}</td>"
                    f"<td style='text-align:right;font-weight:700;font-variant-numeric:tabular-nums'>{nt(it.qty * it.price)}</td></tr>")
    st.markdown(
        "<table class='t'><tr><th>品項</th><th>類別</th><th style='text-align:right'>數量</th>"
        f"<th style='text-align:right'>單價</th><th style='text-align:right'>小計</th></tr>{rowhtml}</table>"
        f"<div style='text-align:right;margin-top:10px;font-size:18px;font-weight:800;color:{AMBER}'>合計 {nt(po_total(items, sel))}</div>",
        unsafe_allow_html=True)
    if po["status"] == "待驗收":
        _receive_form(sel, its, pos, assets)

    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    if po["status"] == "已作廢":
        st.caption("此採購單已作廢。")
    elif has_pending_void(data, "採購單", sel):
        st.info("⏳ 此採購單已有待審核的作廢申請，請等待管理者處理。")
    else:
        with st.expander("🗑️ 申請作廢這張採購單"):
            st.caption("送出後不會立即作廢，需由管理者在「作廢申請」頁審核。")
            vr_reason = st.text_area("作廢原因", key=f"vreason_po_{sel}", placeholder="例如：供應商或金額填錯、重複建立…")
            if st.button("送出作廢申請", key=f"vbtn_po_{sel}", disabled=not vr_reason.strip()):
                submit_void_request(data, "採購單", sel, vr_reason)
                flash("已送出作廢申請，待管理者審核")
                st.rerun()


def page_procurement(data):
    pos, items, assets, sups = data["purchase_orders"], data["po_items"], data["assets"], data["suppliers"]
    st.markdown("<h1>採購</h1><p style='color:#5A6472;margin-top:-8px'>採購單管理；驗收時分類為列管／列帳資產或一般耗材</p>", unsafe_allow_html=True)
    tab_list, tab_query, tab_new = st.tabs(["採購單清單", "🔍 品項查詢", "＋ 新增採購單"])

    with tab_list:
        q = st.text_input("搜尋", placeholder="輸入單號、採購人員、供應商名稱或統一編號查詢…", label_visibility="collapsed")
        rows = pos
        if q.strip():
            kw = q.strip()
            name_map = {str(s.id): s.name for s in sups.itertuples()}
            tax_map = {str(s.id): str(s.tax_id) for s in sups.itertuples()}
            sid = pos["supplier_id"].astype(str)
            sname = sid.map(name_map).fillna("")
            stax = sid.map(tax_map).fillna("")
            rows = pos[pos["id"].astype(str).str.contains(kw, case=False, na=False)
                       | pos["buyer"].astype(str).str.contains(kw, case=False, na=False)
                       | sname.str.contains(kw, case=False, na=False)
                       | stax.str.contains(kw, case=False, na=False)]
        rows = rows.sort_values("date", ascending=False)
        st.caption(f"{len(rows)} 張採購單　·　點任一列展開明細")

        DOT = {"已驗收": "🟢", "待驗收": "🟠", "草稿": "⚪", "已作廢": "🔴"}
        WIDTHS = [1.3, 2, 1, 1.2, 1.2, 1.1]
        cur = st.session_state.get("po_sel")
        if rows.empty:
            st.info("找不到符合的採購單")
        else:
            with st.container():
                st.markdown("<span id='po-rows'></span>", unsafe_allow_html=True)
                # 標題列
                h = st.columns(WIDTHS)
                for col, t, al in zip(h, ["單號", "供應商", "採購人員", "採購日期", "金額", "狀態"],
                                      ["left", "left", "left", "left", "right", "left"]):
                    col.markdown(f"<div class='po-th' style='text-align:{al}'>{t}</div>", unsafe_allow_html=True)
                # 資料列
                for r in rows.itertuples():
                    buyer = getattr(r, "buyer", "") or "—"
                    dot = DOT.get(r.status, "⚪")
                    c = st.columns(WIDTHS)
                    sel_now = (r.id == cur)
                    if c[0].button(("▾ " if sel_now else "▸ ") + str(r.id), key=f"porow_{r.id}", use_container_width=True,
                                   type="primary" if sel_now else "secondary"):
                        st.session_state.po_sel = None if sel_now else r.id
                        st.rerun()
                    c[1].markdown(f"<div class='po-td'>{sup_name(sups, r.supplier_id)}</div>", unsafe_allow_html=True)
                    c[2].markdown(f"<div class='po-td'>{buyer}</div>", unsafe_allow_html=True)
                    c[3].markdown(f"<div class='po-td' style='font-variant-numeric:tabular-nums'>{r.date}</div>", unsafe_allow_html=True)
                    c[4].markdown(f"<div class='po-td' style='text-align:right;font-weight:800;font-variant-numeric:tabular-nums'>{nt(po_total(items, r.id))}</div>", unsafe_allow_html=True)
                    c[5].markdown(f"<div class='po-td'>{dot} {r.status}</div>", unsafe_allow_html=True)
                    if sel_now:
                        with st.container(border=True):
                            _po_detail(data, r.id, pos, items, assets, sups)

    with tab_query:
        st.markdown(f"<div style='color:{SUB};font-size:14px;margin-bottom:8px'>輸入關鍵字，查詢歷次採購紀錄。會比對「品名」與「採購用途」（例如打「影印」找出影印紙；打「主日書房」找出用途含主日書房的採購）。</div>", unsafe_allow_html=True)
        iq = st.text_input("品項關鍵字", placeholder="例如：影印紙、麥克風、主日書房…", label_visibility="collapsed")
        if not iq.strip():
            st.caption("請輸入關鍵字開始查詢。")
        else:
            kw = iq.strip()
            pinfo = pos.set_index("id")
            # 比對品名，或所屬採購單的採購用途
            purpose_map = pinfo["purpose"] if "purpose" in pinfo.columns else pd.Series(dtype=str)
            item_purpose = items["po_id"].map(lambda x: str(purpose_map.get(x, "")) if len(purpose_map) else "")
            name_hit = items["name"].astype(str).str.contains(kw, case=False, na=False)
            purpose_hit = item_purpose.str.contains(kw, case=False, na=False)
            hit = items[name_hit | purpose_hit].copy()
            if hit.empty:
                st.info(f"找不到品名或採購用途包含「{kw}」的採購紀錄。")
            else:
                # 併入採購單的日期／供應商／用途
                hit["日期"] = hit["po_id"].map(lambda x: pinfo["date"].get(x, ""))
                hit["供應商"] = hit["po_id"].map(lambda x: sup_name(sups, pinfo["supplier_id"].get(x, "")))
                hit["用途"] = hit["po_id"].map(lambda x: str(purpose_map.get(x, "")) if len(purpose_map) else "")
                hit["狀態"] = hit["po_id"].map(lambda x: pinfo["status"].get(x, ""))
                hit["小計"] = hit["qty"] * hit["price"]
                hit = hit.sort_values("日期", ascending=False)

                # 總計摘要
                times = len(hit)
                total_qty = int(hit["qty"].sum())
                total_amt = int(hit["小計"].sum())
                avg_price = (total_amt / total_qty) if total_qty else 0

                def stat(label, value, accent, sub=""):
                    return (f"<div class='card'><div class='stat-label'>{label}</div>"
                            f"<div class='stat-value' style='color:{accent}'>{value}</div>"
                            f"<div class='stat-sub'>{sub}</div></div>")
                st.markdown(
                    "<div class='grid4'>"
                    + stat("採購次數", f"{times}", INDIGO, "符合的明細筆數")
                    + stat("累計數量", f"{total_qty}", JADE, "歷次加總")
                    + stat("累計金額", nt(total_amt), AMBER, "歷次小計加總")
                    + stat("平均單價", nt(round(avg_price)), INK, "總金額 ÷ 總量")
                    + "</div><div style='height:12px'></div>", unsafe_allow_html=True)

                # 明細表
                head = ("<tr><th>採購日期</th><th>品名</th><th>採購用途</th><th>供應商</th><th style='text-align:right'>數量</th>"
                        "<th style='text-align:right'>單價</th><th style='text-align:right'>小計</th><th>單號</th></tr>")
                body = ""
                for r in hit.itertuples():
                    body += (f"<tr><td style='font-variant-numeric:tabular-nums'>{r.日期}</td>"
                             f"<td style='font-weight:600'>{r.name}</td><td style='color:{SUB}'>{r.用途 or '—'}</td>"
                             f"<td style='color:{SUB}'>{r.供應商}</td>"
                             f"<td style='text-align:right;font-variant-numeric:tabular-nums'>{int(r.qty)}</td>"
                             f"<td style='text-align:right;font-variant-numeric:tabular-nums'>{nt(r.price)}</td>"
                             f"<td style='text-align:right;font-weight:700;font-variant-numeric:tabular-nums'>{nt(r.小計)}</td>"
                             f"<td style='color:{SUB}'>{r.po_id}</td></tr>")
                st.markdown(f"<div class='card' style='padding:4px'><table class='t'>{head}{body}</table></div>", unsafe_allow_html=True)

    with tab_new:
        sup_map = {sup_name(sups, r.id): r.id for r in sups.itertuples()}
        if not sup_map:
            st.warning("尚無供應商，請先到「供應商」頁新增。")
            return
        chosen = st.selectbox("供應商", list(sup_map.keys()))
        cc1, cc2 = st.columns(2)
        buyer = cc1.text_input("採購人員", placeholder="例：王小明")
        purchase_date = cc2.date_input("採購日期", value=date.today(), format="YYYY-MM-DD")
        purpose = st.text_area("採購用途（必填）", placeholder="例：敬拜團音響升級，汰換老舊混音器與喇叭", height=120)
        note = st.text_area("備註", placeholder="選填", height=68)
        st.caption("在下表新增品項（點最後一列空白處可新增）")
        edited = st.data_editor(
            pd.DataFrame([{"品名": "", "類別": "3C設備", "數量": 1, "單價": 0}]),
            num_rows="dynamic", use_container_width=True, hide_index=True, key="new_po_items",
            column_config={
                "類別": st.column_config.SelectboxColumn(options=CATEGORIES, required=True),
                "數量": st.column_config.NumberColumn(min_value=1, step=1),
                "單價": st.column_config.NumberColumn(min_value=0, step=100),
            })
        valid = edited[edited["品名"].astype(str).str.strip() != ""]
        ok = (not valid.empty) and bool(purpose.strip())
        if not purpose.strip():
            st.caption("⚠️ 採購用途為必填")
        c1, c2 = st.columns(2)
        if c1.button("存為草稿", disabled=not ok, use_container_width=True):
            _create_po(chosen, sup_map, valid, "草稿", pos, items, purpose, buyer, purchase_date, note)
        if c2.button("送出（待驗收）", type="primary", disabled=not ok, use_container_width=True):
            _create_po(chosen, sup_map, valid, "待驗收", pos, items, purpose, buyer, purchase_date, note)


def _create_po(chosen, sup_map, valid, status, pos, items, purpose="", buyer="", purchase_date=None, note=""):
    new_id = next_po_id(pos)
    dstr = purchase_date.isoformat() if purchase_date else date.today().isoformat()
    new_po = pd.DataFrame([[new_id, sup_map[chosen], dstr, status,
                            purpose.strip(), (buyer or "").strip(), (note or "").strip()]], columns=SCHEMAS["purchase_orders"])
    new_items = pd.DataFrame({"po_id": new_id, "name": valid["品名"].values, "category": valid["類別"].values,
                              "qty": valid["數量"].values, "price": valid["單價"].values})[SCHEMAS["po_items"]]
    save("purchase_orders", pd.concat([pos, new_po], ignore_index=True))
    save("po_items", pd.concat([items, new_items], ignore_index=True))
    flash(f"採購單 {new_id} 已建立（{status}）")
    st.rerun()


def _receive_form(po_id, its, pos, assets):
    st.markdown(f"<div class='sect' style='margin-top:14px'>驗收入庫</div>", unsafe_allow_html=True)
    st.caption(f"逐項選擇分類（單價 ≥ {nt(LEDGER_THRESHOLD)} 預設「列帳資產」）。列管→A26、列帳→B26 給編號；一般耗材不列管、不入庫。")
    rows_all = its.reset_index(drop=True)
    choices = {}
    for i, r in enumerate(rows_all.itertuples()):
        with st.container(border=True):
            st.markdown(f"<b>{r.name}</b> <span style='color:{FAINT};font-size:12px'>×{int(r.qty)} ・ {nt(r.price)} ・ {r.category}</span>", unsafe_allow_html=True)
            c1, c2 = st.columns([1.4, 1])
            default = "列帳資產" if r.price >= LEDGER_THRESHOLD else "列管資產"
            aclass = c1.radio("分類", ASSET_CLASS_OPTS, index=ASSET_CLASS_OPTS.index(default),
                              key=f"t_{po_id}_{i}", horizontal=True, label_visibility="collapsed")
            unit = c2.selectbox("單位", UNITS, key=f"u_{po_id}_{i}", label_visibility="collapsed")
            choices[i] = (aclass, unit)
    if st.button("確認入庫", type="primary", key=f"recv_{po_id}", use_container_width=True):
        st.session_state[f"recv_confirm_{po_id}"] = True
        st.rerun()

    # 確認摘要：按下「確認入庫」後先跳出，需勾選＋再按一次才真正寫入
    if st.session_state.get(f"recv_confirm_{po_id}"):
        # 預先計算將入庫的內容
        ledger_names, managed_names, skip_names = [], [], []
        for i, r in enumerate(rows_all.itertuples()):
            aclass, unit = choices[i]
            tag = f"{r.name} ×{int(r.qty)}（{unit}）"
            if aclass == "列帳資產":
                ledger_names.append(tag)
            elif aclass == "列管資產":
                managed_names.append(tag)
            else:
                skip_names.append(f"{r.name} ×{int(r.qty)}")
        n_ledger = sum(int(r.qty) for i, r in enumerate(rows_all.itertuples()) if choices[i][0] == "列帳資產")
        n_managed = sum(int(r.qty) for i, r in enumerate(rows_all.itertuples()) if choices[i][0] == "列管資產")

        with st.container(border=True):
            st.markdown(f"<div style='font-weight:800;color:{AMBER};font-size:15px'>⚠️ 即將入庫，請再次確認分類</div>", unsafe_allow_html=True)
            lines = f"本次將新增 <b>{n_ledger + n_managed}</b> 筆資產："
            if ledger_names:
                lines += f"<br>・<b style='color:{INDIGO}'>列帳資產（B26）{n_ledger} 筆</b>：{'、'.join(ledger_names)}"
            if managed_names:
                lines += f"<br>・<b style='color:{JADE}'>列管資產（A26）{n_managed} 筆</b>：{'、'.join(managed_names)}"
            if skip_names:
                lines += f"<br>・<span style='color:{SUB}'>不列管（一般耗材）略過：{'、'.join(skip_names)}</span>"
            st.markdown(f"<div style='font-size:14px;margin-top:6px;line-height:1.8'>{lines}</div>", unsafe_allow_html=True)
            st.markdown(f"<div style='color:{DANGER};font-size:13px;margin-top:8px'>確認後將正式寫入資產清單，入庫後不可復原（需作廢才能移除）。</div>", unsafe_allow_html=True)

            ck = st.checkbox("我已確認分類無誤", key=f"recv_ck_{po_id}")
            b1, b2 = st.columns(2)
            if b1.button("✅ 確認入庫", type="primary", key=f"recv_go_{po_id}", disabled=not ck, use_container_width=True):
                new_rows, used = [], set()
                for i, r in enumerate(rows_all.itertuples()):
                    aclass, unit = choices[i]
                    if aclass not in ASSET_RECORD_CLASSES:      # 一般耗材：不入庫
                        continue
                    for _ in range(int(r.qty)):
                        aid, _p = new_asset_id(assets, aclass, used)
                        used.add(aid)
                        new_rows.append([aid, r.name, r.category, r.price, po_id,
                                         date.today().isoformat(), aclass, unit, "使用中"])
                if new_rows:
                    save("assets", pd.concat([assets, pd.DataFrame(new_rows, columns=SCHEMAS["assets"])], ignore_index=True))
                save("purchase_orders", pos.assign(status=pos["status"].where(pos["id"] != po_id, "已驗收")))
                nb = sum(1 for r in new_rows if r[6] == "列帳資產")
                na = sum(1 for r in new_rows if r[6] == "列管資產")
                st.session_state[f"recv_confirm_{po_id}"] = False
                flash(f"{po_id} 已驗收入庫：列帳資產 {nb} 筆、列管資產 {na} 筆")
                st.rerun()
            if b2.button("取消", key=f"recv_cancel_{po_id}", use_container_width=True):
                st.session_state[f"recv_confirm_{po_id}"] = False
                st.rerun()


# ============================ 資產 ============================
def _asset_detail(a, assets):
    st.markdown(
        f"<div style='display:flex;gap:10px;flex-wrap:wrap;margin-bottom:8px'>"
        f"{pill(a['asset_type'], ATYPE_STYLE)}{pill(a['status'], ASTATUS_STYLE)}</div>"
        f"<table class='t'>"
        f"<tr><td style='color:{SUB};width:90px'>財產編號</td><td style='font-weight:700'>{a['id']}</td></tr>"
        f"<tr><td style='color:{SUB}'>品名</td><td>{a['name']}</td></tr>"
        f"<tr><td style='color:{SUB}'>類別</td><td>{a['category']}</td></tr>"
        f"<tr><td style='color:{SUB}'>價值</td><td style='font-variant-numeric:tabular-nums'>{nt(a['value'])}</td></tr>"
        f"<tr><td style='color:{SUB}'>取得日期</td><td>{a['acquired']}</td></tr>"
        f"<tr><td style='color:{SUB}'>來源</td><td>{a['source_po']}</td></tr>"
        f"<tr><td style='color:{SUB}'>使用單位</td><td>{a['unit']}</td></tr></table>",
        unsafe_allow_html=True)

    if a["status"] == "已作廢":
        st.caption("此資產已作廢。")
        return
    if not is_admin():
        return  # 一般使用者：只能瀏覽

    st.markdown(f"<div class='sect' style='margin-top:12px'>編輯</div>", unsafe_allow_html=True)
    u, s = st.columns(2)
    nu = u.selectbox("使用單位", UNITS, index=UNITS.index(a["unit"]) if a["unit"] in UNITS else 0, key=f"lu_{a['id']}")
    ns = s.selectbox("狀態", ASSET_STATUS_OPTS, index=ASSET_STATUS_OPTS.index(a["status"]) if a["status"] in ASSET_STATUS_OPTS else 0, key=f"ls_{a['id']}")
    if nu != a["unit"] or ns != a["status"]:
        upd = assets.set_index("id")
        upd.loc[a["id"], "unit"] = nu
        upd.loc[a["id"], "status"] = ns
        save("assets", upd.reset_index()[SCHEMAS["assets"]])
        flash(f"已更新資產 {a['id']}")
        st.rerun()

    with st.expander("🗑️ 作廢這筆資產"):
        st.caption("作廢後資料保留、僅標記為已作廢，不會出現在統計與單位配置。")
        ck = st.checkbox("我確認要作廢這筆資產", key=f"vck_as_{a['id']}")
        if st.button("確認作廢", key=f"vbtn_as_{a['id']}", disabled=not ck):
            upd = assets.set_index("id")
            upd.loc[a["id"], "status"] = "已作廢"
            save("assets", upd.reset_index()[SCHEMAS["assets"]])
            flash(f"資產 {a['id']} 已作廢")
            st.rerun()


def page_assets(data):
    assets = data["assets"]
    st.markdown("<h1>資產</h1><p style='color:#5A6472;margin-top:-8px'>所有列管資產皆可追溯來源採購單</p>", unsafe_allow_html=True)
    tab_list, tab_board, tab_add = st.tabs(["資產清單", "單位配置（列管＋列帳）", "＋ 手動新增資產"])

    with tab_list:
        aq = st.text_input("資產搜尋", placeholder="輸入資產名稱或財產編號（A26/B26）查詢…", label_visibility="collapsed")
        c1, c2, c3, c4 = st.columns(4)
        f_type = c1.selectbox("查詢範圍", ["總資產", "列管資產", "列帳資產"])
        f_cat = c2.selectbox("類別", ["全部"] + CATEGORIES)
        f_unit = c3.selectbox("使用單位", ["全部"] + UNITS)
        f_status = c4.selectbox("狀態", ["全部"] + ASSET_STATUS_OPTS)
        mask = pd.Series(True, index=assets.index)
        if aq.strip():
            kw = aq.strip()
            mask &= (assets["name"].astype(str).str.contains(kw, case=False, na=False)
                     | assets["id"].astype(str).str.contains(kw, case=False, na=False))
        if f_type != "總資產":  mask &= assets["asset_type"] == f_type
        if f_cat != "全部":    mask &= assets["category"] == f_cat
        if f_unit != "全部":   mask &= assets["unit"] == f_unit
        if f_status != "全部": mask &= assets["status"] == f_status
        view = assets[mask]
        total_val = view[view["status"] != "已報廢"]["value"].sum()
        st.caption(f"{len(view)} 筆 ・ 合計（排除已報廢）{nt(total_val)}")

        if view.empty:
            st.info("沒有符合條件的資產")
        else:
            ADOT = {"使用中": "🟢", "維修中": "🟠", "已報廢": "⚫", "已作廢": "🔴"}
            CDOT = {"列帳資產": "🔵", "列管資產": "🟢"}
            AW = [1.3, 2, 1.2, 1, 1.1, 1]
            cur = st.session_state.get("as_sel")
            with st.container():
                st.markdown("<span id='as-rows'></span>", unsafe_allow_html=True)
                h = st.columns(AW)
                for col, t, al in zip(h, ["財產編號", "品名", "分類", "單位", "價值", "狀態"],
                                      ["left", "left", "left", "left", "right", "left"]):
                    col.markdown(f"<div class='po-th' style='text-align:{al}'>{t}</div>", unsafe_allow_html=True)
                for a in view.to_dict("records"):
                    sel_now = (a["id"] == cur)
                    c = st.columns(AW)
                    if c[0].button(("▾ " if sel_now else "▸ ") + str(a["id"]), key=f"asrow_{a['id']}",
                                   use_container_width=True, type="primary" if sel_now else "secondary"):
                        st.session_state.as_sel = None if sel_now else a["id"]
                        st.rerun()
                    c[1].markdown(f"<div class='po-td'>{a['name']}</div>", unsafe_allow_html=True)
                    c[2].markdown(f"<div class='po-td'>{CDOT.get(a['asset_type'],'')} {a['asset_type']}</div>", unsafe_allow_html=True)
                    c[3].markdown(f"<div class='po-td'>{a['unit']}</div>", unsafe_allow_html=True)
                    c[4].markdown(f"<div class='po-td' style='text-align:right;font-weight:800;font-variant-numeric:tabular-nums'>{nt(a['value'])}</div>", unsafe_allow_html=True)
                    c[5].markdown(f"<div class='po-td'>{ADOT.get(a['status'],'')} {a['status']}</div>", unsafe_allow_html=True)
                    if sel_now:
                        with st.container(border=True):
                            _asset_detail(a, assets)

    with tab_board:
        st.markdown(
            f"<div style='display:flex;gap:16px;align-items:center;flex-wrap:wrap;background:#fff;border:1px solid {LINE};"
            f"border-radius:12px;padding:10px 14px;font-size:13px;font-weight:600;margin-bottom:12px'>"
            f"<span style='color:{SUB};font-weight:500'>顯示各單位使用中的資產（改派單位請至「資產清單」）：</span>"
            f"<span style='display:inline-flex;align-items:center;gap:6px'><span style='width:12px;height:12px;border-radius:3px;background:{INDIGO};display:inline-block'></span>列帳資產 B26</span>"
            f"<span style='display:inline-flex;align-items:center;gap:6px'><span style='width:12px;height:12px;border-radius:3px;background:{JADE};display:inline-block'></span>列管資產 A26</span>"
            f"</div>",
            unsafe_allow_html=True)
        fixed = assets[(assets["asset_type"].isin(ASSET_RECORD_CLASSES)) & (~assets["status"].isin(["已報廢", "已作廢"]))]
        for u in UNITS:
            sub = fixed[fixed["unit"] == u]
            ledger = sub[sub["asset_type"] == "列帳資產"]
            managed = sub[sub["asset_type"] == "列管資產"]
            title = (f"🏢 {u}　·　{len(sub)} 件　·　"
                     f"列帳 {nt(ledger['value'].sum())}　列管 {nt(managed['value'].sum())}")
            with st.expander(title, expanded=False):
                if sub.empty:
                    st.caption("此單位目前沒有資產。")
                else:
                    rowhtml = ""
                    for a in sub.sort_values("asset_type").to_dict("records"):
                        is_ledger = a["asset_type"] == "列帳資產"
                        accent = INDIGO if is_ledger else JADE
                        soft = INDIGO_SOFT if is_ledger else JADE_SOFT
                        rowhtml += (f"<tr><td style='border-left:4px solid {accent};padding-left:10px;font-weight:600'>{a['name']}</td>"
                                    f"<td><span style='background:{soft};color:{accent};font-size:11px;font-weight:700;padding:1px 8px;border-radius:8px;white-space:nowrap'>{a['asset_type']}</span></td>"
                                    f"<td style='color:{SUB};font-variant-numeric:tabular-nums'>{a['id']}</td>"
                                    f"<td style='text-align:right;font-weight:700;font-variant-numeric:tabular-nums'>{nt(a['value'])}</td></tr>")
                    st.markdown(
                        "<table class='t'><tr><th>品名</th><th>分類</th><th>財產編號</th>"
                        f"<th style='text-align:right'>價值</th></tr>{rowhtml}</table>",
                        unsafe_allow_html=True)

    with tab_add:
        if not is_admin():
            st.info("僅管理者可手動新增資產。")
        else:
            _asset_manual_add(assets)


def _asset_manual_add(assets):
    if True:
        st.markdown(f"<div style='color:{SUB};font-size:14px;margin-bottom:6px'>登錄以前就買好、不是透過本系統採購的舊資產。來源會標記為「手動新增」。耗材不在此登錄。</div>", unsafe_allow_html=True)
        a1, a2 = st.columns(2)
        m_name = a1.text_input("品名", key="m_name", placeholder="例如：筆記型電腦")
        m_cat = a2.selectbox("類別", CATEGORIES, key="m_cat")
        a3, a4 = st.columns(2)
        m_value = a3.number_input("單筆價值（NT$）", min_value=0, step=100, key="m_value")
        m_qty = a4.number_input("數量", min_value=1, step=1, value=1, key="m_qty",
                                help="一次新增多筆相同資產（會各給一個獨立財產編號）")
        a5, a6 = st.columns(2)
        default_class = "列帳資產" if m_value >= LEDGER_THRESHOLD else "列管資產"
        if m_value >= LEDGER_THRESHOLD:
            st.info(f"💡 金額已達 {nt(LEDGER_THRESHOLD)}，建議分類為「列帳資產」（編號 B26）。已自動帶入，如需調整可改選下方分類。")
        m_type = a5.radio("資產分類", ASSET_RECORD_CLASSES, index=ASSET_RECORD_CLASSES.index(default_class),
                          horizontal=True, key="m_type",
                          help=f"≥ {nt(LEDGER_THRESHOLD)} 建議列帳資產（B26）；其餘列管資產（A26）")
        m_status = a6.selectbox("狀態", ASSET_STATUS_OPTS, key="m_status")
        a7, a8 = st.columns(2)
        m_unit = a7.selectbox("使用單位", UNITS, key="m_unit")
        m_acq = a8.date_input("取得日期", key="m_acq")
        if st.button("新增資產", type="primary", disabled=not m_name.strip()):
            rows, used = [], set()
            for _ in range(int(m_qty)):
                aid, _p = new_asset_id(assets, m_type, used)
                used.add(aid)
                rows.append([aid, m_name.strip(), m_cat, m_value, "手動新增",
                             m_acq.isoformat(), m_type, m_unit, m_status])
            save("assets", pd.concat([assets, pd.DataFrame(rows, columns=SCHEMAS["assets"])], ignore_index=True))
            flash(f"已新增 {len(rows)} 筆「{m_name.strip()}」{m_type}")
            st.rerun()


# ============================ 供應商 ============================
def page_suppliers(data):
    sups, pos, items = data["suppliers"], data["purchase_orders"], data["po_items"]
    st.markdown("<h1>供應商</h1><p style='color:#5A6472;margin-top:-8px'>採購與資產共用同一份供應商名單；可點選查看各供應商的採購單</p>", unsafe_allow_html=True)
    tab_list, tab_new = st.tabs(["清單", "＋ 新增供應商"])

    with tab_list:
        sq = st.text_input("供應商搜尋", placeholder="輸入公司名稱或統一編號查詢…", label_visibility="collapsed")
        shown = sups
        if sq.strip():
            kw = sq.strip()
            shown = sups[sups["name"].astype(str).str.contains(kw, case=False, na=False)
                         | sups["tax_id"].astype(str).str.contains(kw, case=False, na=False)]
        st.caption(f"{len(shown)} 家供應商")
        cards = ""
        for s in shown.itertuples():
            n = int((pos["supplier_id"] == s.id).sum())
            note = (getattr(s, "note", "") or "").strip()
            note_html = (f"<div style='border-top:1px solid {LINE_SOFT};margin-top:10px;padding-top:10px;"
                         f"color:{SUB};font-size:13px'>📝 {note}</div>") if note else ""
            cards += (f"<div class='card'><div style='display:flex;justify-content:space-between;align-items:flex-start'>"
                      f"<div><div style='font-weight:800;font-size:16px'>{s.name}</div>"
                      f"<div style='color:{SUB};font-size:12px;margin-top:2px'>統編 {s.tax_id}</div></div>"
                      f"<span class='pill' style='background:{JADE_SOFT};color:{JADE}'>{n} 張採購單</span></div>"
                      f"<div style='border-top:1px solid {LINE_SOFT};margin-top:12px;padding-top:12px;color:{SUB};font-size:14px;display:flex;gap:16px'>"
                      f"<span>{s.contact}</span><span>{s.phone}</span></div>{note_html}</div>")
        if cards:
            st.markdown(f"<div class='grid2'>{cards}</div>", unsafe_allow_html=True)
        else:
            st.info("找不到符合的供應商")

        # ---- 點選供應商，查看其採購單 ----
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        if len(sups):
            name_to_id = {s.name: s.id for s in sups.itertuples()}
            chosen = st.selectbox("查看供應商的採購單", list(name_to_id.keys()))
            sid = name_to_id[chosen]
            their = pos[pos["supplier_id"] == sid]
            with st.container(border=True):
                total_amt = sum(po_total(items, p) for p in their["id"])
                st.markdown(
                    f"<div style='display:flex;justify-content:space-between;align-items:center'>"
                    f"<h3 style='margin:0'>{chosen}</h3>"
                    f"<span style='color:{SUB};font-size:14px'>共 {len(their)} 張 ・ 累計 "
                    f"<b style='color:{AMBER};font-variant-numeric:tabular-nums'>{nt(total_amt)}</b></span></div>",
                    unsafe_allow_html=True)
                if their.empty:
                    st.caption("此供應商目前沒有採購單")
                else:
                    body = ""
                    for r in their.itertuples():
                        purpose = getattr(r, "purpose", "") or "—"
                        body += (f"<tr><td style='font-weight:700'>{r.id}</td><td style='color:{SUB}'>{r.date}</td>"
                                 f"<td style='color:{SUB}'>{purpose}</td>"
                                 f"<td style='text-align:right;font-weight:800;font-variant-numeric:tabular-nums'>{nt(po_total(items, r.id))}</td>"
                                 f"<td style='text-align:center'>{pill(r.status, STATUS_STYLE)}</td></tr>")
                    st.markdown(
                        "<table class='t'><tr><th>單號</th><th>日期</th><th>採購用途</th>"
                        f"<th style='text-align:right'>金額</th><th style='text-align:center'>狀態</th></tr>{body}</table>",
                        unsafe_allow_html=True)

    with tab_new:
        name = st.text_input("供應商名稱")
        tax = st.text_input("統一編號（8 碼）", max_chars=8)
        c1, c2 = st.columns(2)
        contact = c1.text_input("聯絡人")
        phone = c2.text_input("電話")
        note = st.text_area("備註", placeholder="例如：可月結、報價含安裝、配合多年…", height=80)
        tax_ok = tax.isdigit() and len(tax) == 8
        dup = tax in sups["tax_id"].astype(str).tolist()
        if tax and not tax_ok:
            st.warning("統編需為 8 位數字")
        if dup:
            st.warning("此統編已存在")
        if st.button("新增", type="primary", disabled=not (name.strip() and tax_ok and not dup)):
            new = pd.DataFrame([[f"S{len(sups) + 1}", name.strip(), tax, contact, phone, note.strip()]], columns=SCHEMAS["suppliers"])
            save("suppliers", pd.concat([sups, new], ignore_index=True))
            flash(f"已新增供應商「{name}」")
            st.rerun()


# ============================ 作廢申請（審核） ============================
def page_void(data):
    vr = data["void_requests"]
    pos = data["purchase_orders"]
    items = data["po_items"]
    assets = data["assets"]
    sups = data["suppliers"]

    st.markdown("<h1>作廢申請</h1><p style='color:#5A6472;margin-top:-8px'>前台送出的作廢申請會列在這裡，由你核准或駁回；核准後才會真正作廢（資料保留、僅標記）</p>", unsafe_allow_html=True)

    pending = vr[vr["status"] == "待審核"]
    done = vr[vr["status"] != "待審核"]

    st.markdown(f"<div class='sect'>待審核（{len(pending)}）</div>", unsafe_allow_html=True)
    if pending.empty:
        st.info("目前沒有待審核的作廢申請。")
    for r in pending.itertuples():
        with st.container(border=True):
            # 明細
            if r.target_type == "採購單":
                po = pos[pos["id"] == r.target_id]
                if not po.empty:
                    po = po.iloc[0]
                    detail = (f"供應商 {sup_name(sups, po['supplier_id'])} ・ 採購日期 {po['date']}"
                              f" ・ 採購人員 {po.get('buyer') or '—'} ・ 合計 {nt(po_total(items, r.target_id))}"
                              f"<br>用途：{po.get('purpose') or '—'}")
                else:
                    detail = "<span style='color:#B42318'>（找不到此採購單，可能已被移除）</span>"
            else:
                a = assets[assets["id"] == r.target_id]
                if not a.empty:
                    a = a.iloc[0]
                    detail = (f"{a['name']} ・ {a['category']} ・ {a['asset_type']}"
                              f" ・ {nt(a['value'])} ・ {a['unit']} ・ 狀態 {a['status']}")
                else:
                    detail = "<span style='color:#B42318'>（找不到此資產，可能已被移除）</span>"
            st.markdown(
                f"<div style='display:flex;justify-content:space-between;align-items:flex-start'>"
                f"<div><b>{r.target_type}　{r.target_id}</b>"
                f"<div style='color:{SUB};font-size:13px;margin-top:4px'>{detail}</div></div>"
                f"<span style='color:{FAINT};font-size:12px;white-space:nowrap'>{r.requested_at}</span></div>"
                f"<div style='margin-top:8px;background:{AMBER_SOFT};color:{AMBER};border-radius:8px;padding:6px 10px;font-size:13px'>"
                f"📝 作廢原因：{r.reason}</div>",
                unsafe_allow_html=True)
            b1, b2 = st.columns(2)
            if b1.button("✅ 核准作廢", key=f"ok_{r.req_id}", type="primary", use_container_width=True):
                _apply_void(data, r.target_type, r.target_id)
                _set_void_status(vr, r.req_id, "已核准")
                flash(f"已核准作廢：{r.target_type} {r.target_id}")
                st.rerun()
            if b2.button("✕ 駁回", key=f"no_{r.req_id}", use_container_width=True):
                _set_void_status(vr, r.req_id, "已駁回")
                flash(f"已駁回作廢申請：{r.target_type} {r.target_id}")
                st.rerun()

    if not done.empty:
        st.markdown(f"<div class='sect' style='margin-top:18px'>已處理紀錄</div>", unsafe_allow_html=True)
        rowhtml = ""
        for r in done.sort_values("requested_at", ascending=False).itertuples():
            rowhtml += (f"<tr><td>{r.requested_at}</td><td>{r.target_type}</td><td style='font-weight:600'>{r.target_id}</td>"
                        f"<td style='color:{SUB}'>{r.reason}</td><td style='text-align:center'>{pill(r.status, STATUS_STYLE if r.status in STATUS_STYLE else ASTATUS_STYLE)}</td></tr>")
        st.markdown(f"<div class='card' style='padding:4px'><table class='t'>"
                    f"<tr><th>時間</th><th>類型</th><th>編號</th><th>原因</th><th style='text-align:center'>結果</th></tr>{rowhtml}</table></div>",
                    unsafe_allow_html=True)


def _set_void_status(vr, req_id, status):
    upd = vr.set_index("req_id")
    upd.loc[req_id, "status"] = status
    save("void_requests", upd.reset_index()[SCHEMAS["void_requests"]])


def _apply_void(data, target_type, target_id):
    if target_type == "採購單":
        pos = data["purchase_orders"]
        save("purchase_orders", pos.assign(status=pos["status"].where(pos["id"] != target_id, "已作廢")))
    else:
        assets = data["assets"]
        upd = assets.set_index("id")
        if target_id in upd.index:
            upd.loc[target_id, "status"] = "已作廢"
            save("assets", upd.reset_index()[SCHEMAS["assets"]])


# ============================ 設定（單位管理） ============================
def page_settings(data):
    assets = data["assets"]
    st.markdown("<h1>設定</h1><p style='color:#5A6472;margin-top:-8px'>管理使用單位；變更後採購、資產、單位配置等頁面會即時套用，不用改程式</p>", unsafe_allow_html=True)
    st.markdown("<div class='sect'>使用單位</div>", unsafe_allow_html=True)

    # 新增單位
    c1, c2 = st.columns([4, 1])
    new = c1.text_input("新增單位名稱", placeholder="例如：青年牧區", label_visibility="collapsed")
    if c2.button("新增", type="primary", use_container_width=True, disabled=not new.strip()):
        nm = new.strip()
        if nm in UNITS:
            flash(f"單位「{nm}」已存在")
        else:
            save("units", pd.DataFrame({"name": UNITS + [nm]}))
            flash(f"已新增單位「{nm}」")
        st.rerun()

    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

    # 既有單位清單 + 刪除
    for u in UNITS:
        cnt = int((assets["unit"] == u).sum())
        col1, col2 = st.columns([4, 1])
        col1.markdown(
            f"<div class='card' style='padding:12px 16px;display:flex;justify-content:space-between;align-items:center'>"
            f"<span style='font-weight:700'>🏢 {u}</span>"
            f"<span style='color:{SUB};font-size:13px'>{cnt} 項資產使用中</span></div>",
            unsafe_allow_html=True)
        disabled = (cnt > 0) or (len(UNITS) <= 1)
        if col2.button("刪除", key=f"del_unit_{u}", use_container_width=True, disabled=disabled):
            save("units", pd.DataFrame({"name": [x for x in UNITS if x != u]}))
            flash(f"已刪除單位「{u}」")
            st.rerun()
        if cnt > 0:
            col2.caption("有資產使用中")

    st.caption("提示：若要刪除某單位，請先到「資產」把該單位的資產改派到其他單位，數量歸零後才能刪除。")


# ============================ 主程式 ============================
def main():
    inject_css()
    if st.session_state.get("_flash"):
        st.toast(st.session_state.pop("_flash"), icon="✅")
    if not require_login():
        return

    if "page" not in st.session_state:
        st.session_state.page = "儀表板"

    with st.sidebar:
        role_label = "管理者" if is_admin() else "一般使用者"
        role_color = JADE if is_admin() else SUB
        st.markdown("<div style='display:flex;align-items:center;gap:9px;padding:6px 2px 4px'>"
                    f"<div style='width:34px;height:34px;border-radius:10px;background:{JADE};display:flex;"
                    "align-items:center;justify-content:center;font-size:18px'>📦</div>"
                    "<div><div class='brand-title'>行政後勤</div>"
                    "<div style='font-size:12px'>採購・資產整合平台</div></div></div>", unsafe_allow_html=True)
        st.markdown(f"<div style='font-size:12px;color:{role_color};font-weight:700;padding:0 2px 12px'>● {role_label}</div>",
                    unsafe_allow_html=True)

        nav_items = [("儀表板", "📊"), ("採購", "🛒"), ("資產", "📦"), ("供應商", "👥")]
        if is_admin():
            nav_items += [("作廢申請", "🗑️"), ("設定", "⚙️")]
        for label, icon in nav_items:
            active = st.session_state.page == label
            if st.button(f"{icon}\u2002{label}", key=f"nav_{label}", use_container_width=True,
                         type="primary" if active else "secondary"):
                st.session_state.page = label
                st.rerun()

        st.divider()
        if st.button("登出", key="logout", use_container_width=True):
            st.session_state.authed = False
            st.session_state.role = None
            st.rerun()
        if is_admin():
            with st.expander("⚙️ 首次設定 / 重建資料（危險）"):
                st.caption("⚠️ 僅限第一次建置使用。按下會「清空並覆蓋」現有所有資料，改回範例資料，無法復原。")
                confirm = st.checkbox("我了解這會清除現有資料，仍要重建", key="init_confirm")
                if st.button("初始化試算表（含範例）", disabled=not confirm):
                    init_sheets()
                    st.session_state.init_confirm = False
                    flash("已建立／重建工作表與範例資料")
                    st.rerun()

    try:
        data = load_all()
    except Exception as e:
        st.error("無法連線 Google Sheet，請確認 secrets 設定與試算表分享權限。")
        st.exception(e)
        return

    # 從 units 分頁載入單位清單（覆蓋預設）；沒有資料就用預設
    global UNITS
    unit_list = [u for u in data["units"]["name"].astype(str).tolist() if u.strip()]
    UNITS = unit_list if unit_list else list(DEFAULT_UNITS)

    if all(data[n].empty for n in SCHEMAS):
        st.info("試算表是空的。請開啟左側「⚙️ 首次設定」並按「初始化試算表」建立範例資料。")

    # 權限保護：非管理者不得進入管理頁（即使手動切換）
    admin_only = {"作廢申請", "設定"}
    if st.session_state.page in admin_only and not is_admin():
        st.session_state.page = "儀表板"

    {"儀表板": page_dashboard, "採購": page_procurement,
     "資產": page_assets, "供應商": page_suppliers,
     "作廢申請": page_void, "設定": page_settings}[st.session_state.page](data)


if __name__ == "__main__":
    main()
