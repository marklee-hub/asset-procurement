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
from datetime import date

# ============================ 設定 ============================
st.set_page_config(page_title="採購・資產整合平台", page_icon="📦", layout="wide")

# ---- 配色（對齊 React 版 tokens）----
INK, SUB, FAINT = "#161A22", "#5A6472", "#8A93A2"
LINE, LINE_SOFT = "#E4E8EE", "#EEF1F5"
JADE, JADE_SOFT = "#0F766E", "#E2F1EF"
AMBER, AMBER_SOFT = "#B45309", "#FBEEDC"
INDIGO, INDIGO_SOFT = "#4338CA", "#E7E7FA"
DANGER = "#B42318"

CATEGORIES = ["3C設備", "辦公家具", "音響設備", "文宣耗材"]
DEFAULT_UNITS = ["傳道部", "行銷部", "影音部", "行政部", "神學院", "財務部"]
UNITS = list(DEFAULT_UNITS)   # 啟動時會從 Google Sheet 的 units 分頁覆蓋
FIXED_THRESHOLD = 10000
TRACKED = [c for c in CATEGORIES if c != "文宣耗材"]

STATUS_OPTS = ["草稿", "待驗收", "已驗收"]
ASSET_STATUS_OPTS = ["使用中", "維修中", "已報廢"]
ASSET_TYPE_OPTS = ["固定資產", "一般資產"]

STATUS_STYLE = {"草稿": (LINE_SOFT, SUB), "待驗收": (AMBER_SOFT, AMBER), "已驗收": (JADE_SOFT, JADE)}
ASTATUS_STYLE = {"使用中": (JADE_SOFT, JADE), "維修中": (AMBER_SOFT, AMBER), "已報廢": ("#F3EDED", DANGER)}
ATYPE_STYLE = {"固定資產": (INDIGO_SOFT, INDIGO), "一般資產": (LINE_SOFT, SUB)}

SCHEMAS = {
    "suppliers":       ["id", "name", "tax_id", "contact", "phone"],
    "purchase_orders": ["id", "supplier_id", "date", "status", "purpose", "delivery_date", "note"],
    "po_items":        ["po_id", "name", "category", "qty", "price"],
    "assets":          ["id", "name", "category", "value", "source_po",
                        "acquired", "asset_type", "unit", "status"],
    "units":           ["name"],
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
    .stApp {{ background: #F4F6F8; }}
    #MainMenu, footer {{ visibility: hidden; }}
    [data-testid="stHeader"] {{ background: transparent; }}
    .block-container {{ padding-top: 2.2rem; padding-bottom: 3rem; max-width: 1120px; }}

    /* 側欄深色 */
    section[data-testid="stSidebar"] {{ background: #15201F; }}
    section[data-testid="stSidebar"] * {{ color: #C9D6D2; }}
    section[data-testid="stSidebar"] .brand-title {{ color:#fff; font-weight:800; font-size:1rem; }}
    /* 側欄導覽按鈕 */
    section[data-testid="stSidebar"] .stButton > button {{
        text-align:left; justify-content:flex-start; border:none; background:transparent;
        color:#C9D6D2; font-weight:600; border-radius:12px; padding:9px 14px; font-size:15px;
    }}
    section[data-testid="stSidebar"] .stButton > button:hover {{ background:rgba(255,255,255,.08); color:#fff; }}
    section[data-testid="stSidebar"] .stButton > button[kind="primary"] {{ background:#0F766E; color:#fff; }}
    section[data-testid="stSidebar"] .stButton > button[kind="primary"]:hover {{ background:#0c655e; }}

    /* 主按鈕 */
    .stButton > button {{ border-radius:12px; font-weight:600; border:1px solid {LINE}; }}
    .stButton > button[kind="primary"] {{ background:{JADE}; border-color:{JADE}; }}
    .stButton > button[kind="primary"]:hover {{ background:#0c655e; border-color:#0c655e; }}

    /* 側欄導覽：乾淨無框、左對齊、選中highlight */
    section[data-testid="stSidebar"] .stButton > button {{
        border:none !important; background:transparent !important; color:#C9D6D2 !important;
        justify-content:flex-start !important; text-align:left !important;
        font-weight:600; border-radius:12px; padding:9px 14px; box-shadow:none !important;
    }}
    section[data-testid="stSidebar"] .stButton > button p {{ text-align:left; width:100%; }}
    section[data-testid="stSidebar"] .stButton > button:hover {{
        background:rgba(255,255,255,.08) !important; color:#fff !important;
    }}
    section[data-testid="stSidebar"] .stButton > button[kind="primary"] {{
        background:{JADE} !important; color:#fff !important;
    }}
    section[data-testid="stSidebar"] .stButton > button[kind="primary"]:hover {{
        background:#0c655e !important;
    }}
    /* 登出鈕：細邊框與選單區隔 */
    section[data-testid="stSidebar"] [data-testid="stButton"]:has(#logout) button,
    section[data-testid="stSidebar"] .stButton:last-of-type > button {{
        color:#8A93A2 !important;
    }}

    /* 輸入元件圓角 */
    [data-baseweb="select"] > div, .stTextInput input, .stNumberInput input {{
        border-radius:12px !important;
    }}

    /* 標題 */
    h1 {{ font-weight:800 !important; letter-spacing:-.01em; }}
    h2, h3 {{ font-weight:700 !important; }}

    /* 卡片系統 */
    .card {{ background:#fff; border:1px solid {LINE}; border-radius:18px; padding:18px; }}
    .grid4 {{ display:grid; grid-template-columns:repeat(4,1fr); gap:14px; }}
    .grid2 {{ display:grid; grid-template-columns:repeat(2,1fr); gap:14px; }}
    @media (max-width:760px){{ .grid4{{grid-template-columns:repeat(2,1fr)}} .grid2{{grid-template-columns:1fr}} }}
    .stat-label {{ color:{FAINT}; font-size:12px; font-weight:700; letter-spacing:.04em; text-transform:uppercase; }}
    .stat-value {{ font-size:26px; font-weight:800; line-height:1.1; margin-top:4px; font-variant-numeric:tabular-nums; }}
    .stat-sub {{ color:{SUB}; font-size:12px; margin-top:2px; }}
    .pill {{ display:inline-flex; align-items:center; padding:2px 9px; border-radius:999px; font-size:12px; font-weight:700; white-space:nowrap; }}
    .flow {{ display:flex; align-items:center; gap:10px; flex-wrap:wrap; }}
    .flowcard {{ border-radius:14px; padding:12px 16px; flex:1; min-width:150px; }}
    .flowcard .fl {{ font-size:12px; font-weight:700; }}
    .flowcard .fv {{ font-weight:800; color:{INK}; font-variant-numeric:tabular-nums; }}
    .sect {{ color:{FAINT}; font-size:12px; font-weight:700; letter-spacing:.04em; text-transform:uppercase; margin-bottom:10px; }}
    table.t {{ width:100%; border-collapse:collapse; font-size:14px; }}
    table.t th {{ text-align:left; color:{FAINT}; font-size:11px; text-transform:uppercase; font-weight:700; padding:10px 14px; border-bottom:1px solid {LINE}; }}
    table.t td {{ padding:11px 14px; border-bottom:1px solid {LINE_SOFT}; }}
    .acard {{ background:#fff; border:1px solid {LINE}; border-radius:16px; padding:14px; margin-bottom:6px; }}
    .acard .aid {{ color:{FAINT}; font-size:12px; font-weight:700; font-variant-numeric:tabular-nums; }}
    .acard .anm {{ font-weight:800; }}
    .acard .acat {{ color:{SUB}; font-size:12px; }}
    .acard .aval {{ font-weight:800; color:{INK}; font-variant-numeric:tabular-nums; }}
    .unit-head {{ font-weight:800; display:flex; align-items:center; justify-content:space-between; margin-bottom:2px; }}
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
    _, mid, _ = st.columns([1, 1.4, 1])
    with mid:
        st.markdown("<div style='height:8vh'></div>", unsafe_allow_html=True)
        st.markdown(
            f"<div class='card' style='text-align:center;padding:32px'>"
            f"<div style='font-size:34px'>📦</div>"
            f"<h2 style='margin:.3rem 0'>採購・資產整合平台</h2>"
            f"<p style='color:{SUB};margin-top:0'>請輸入共用密碼</p></div>",
            unsafe_allow_html=True)
        pw = st.text_input("密碼", type="password", label_visibility="collapsed", placeholder="密碼")
        if st.button("登入", type="primary", use_container_width=True):
            if pw == st.secrets.get("app_password", ""):
                st.session_state.authed = True
                st.rerun()
            else:
                st.error("密碼錯誤")
    return False


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
        ["S1", "全美電腦資訊", "12345675", "王經理", "02-2222-3333"],
        ["S2", "永興辦公家具", "23456781", "林小姐", "02-2555-6666"],
        ["S3", "聲學音響工程", "34567892", "陳先生", "03-3333-4444"],
        ["S4", "印刷大師文宣", "45678903", "李主任", "02-2888-9999"],
    ], columns=SCHEMAS["suppliers"])


def _seed_pos():
    return pd.DataFrame([
        ["PO-2026-001", "S1", "2026-04-12", "已驗收", "辦公室設備汰換", "2026-04-18", ""],
        ["PO-2026-002", "S2", "2026-05-03", "已驗收", "主日場地桌椅", "2026-05-09", ""],
        ["PO-2026-003", "S3", "2026-06-01", "待驗收", "敬拜團音響升級", "2026-06-15", "含現場安裝"],
        ["PO-2026-004", "S4", "2026-06-06", "草稿", "主日文宣印製", "2026-06-12", ""],
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
        ["PO-2026-004", "主日文宣海報", "文宣耗材", 500, 12],
    ], columns=SCHEMAS["po_items"])


def _seed_assets():
    rows, n = [], 1

    def mk(name, cat, val, po, d, atype, unit, status="使用中"):
        nonlocal n
        rows.append([f"A26-{n:04d}", name, cat, val, po, d, atype, unit, status])
        n += 1

    mk("筆記型電腦", "3C設備", 28000, "PO-2026-001", "2026-04-18", "固定資產", "行政辦公室")
    mk("筆記型電腦", "3C設備", 28000, "PO-2026-001", "2026-04-18", "固定資產", "媒體組", "維修中")
    mk("短焦投影機", "3C設備", 24000, "PO-2026-001", "2026-04-18", "固定資產", "主堂")
    mk("短焦投影機", "3C設備", 24000, "PO-2026-001", "2026-04-18", "固定資產", "兒童主日學")
    mk("雷射印表機", "3C設備", 15000, "PO-2026-001", "2026-04-18", "固定資產", "行政辦公室")
    for i in range(4):
        mk("折疊長桌", "辦公家具", 1800, "PO-2026-002", "2026-05-09", "一般資產", UNITS[i % len(UNITS)])
    for i in range(8):
        mk("堆疊摺疊椅", "辦公家具", 500, "PO-2026-002", "2026-05-09", "一般資產", UNITS[i % len(UNITS)],
           "已報廢" if i == 7 else "使用中")
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


def next_asset_seq(assets):
    if assets.empty:
        return 1
    nums = assets["id"].astype(str).str.extract(r"A26-(\d+)")[0].dropna().astype(int)
    return (nums.max() + 1) if len(nums) else 1


def po_total(items, po_id):
    rows = items[items["po_id"] == po_id]
    return float((rows["qty"] * rows["price"]).sum())


def sup_name(suppliers, sid):
    m = suppliers[suppliers["id"] == sid]
    return m.iloc[0]["name"] if len(m) else "—"


# ============================ 儀表板 ============================
def page_dashboard(data):
    pos, items, assets, sups = data["purchase_orders"], data["po_items"], data["assets"], data["suppliers"]
    received = pos[pos["status"] == "已驗收"]["id"].tolist()
    pending = pos[pos["status"] == "待驗收"]["id"].tolist()
    spend = sum(po_total(items, p) for p in received)
    pending_amt = sum(po_total(items, p) for p in pending)
    live = assets[assets["status"] != "已報廢"]
    fixed_val = live[live["asset_type"] == "固定資產"]["value"].sum()
    gen_val = live[live["asset_type"] == "一般資產"]["value"].sum()

    st.markdown("<h1>儀表板</h1><p style='color:#5A6472;margin-top:-8px'>採購支出與資產價值一覽</p>", unsafe_allow_html=True)

    def stat(label, value, accent, sub=""):
        return (f"<div class='card'><div class='stat-label'>{label}</div>"
                f"<div class='stat-value' style='color:{accent}'>{value}</div>"
                f"<div class='stat-sub'>{sub}</div></div>")

    st.markdown(
        "<div class='grid4'>"
        + stat("本年採購支出", nt(spend), AMBER, "已驗收採購單")
        + stat("固定資產總值", nt(fixed_val), INDIGO, "可調整使用單位")
        + stat("一般資產總值", nt(gen_val), JADE, "排除已報廢")
        + stat("待驗收金額", nt(pending_amt), f"{len(pending)} 張單 ・ {len(sups)} 家供應商")
        + "</div>", unsafe_allow_html=True)

    # 轉換流向
    def fc(bg, color, label, value):
        return (f"<div class='flowcard' style='background:{bg}'>"
                f"<div class='fl' style='color:{color}'>{label}</div>"
                f"<div class='fv'>{value}</div></div>")
    arrow = "<div style='color:#8A93A2;font-size:20px'>→</div>"
    st.markdown(
        "<div class='card' style='margin-top:14px'><div class='sect'>採購到資產的轉換</div><div class='flow'>"
        + fc(AMBER_SOFT, AMBER, "採購支出", nt(spend)) + arrow
        + fc(LINE_SOFT, SUB, "驗收分類", f"{len(received)} 張單") + arrow
        + fc(INDIGO_SOFT, INDIGO, "固定資產", nt(fixed_val))
        + fc(JADE_SOFT, JADE, "一般資產", nt(gen_val))
        + "</div></div>", unsafe_allow_html=True)

    left, right = st.columns([3, 2])
    with left:
        with st.container(border=True):
            st.markdown("<div class='sect'>各類別：採購支出 vs 資產價值</div>", unsafe_allow_html=True)
            rows = []
            for cat in TRACKED:
                sp = sum((items[(items["po_id"] == p) & (items["category"] == cat)]["qty"]
                          * items[(items["po_id"] == p) & (items["category"] == cat)]["price"]).sum() for p in received)
                rows.append({"類別": cat, "採購支出": sp, "資產價值": live[live["category"] == cat]["value"].sum()})
            st.bar_chart(pd.DataFrame(rows).set_index("類別"), color=[AMBER, JADE], height=240)
    with right:
        with st.container(border=True):
            st.markdown("<div class='sect'>待驗收</div>", unsafe_allow_html=True)
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
def page_procurement(data):
    pos, items, assets, sups = data["purchase_orders"], data["po_items"], data["assets"], data["suppliers"]
    st.markdown("<h1>採購</h1><p style='color:#5A6472;margin-top:-8px'>採購單管理；驗收時分類為固定／一般資產</p>", unsafe_allow_html=True)
    tab_list, tab_new = st.tabs(["採購單清單", "＋ 新增採購單"])

    with tab_list:
        head = "<tr><th>單號</th><th>供應商</th><th>日期</th><th style='text-align:right'>金額</th><th style='text-align:center'>狀態</th></tr>"
        body = ""
        for r in pos.itertuples():
            body += (f"<tr><td style='font-weight:700'>{r.id}</td><td style='color:{SUB}'>{sup_name(sups, r.supplier_id)}</td>"
                     f"<td style='color:{SUB}'>{r.date}</td>"
                     f"<td style='text-align:right;font-weight:800;font-variant-numeric:tabular-nums'>{nt(po_total(items, r.id))}</td>"
                     f"<td style='text-align:center'>{pill(r.status, STATUS_STYLE)}</td></tr>")
        st.markdown(f"<div class='card' style='padding:4px'><table class='t'>{head}{body}</table></div>", unsafe_allow_html=True)

        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        if len(pos):
            sel = st.selectbox("查看採購單明細", pos["id"].tolist())
            po = pos[pos["id"] == sel].iloc[0]
            its = items[items["po_id"] == sel].copy()
            with st.container(border=True):
                meta = f"{sup_name(sups, po['supplier_id'])} ・ 建立日 {po['date']}"
                if po.get("delivery_date"):
                    meta += f" ・ 交貨日 {po['delivery_date']}"
                purpose_html = f"<div style='margin:.2rem 0'><b>採購用途：</b>{po.get('purpose') or '—'}</div>"
                note_html = f"<div style='color:{SUB}'><b>備註：</b>{po.get('note')}</div>" if po.get("note") else ""
                st.markdown(f"<h3 style='margin:0'>{sel} &nbsp; {pill(po['status'], STATUS_STYLE)}</h3>"
                            f"<p style='color:{SUB};margin:.2rem 0 .3rem'>{meta}</p>"
                            f"{purpose_html}{note_html}",
                            unsafe_allow_html=True)
                rowhtml = ""
                for it in its.itertuples():
                    tag = "" if it.category in TRACKED else f"<span style='color:{FAINT};font-size:12px'>（耗材・不列管）</span>"
                    rowhtml += (f"<tr><td style='font-weight:600'>{it.name} {tag}</td><td style='color:{SUB}'>{it.category}</td>"
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

    with tab_new:
        sup_map = {sup_name(sups, r.id): r.id for r in sups.itertuples()}
        if not sup_map:
            st.warning("尚無供應商，請先到「供應商」頁新增。")
            return
        chosen = st.selectbox("供應商", list(sup_map.keys()))
        cc1, cc2 = st.columns(2)
        purpose = cc1.text_input("採購用途（必填）", placeholder="例：敬拜團音響升級")
        delivery = cc2.date_input("交貨日期", value=None, format="YYYY-MM-DD")
        note = st.text_area("備註", placeholder="選填", height=70)
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
            _create_po(chosen, sup_map, valid, "草稿", pos, items, purpose, delivery, note)
        if c2.button("送出（待驗收）", type="primary", disabled=not ok, use_container_width=True):
            _create_po(chosen, sup_map, valid, "待驗收", pos, items, purpose, delivery, note)


def _create_po(chosen, sup_map, valid, status, pos, items, purpose="", delivery=None, note=""):
    new_id = next_po_id(pos)
    dstr = delivery.isoformat() if delivery else ""
    new_po = pd.DataFrame([[new_id, sup_map[chosen], date.today().isoformat(), status,
                            purpose.strip(), dstr, (note or "").strip()]], columns=SCHEMAS["purchase_orders"])
    new_items = pd.DataFrame({"po_id": new_id, "name": valid["品名"].values, "category": valid["類別"].values,
                              "qty": valid["數量"].values, "price": valid["單價"].values})[SCHEMAS["po_items"]]
    save("purchase_orders", pd.concat([pos, new_po], ignore_index=True))
    save("po_items", pd.concat([items, new_items], ignore_index=True))
    flash(f"採購單 {new_id} 已建立（{status}）")
    st.rerun()


def _receive_form(po_id, its, pos, assets):
    st.markdown(f"<div class='sect' style='margin-top:14px'>驗收入庫</div>", unsafe_allow_html=True)
    st.caption(f"逐項確認分類與使用單位（單價 ≥ {nt(FIXED_THRESHOLD)} 預設為固定資產），耗材不列管。")
    tracked = its[its["category"].isin(TRACKED)].reset_index(drop=True)
    skipped = its[~its["category"].isin(TRACKED)]
    choices = {}
    for i, r in enumerate(tracked.itertuples()):
        with st.container(border=True):
            st.markdown(f"<b>{r.name}</b> <span style='color:{FAINT};font-size:12px'>×{int(r.qty)} ・ {nt(r.price)} ・ {r.category}</span>", unsafe_allow_html=True)
            c1, c2 = st.columns([1.2, 1])
            default = "固定資產" if r.price >= FIXED_THRESHOLD else "一般資產"
            atype = c1.radio("分類", ASSET_TYPE_OPTS, index=ASSET_TYPE_OPTS.index(default),
                             key=f"t_{po_id}_{i}", horizontal=True, label_visibility="collapsed")
            unit = c2.selectbox("單位", UNITS, key=f"u_{po_id}_{i}", label_visibility="collapsed")
            choices[i] = (atype, unit)
    if not skipped.empty:
        st.caption("不列管（耗材）：" + "、".join(skipped["name"].tolist()))
    if st.button("確認入庫", type="primary", key=f"recv_{po_id}", use_container_width=True):
        seq = next_asset_seq(assets)
        new_rows = []
        for i, r in enumerate(tracked.itertuples()):
            atype, unit = choices[i]
            for _ in range(int(r.qty)):
                new_rows.append([f"A26-{seq:04d}", r.name, r.category, r.price, po_id,
                                 date.today().isoformat(), atype, unit, "使用中"])
                seq += 1
        save("assets", pd.concat([assets, pd.DataFrame(new_rows, columns=SCHEMAS["assets"])], ignore_index=True))
        save("purchase_orders", pos.assign(status=pos["status"].where(pos["id"] != po_id, "已驗收")))
        f = sum(1 for r in new_rows if r[6] == "固定資產")
        flash(f"{po_id} 已驗收入庫：固定資產 {f} 筆、一般資產 {len(new_rows) - f} 筆")
        st.rerun()


# ============================ 資產 ============================
def page_assets(data):
    assets = data["assets"]
    st.markdown("<h1>資產</h1><p style='color:#5A6472;margin-top:-8px'>所有列管資產皆可追溯來源採購單</p>", unsafe_allow_html=True)
    tab_list, tab_board = st.tabs(["資產清單", "單位配置（固定資產）"])

    with tab_list:
        c1, c2, c3, c4 = st.columns(4)
        f_type = c1.selectbox("類型", ["全部"] + ASSET_TYPE_OPTS)
        f_cat = c2.selectbox("類別", ["全部"] + TRACKED)
        f_unit = c3.selectbox("使用單位", ["全部"] + UNITS)
        f_status = c4.selectbox("狀態", ["全部"] + ASSET_STATUS_OPTS)
        mask = pd.Series(True, index=assets.index)
        if f_type != "全部":   mask &= assets["asset_type"] == f_type
        if f_cat != "全部":    mask &= assets["category"] == f_cat
        if f_unit != "全部":   mask &= assets["unit"] == f_unit
        if f_status != "全部": mask &= assets["status"] == f_status
        view = assets[mask]
        st.caption(f"{len(view)} 筆 ・ {view['name'].nunique()} 種品項（點品項展開看個別明細）")

        if view.empty:
            st.info("沒有符合條件的資產")
        else:
            for name, g in view.groupby("name", sort=False):
                total_val = g["value"].sum()
                atype = "固定資產" if (g["asset_type"] == "固定資產").any() else "一般資產"
                with st.expander(f"{name}　×{len(g)}　・　{atype}　・　{nt(total_val)}"):
                    by_unit = g.groupby("unit").size()
                    dist = "　".join(f"{u} {c}" for u, c in by_unit.items())
                    st.markdown(f"<div style='color:{SUB};font-size:13px;margin-bottom:8px'>單位分佈：{dist}</div>", unsafe_allow_html=True)

                    # 批次改派：把 N 件從某單位改派到另一單位
                    present = list(by_unit.index)
                    b1, b2, b3, b4 = st.columns([1.2, 1.2, 1, 0.8])
                    src = b1.selectbox("來源單位", present, key=f"src_{name}")
                    dst = b2.selectbox("改派到", UNITS, key=f"dst_{name}")
                    avail = int((g["unit"] == src).sum())
                    qty = b3.number_input("件數", min_value=1, max_value=max(1, avail), value=1, step=1, key=f"qty_{name}")
                    if b4.button("改派", key=f"mvbtn_{name}", use_container_width=True):
                        ids = g[g["unit"] == src]["id"].head(int(qty)).tolist()
                        upd = assets.set_index("id")
                        upd.loc[ids, "unit"] = dst
                        save("assets", upd.reset_index()[SCHEMAS["assets"]])
                        flash(f"已將「{name}」{len(ids)} 件從 {src} 改派到 {dst}")
                        st.rerun()

                    st.markdown(f"<div style='border-top:1px solid {LINE_SOFT};margin:10px 0 6px'></div>"
                                f"<div style='color:{FAINT};font-size:12px;font-weight:700'>個別明細</div>", unsafe_allow_html=True)

                    cap = 100
                    recs = g.to_dict("records")
                    for a in recs[:cap]:
                        d1, d2, d3 = st.columns([1.3, 1, 1])
                        d1.markdown(f"<div style='padding-top:8px;font-size:13px'><b>{a['id']}</b>　{pill(a['status'], ASTATUS_STYLE)}</div>", unsafe_allow_html=True)
                        nu = d2.selectbox("單位", UNITS, index=UNITS.index(a["unit"]) if a["unit"] in UNITS else 0,
                                          key=f"lu_{a['id']}", label_visibility="collapsed")
                        ns = d3.selectbox("狀態", ASSET_STATUS_OPTS, index=ASSET_STATUS_OPTS.index(a["status"]) if a["status"] in ASSET_STATUS_OPTS else 0,
                                          key=f"ls_{a['id']}", label_visibility="collapsed")
                        if nu != a["unit"] or ns != a["status"]:
                            upd = assets.set_index("id")
                            upd.loc[a["id"], "unit"] = nu
                            upd.loc[a["id"], "status"] = ns
                            save("assets", upd.reset_index()[SCHEMAS["assets"]])
                            flash(f"已更新資產 {a['id']}")
                            st.rerun()
                    if len(recs) > cap:
                        st.caption(f"此品項共 {len(recs)} 筆，個別明細僅顯示前 {cap} 筆；大量調整請用上方「批次改派」。")

    with tab_board:
        st.markdown(f"<div style='background:{INDIGO_SOFT};color:{INDIGO};border-radius:12px;padding:10px 14px;"
                    f"font-size:14px;font-weight:600'>用各單位欄位下方的選單即可把固定資產機動調整到別的單位。僅顯示使用中的固定資產。</div>",
                    unsafe_allow_html=True)
        fixed = assets[(assets["asset_type"] == "固定資產") & (assets["status"] != "已報廢")]
        cols = st.columns(len(UNITS))
        for i, u in enumerate(UNITS):
            with cols[i]:
                sub = fixed[fixed["unit"] == u]
                st.markdown(f"<div class='unit-head'><span>🏢 {u}</span><span style='color:{FAINT};font-size:12px'>{len(sub)} 件</span></div>"
                            f"<div style='color:{INDIGO};font-size:12px;font-weight:700;margin-bottom:6px'>{nt(sub['value'].sum())}</div>",
                            unsafe_allow_html=True)
                for a in sub.to_dict("records"):
                    st.markdown(f"<div class='acard' style='padding:10px'><div style='font-weight:700;font-size:13px'>{a['name']}</div>"
                                f"<div style='color:{SUB};font-size:11px;font-variant-numeric:tabular-nums'>{a['id']} ・ {nt(a['value'])}</div></div>",
                                unsafe_allow_html=True)
                    nu = st.selectbox("移到", UNITS, index=i, key=f"mv_{a['id']}", label_visibility="collapsed")
                    if nu != u:
                        upd = assets.set_index("id")
                        upd.loc[a["id"], "unit"] = nu
                        save("assets", upd.reset_index()[SCHEMAS["assets"]])
                        flash(f"{a['id']} 已移至 {nu}")
                        st.rerun()


# ============================ 供應商 ============================
def page_suppliers(data):
    sups, pos, items = data["suppliers"], data["purchase_orders"], data["po_items"]
    st.markdown("<h1>供應商</h1><p style='color:#5A6472;margin-top:-8px'>採購與資產共用同一份供應商名單；可點選查看各供應商的採購單</p>", unsafe_allow_html=True)
    tab_list, tab_new = st.tabs(["清單", "＋ 新增供應商"])

    with tab_list:
        cards = ""
        for s in sups.itertuples():
            n = int((pos["supplier_id"] == s.id).sum())
            cards += (f"<div class='card'><div style='display:flex;justify-content:space-between;align-items:flex-start'>"
                      f"<div><div style='font-weight:800;font-size:16px'>{s.name}</div>"
                      f"<div style='color:{SUB};font-size:12px;margin-top:2px'>統編 {s.tax_id}</div></div>"
                      f"<span class='pill' style='background:{JADE_SOFT};color:{JADE}'>{n} 張採購單</span></div>"
                      f"<div style='border-top:1px solid {LINE_SOFT};margin-top:12px;padding-top:12px;color:{SUB};font-size:14px;display:flex;gap:16px'>"
                      f"<span>{s.contact}</span><span>{s.phone}</span></div></div>")
        st.markdown(f"<div class='grid2'>{cards}</div>", unsafe_allow_html=True)

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
        tax_ok = tax.isdigit() and len(tax) == 8
        dup = tax in sups["tax_id"].astype(str).tolist()
        if tax and not tax_ok:
            st.warning("統編需為 8 位數字")
        if dup:
            st.warning("此統編已存在")
        if st.button("新增", type="primary", disabled=not (name.strip() and tax_ok and not dup)):
            new = pd.DataFrame([[f"S{len(sups) + 1}", name.strip(), tax, contact, phone]], columns=SCHEMAS["suppliers"])
            save("suppliers", pd.concat([sups, new], ignore_index=True))
            flash(f"已新增供應商「{name}」")
            st.rerun()


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
        st.markdown("<div style='display:flex;align-items:center;gap:8px;padding:6px 2px 14px'>"
                    "<div style='width:32px;height:32px;border-radius:9px;background:#0F766E;display:flex;"
                    "align-items:center;justify-content:center;font-size:18px'>📦</div>"
                    "<div><div class='brand-title'>行政後勤</div>"
                    "<div style='font-size:12px'>採購・資產整合平台</div></div></div>", unsafe_allow_html=True)

        for label, icon in [("儀表板", "📊"), ("採購", "🛒"), ("資產", "📦"), ("供應商", "👥"), ("設定", "⚙️")]:
            active = st.session_state.page == label
            if st.button(f"{icon}\u2002{label}", key=f"nav_{label}", use_container_width=True,
                         type="primary" if active else "secondary"):
                st.session_state.page = label
                st.rerun()

        st.divider()
        if st.button("登出", key="logout", use_container_width=True):
            st.session_state.authed = False
            st.rerun()
        with st.expander("⚙️ 首次設定 / 重建資料"):
            st.caption("第一次使用，或想用範例資料重建工作表時按此。會覆蓋現有四個工作表。")
            if st.button("初始化試算表（含範例）"):
                init_sheets()
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

    {"儀表板": page_dashboard, "採購": page_procurement,
     "資產": page_assets, "供應商": page_suppliers,
     "設定": page_settings}[st.session_state.page](data)


if __name__ == "__main__":
    main()
