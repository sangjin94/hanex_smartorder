# -*- coding: utf-8 -*-
"""스마트오더 라벨 생성 핵심 엔진 (채널: CU / GS / E24)

설계 원칙
 - 매핑은 '코드' 기반: 센터코드→거점센터, 상품코드→대표상품명 (E24 센터는 코드가 없어 입고센터명 키)
 - 마스터는 JSON 파일. 업로드 중 미매핑 코드는 화면에서 즉시 등록 → 재사용
 - 부착양식(A4출력)은 기존 xlsx와 픽셀 단위로 동일한 스타일로 새로 생성(정적값·정확한 라벨 수)
"""
import os, json, io
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Color
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
# 마스터 저장 경로: 배포 시 git 트리 밖(영구 디렉터리)으로 분리해 update(git pull)에도 보존.
MASTER_DIR = os.environ.get("SMARTORDER_MASTERS") or os.path.join(BASE, "masters")
SEED_DIR = os.path.join(BASE, "masters")

def ensure_masters_seeded():
    """MASTER_DIR가 비어 있으면 코드 동봉 seed(masters/)를 복사한다(최초 배포용)."""
    if os.path.abspath(MASTER_DIR) == os.path.abspath(SEED_DIR):
        return
    os.makedirs(MASTER_DIR, exist_ok=True)
    import shutil
    for fn in os.listdir(SEED_DIR):
        if not fn.endswith(".json"):
            continue
        dst = os.path.join(MASTER_DIR, fn)
        if not os.path.exists(dst):
            shutil.copy2(os.path.join(SEED_DIR, fn), dst)

# ----------------------------------------------------------------------------
# 채널 정의
# ----------------------------------------------------------------------------
CHANNELS = {
    "cu": {
        "name": "CU",
        "tag": "CU",
        # {jm} = 대월직매장 / 프리미엄직매장 (상품 구분에 따라 라벨별로 결정)
        "title_tpl": "{jm}\n(BGF리테일 CU스마트오더)",
        "center_label": "CU센터",
        "center_master": "center_cu.json",
        "center_key": "code",       # 센터코드로 매핑
        "label_product": "raw",     # 라벨 상품명 = 화주상품명(원본). 표지는 대표(한익스상품명)
        "col_width": {"A": 39.0, "B": 99.9, "C": 41.7, "D": 8.7},
        "row_h": 62.4,
        "lead_spacer": True,
        # RAW 헤더 후보 (부분일치)
        "raw": {
            "sheet_hint": ["발주확인", "스티커", "lowdata", "Sheet"],
            "header_scan_rows": 3,
            "center_code": ["센터 코드", "센터코드"],
            "center_name": ["센터명", "센터 명"],
            "store_code":  ["센터 코드", "센터코드"],
            "store_name":  ["센터명"],
            "prod_code":   ["상품 코드", "상품코드"],
            "prod_name":   ["상품명"],
            "qty":         ["총수량", "수량"],
            "date":        ["센터납품 예정일자", "예정일자", "고객인도일"],
        },
    },
    "gs": {
        "name": "GS",
        "tag": "GS",
        "title_tpl": "{jm}\n(GS슈퍼 GS스마트오더)",
        "center_label": "GS센터",
        "center_master": "center_gs.json",
        "center_key": "code",
        "label_product": "raw",
        "col_width": {"A": 39.0, "B": 99.9, "C": 41.7, "D": 8.7},
        "row_h": 62.4,
        "lead_spacer": True,
        "raw": {
            "sheet_hint": ["lowdata", "발주", "Wine", "Sheet"],
            "header_scan_rows": 2,
            "center_code": ["센터코드"],
            "center_name": ["센터"],          # exact 우선 처리
            "store_code":  ["점포코드", "최초점포코드"],
            "store_name":  ["점포명"],
            "prod_code":   ["상품코드"],
            "prod_name":   ["상품명"],
            "qty":         ["수량"],
            "date":        ["예약일자", "수령일자"],
        },
    },
    "e24": {
        "name": "이마트24",
        "tag": "E-24",
        "title_tpl": "오비맥주 {jm}\n(E-24 스마트오더)",
        "center_label": "E-24센터",
        "center_master": "center_e24.json",
        "center_key": "name",       # E24는 입고센터명으로 매핑(코드 없음)
        "label_product": "raw",     # 라벨 상품명 = 화주상품명(원본). 표지는 대표(한익스상품명)
        "col_width": {"A": 47.8, "B": 110.7, "C": 41.7, "D": 8.7},
        "row_h": 83.4,
        "lead_spacer": False,
        "raw": {
            "sheet_hint": ["lowdata", "Sheet"],
            "header_scan_rows": 2,
            "center_code": ["입고센터명"],
            "center_name": ["입고센터명"],
            "store_code":  ["점포코드"],
            "store_name":  ["점포명"],
            "prod_code":   ["상품코드"],
            "prod_name":   ["앱 상품명", "상품명(대표)", "앱상품명", "상품명"],
            "qty":         ["주문수량", "수량"],
            "date":        ["납품예정일", "픽업일자", "발주일자"],
        },
    },
}

# 상품 구분(직매장 종류)
PRODUCT_CATS = ["대월", "프리미엄"]
DEFAULT_CAT = "대월"
TOTAL_TITLE_TPL = "{jm}({tag})스마트오더 TOTAL 수량"


def product_rep_cat(pm, code):
    """상품마스터에서 (대표=한익스상품명, 구분) 반환. 값이 문자열이면 레거시(구분=대월).
    미등록이면 (None, None)."""
    v = pm.get(code)
    if v is None:
        return None, None
    if isinstance(v, dict):
        return v.get("rep", ""), v.get("cat", DEFAULT_CAT)
    return v, DEFAULT_CAT


def make_title(cfg, cat):
    """상품 구분에 따른 라벨 제목. 대월→대월직매장, 프리미엄→프리미엄직매장."""
    jm = (cat or DEFAULT_CAT) + "직매장"
    return cfg["title_tpl"].format(jm=jm)


def cover_title(cfg, rows):
    """표지 제목. 포함된 구분이 하나면 그 구분, 여러 개면 '대월·프리미엄'."""
    cats = [r.get("구분") or DEFAULT_CAT for r in rows]
    uniq = sorted(set(cats), key=lambda c: (PRODUCT_CATS.index(c) if c in PRODUCT_CATS else 9))
    jm = ("·".join(uniq) if uniq else DEFAULT_CAT) + "직매장"
    return TOTAL_TITLE_TPL.format(jm=jm, tag=cfg["tag"])


# ----------------------------------------------------------------------------
# 유틸
# ----------------------------------------------------------------------------
def norm(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s

def to_int(v):
    try:
        f = float(v)
        return int(f) if f.is_integer() else f
    except Exception:
        return v

# ----------------------------------------------------------------------------
# 마스터 로드/저장
# ----------------------------------------------------------------------------
def load_master(fname):
    p = os.path.join(MASTER_DIR, fname)
    if not os.path.exists(p):
        return {}
    with open(p, "r", encoding="utf-8") as f:
        return json.load(f)

def save_master(fname, data):
    p = os.path.join(MASTER_DIR, fname)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1, sort_keys=True)

def all_hubs():
    """거점센터(이고센터) 후보 목록 — 관리 UI 드롭다운용."""
    hubs = set()
    for fn in ("center_cu.json", "center_gs.json", "center_e24.json"):
        hubs.update(load_master(fn).values())
    return sorted(h for h in hubs if h)

# ----------------------------------------------------------------------------
# RAW 파싱
# ----------------------------------------------------------------------------
def _read_any(path):
    """xlsx/xls 모두 [ [row], ... ] 2차원 리스트로."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        import xlrd
        wb = xlrd.open_workbook(path)
        ws = wb.sheet_by_index(0)
        return ws.name, [[ws.cell_value(i, j) for j in range(ws.ncols)] for i in range(ws.nrows)]
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    name = ws.title
    wb.close()
    return name, rows

def _pick_sheet(path, hints):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        import xlrd
        wb = xlrd.open_workbook(path)
        best = wb.sheet_by_index(0)
        for ws in wb.sheets():
            if any(h in ws.name for h in hints):
                best = ws; break
        return best.name, [[best.cell_value(i, j) for j in range(best.ncols)] for i in range(best.nrows)]
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    target = wb.worksheets[0]
    for ws in wb.worksheets:
        if any(h in ws.title for h in hints):
            target = ws; break
    rows = [list(r) for r in target.iter_rows(values_only=True)]
    name = target.title
    wb.close()
    return name, rows

def _find_header(rows, cfg):
    """헤더행 index와 각 필드 컬럼 index를 찾는다."""
    scan = cfg["raw"].get("header_scan_rows", 3)
    want = {k: cfg["raw"][k] for k in ("center_code","center_name","store_code","store_name","prod_code","prod_name","qty","date")}
    best_row, best_idx, best_hits = 0, {}, -1
    for r in range(min(scan + 1, len(rows))):
        hdr = [norm(x) for x in rows[r]]
        idx = {}
        for key, cands in want.items():
            # exact 우선, 없으면 부분일치
            found = None
            for i, h in enumerate(hdr):
                if h in cands:
                    found = i; break
            if found is None:
                for i, h in enumerate(hdr):
                    if any(c in h for c in cands):
                        found = i; break
            if found is not None:
                idx[key] = found
        hits = len(idx)
        if hits > best_hits:
            best_row, best_idx, best_hits = r, idx, hits
    return best_row, best_idx

def parse_raw(path, channel):
    cfg = CHANNELS[channel]
    _, rows = _pick_sheet(path, cfg["raw"]["sheet_hint"])
    hrow, idx = _find_header(rows, cfg)
    required = ["center_name", "prod_code", "qty"]
    missing = [k for k in required if k not in idx]
    if missing:
        raise ValueError("RAW에서 필수 컬럼을 찾지 못했습니다: %s (헤더행 인식 실패)" % ", ".join(missing))

    def g(r, key):
        i = idx.get(key)
        if i is None or i >= len(r):
            return ""
        return norm(r[i])

    records = []
    for r in rows[hrow + 1:]:
        if not any(norm(x) for x in r):
            continue
        cname = g(r, "center_name")
        ccode = g(r, "center_code")
        pcode = g(r, "prod_code")
        pname = g(r, "prod_name")
        qtxt  = g(r, "qty")
        if not (cname or ccode) and not pcode:
            continue
        if not pcode and not pname:
            continue
        qty = to_int(qtxt) if qtxt else ""
        records.append({
            "center_code": ccode,
            "center_name": cname,
            "store_code": g(r, "store_code"),
            "store_name": g(r, "store_name") or cname,
            "prod_code": pcode,
            "prod_name": pname,
            "qty": qty,
            "date": g(r, "date"),
        })
    return records

# ----------------------------------------------------------------------------
# 매핑 처리
# ----------------------------------------------------------------------------
def process(records, channel):
    """레코드에 거점센터/대표상품명을 코드 기반으로 매핑.
    반환: (rows, unmapped_centers, unmapped_products)
    rows: 붙여넣기(데이터추출)용 dict 리스트
    """
    cfg = CHANNELS[channel]
    center_master = load_master(cfg["center_master"])
    product_master = load_master("product.json")
    # 보조: 화주상품명 -> {rep,cat} (코드 미등록 시 이름 폴백). 정규화 키로 인덱싱.
    product_names = load_master("product_names.json")
    pn_idx = {norm(k): v for k, v in product_names.items()}
    ckey = cfg["center_key"]

    rows = []
    unmapped_centers = {}   # key -> 대표명칭(표시용)
    unmapped_products = {}  # code -> name
    for rec in records:
        ckeyval = rec["center_code"] if ckey == "code" else rec["center_name"]
        hub = center_master.get(ckeyval, "")
        if not hub and ckeyval:
            unmapped_centers[ckeyval] = rec["center_name"] or rec["center_code"]
        # 1) 코드 우선 → 2) 화주상품명 폴백
        rep, cat = product_rep_cat(product_master, rec["prod_code"])
        if rep is None:
            v = pn_idx.get(norm(rec["prod_name"]))
            if v:
                rep, cat = v.get("rep", ""), v.get("cat", DEFAULT_CAT)
        if rep is None and rec["prod_code"]:
            unmapped_products[rec["prod_code"]] = rec["prod_name"]
        cat = cat or DEFAULT_CAT
        # 부착양식 라벨 = 화주상품명(원본 제품명), 표지/집계 = 대표(한익스상품명)
        label_prod = rec["prod_name"] if cfg["label_product"] == "raw" else (rep or rec["prod_name"])
        rows.append({
            "센터": rec["center_name"],
            "배송일자": rec["date"],
            "점포코드": rec["store_code"],
            "점포명": rec["store_name"],
            "거점센터": hub,
            "상품명": label_prod,          # 라벨 표시용 = 화주상품명
            "수량": rec["qty"],
            "대표": rep or rec["prod_name"],   # 표지 표시용 = 한익스상품명(대표)
            "구분": cat,                       # 대월 / 프리미엄 (라벨 제목 분기)
            "_ckey": ckeyval,
            "_pcode": rec["prod_code"],
        })
    return rows, unmapped_centers, unmapped_products

def sort_rows(rows):
    """구분(대월/프리미엄) → 거점센터 → 대표상품 → 점포명 순 정렬(직매장·창고 분류 최적화)."""
    def catkey(c):
        return PRODUCT_CATS.index(c) if c in PRODUCT_CATS else 9
    return sorted(rows, key=lambda x: (catkey(x.get("구분") or DEFAULT_CAT),
                                       x["거점센터"] or "zzz", x["대표"] or "", x["점포명"] or ""))


def compute_totals(rows):
    """TOTAL 표지용 집계: 상품(대표)별 수량, 거점센터별 수량. (수량 내림차순)"""
    prod, hub = {}, {}
    for row in rows:
        q = row["수량"] if isinstance(row["수량"], (int, float)) else 0
        if row["대표"]:
            prod[row["대표"]] = prod.get(row["대표"], 0) + q
        if row["거점센터"]:
            hub[row["거점센터"]] = hub.get(row["거점센터"], 0) + q
    prod_list = sorted(prod.items(), key=lambda x: -x[1])
    hub_list = sorted(hub.items(), key=lambda x: -x[1])
    return prod_list, hub_list

# ----------------------------------------------------------------------------
# 스타일 (기존 부착양식과 동일)
# ----------------------------------------------------------------------------
FONT = "맑은 고딕"
_thin = Side(style="thin", color="FF000000")
BORDER_BOX = Border(top=_thin, bottom=_thin, left=_thin, right=_thin)
FILL_YELLOW = PatternFill("solid", fgColor="FFFFFF00")
FILL_WARN = PatternFill("solid", fgColor=Color(theme=1, tint=0.05))
AL_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_C_NW = Alignment(horizontal="center", vertical="center", wrap_text=False)
AL_V = Alignment(vertical="center")
AL_CNUM = Alignment(vertical="center", wrap_text=True)   # 노란 C열(라벨번호) — 원본과 동일

def _cnum_cell(cell, value=None):
    cell.value = value
    cell.font = Font(name=FONT, size=48, bold=True)
    cell.fill = FILL_YELLOW
    cell.alignment = AL_CNUM

def _style_cell(cell, value, size, bold=True, border=True, fill=None, align=AL_C_NW):
    cell.value = value
    cell.font = Font(name=FONT, size=size, bold=bold)
    cell.alignment = align
    if border:
        cell.border = BORDER_BOX
    if fill:
        cell.fill = fill

def build_label_sheet(ws, rows, cfg):
    # 열 너비
    for col, w in cfg["col_width"].items():
        ws.column_dimensions[col].width = w
    r = 1
    if cfg["lead_spacer"]:
        ws.row_dimensions[r].height = 25.5
        _cnum_cell(ws.cell(r, 3))
        r += 1
    rh = cfg["row_h"]
    for n, row in enumerate(rows, start=1):
        # 1) 제목 (상품 구분에 따라 대월/프리미엄 직매장)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        _style_cell(ws.cell(r, 1), make_title(cfg, row.get("구분")), 62, fill=None, align=AL_C)
        _cnum_cell(ws.cell(r, 3), n)
        ws.row_dimensions[r].height = 169.95
        r += 1
        # 2) 개봉금지
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        _style_cell(ws.cell(r, 1), "※수령자 외 절대 개봉금지※", 40, fill=FILL_WARN, align=AL_C)
        _cnum_cell(ws.cell(r, 3))
        ws.row_dimensions[r].height = 63.6
        r += 1
        # 3~7) 라벨 항목
        items = [
            ("이고센터", row["거점센터"], 42),
            (cfg["center_label"], row["센터"], 42),
            (" 점포명", row["점포명"], 48),
            (" 상품 ", row["상품명"], 48),
            (" 수량", row["수량"], 48),
        ]
        for label, value, bsize in items:
            _style_cell(ws.cell(r, 1), label, 42)
            _style_cell(ws.cell(r, 2), value, bsize)
            _cnum_cell(ws.cell(r, 3))
            ws.row_dimensions[r].height = rh
            r += 1
        # 8) 여백
        cc = ws.cell(r, 3); cc.font = Font(name=FONT, size=48, bold=True); cc.fill = FILL_YELLOW; cc.alignment = AL_V
        ws.row_dimensions[r].height = 26.0
        r += 1

def build_data_sheet(ws, rows):
    headers = ["센터", "배송일자", "점포코드", "점포명", "거점센터", "상품명", "수량", "상품명(대표)"]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        ws.cell(1, c).font = Font(name=FONT, size=11, bold=True)
    for row in rows:
        ws.append([row["센터"], row["배송일자"], row["점포코드"], row["점포명"],
                   row["거점센터"], row["상품명"], row["수량"], row["대표"]])
    widths = [22, 12, 10, 24, 14, 30, 8, 26]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def build_total_sheet(ws, rows, cfg):
    from collections import OrderedDict
    prod = OrderedDict()
    hub = OrderedDict()
    for row in rows:
        q = row["수량"] if isinstance(row["수량"], (int, float)) else 0
        prod[row["대표"]] = prod.get(row["대표"], 0) + q
        if row["거점센터"]:
            hub[row["거점센터"]] = hub.get(row["거점센터"], 0) + q
    ws.cell(1, 2, cover_title(cfg, rows)).font = Font(name=FONT, size=14, bold=True)
    hdr = ["번호", "상품", "수량"]
    for i, h in enumerate(hdr):
        ws.cell(2, 2 + i, h).font = Font(name=FONT, size=11, bold=True)
    ws.cell(2, 6, "이고센터명").font = Font(name=FONT, size=11, bold=True)
    ws.cell(2, 7, "이고수량").font = Font(name=FONT, size=11, bold=True)
    r = 3
    for i, (k, v) in enumerate(sorted(prod.items(), key=lambda x: -x[1]), 1):
        ws.cell(r, 2, i); ws.cell(r, 3, k); ws.cell(r, 4, v); r += 1
    r = 3
    for k, v in sorted(hub.items(), key=lambda x: -x[1]):
        ws.cell(r, 6, k); ws.cell(r, 7, v); r += 1
    for col, w in {"B": 6, "C": 34, "D": 8, "E": 3, "F": 16, "G": 10}.items():
        ws.column_dimensions[col].width = w

def generate_workbook(rows, channel):
    cfg = CHANNELS[channel]
    rows = sort_rows(rows)
    wb = openpyxl.Workbook()
    try:
        wb._named_styles["Normal"].font = Font(name=FONT, size=11)  # 빈 셀 기본폰트도 맑은 고딕
    except Exception:
        pass
    ws_total = wb.active; ws_total.title = "TOTAL(표지)"
    build_total_sheet(ws_total, rows, cfg)
    ws_label = wb.create_sheet("부착양식(A4출력)")
    build_label_sheet(ws_label, rows, cfg)
    ws_data = wb.create_sheet("붙여넣기(데이터추출)")
    build_data_sheet(ws_data, rows)
    # 인쇄 설정: A4 세로, 여백 최소, 열 폭 맞춤
    ws_label.page_setup.orientation = "portrait"
    ws_label.page_setup.paperSize = 9  # A4
    ws_label.page_setup.fitToWidth = 1
    ws_label.page_setup.fitToHeight = 0
    ws_label.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws_label.print_area = "A1:C%d" % ws_label.max_row
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio
