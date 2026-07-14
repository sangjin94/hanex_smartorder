# -*- coding: utf-8 -*-
"""원본 'OB스마트오더 자동화' 엑셀에서 마스터를 재생성한다(1회용 / 복구용).

생성물 (masters/)
 - center.json        : {센터명: 이고센터}      전 채널 공통. 원본 '점포마스터' 시트가 원 출처
 - product.json       : {상품코드: {rep, cat}}  전 채널 공통. 같은 상품이라도 채널마다 코드가 다름
 - product_names.json : {화주상품명: {rep, cat}} 상품코드 미등록 시 이름 폴백

사용:  py bootstrap_masters.py [원본폴더]
      (생략 시 SMARTORDER_SRC 환경변수 → 바탕화면 'OB스마트오더 자동화')
"""
import os, sys, io
import openpyxl
import smartorder_core as sc

DEFAULT_SRC = os.path.join(os.path.expanduser("~"), "Desktop", "OB스마트오더 자동화")
SRC = (sys.argv[1] if len(sys.argv) > 1 else os.environ.get("SMARTORDER_SRC") or DEFAULT_SRC)

# 출력양식 파일: 점포마스터(센터명→이고센터) + 상품마스터(대표) + lowdata(라벨출력) 보유
FORMS = [
    ("cu",  "(CU)스마트오더 출력양식.xlsx",     "점포마스터"),
    ("gs",  "GS 스마트오더 출력양식.xlsx",      "점포마스터"),
    ("e24", "이마트24 스마트오더 출력양식.xlsx", "E24 점포마스터"),
]
# 리테일러가 보내주는 RAW 원본(상품코드 수집용)
RAWS = [
    ("cu",  "CU 스마트오더.xlsx"),
    ("gs",  "GS 스마트오더.xls"),
    ("e24", "이마트24 스마트오더.xlsx"),
]
# 원본 점포마스터의 알려진 오류 교정 — 제주 물량은 이천센터에서 나간다.
# (시트에 같은 센터명이 '제주도'/'제주센터'로도 중복 등록돼 있음)
CENTER_OVERRIDES = {
    "BGF로지스제주": "이천센터",
    "BGF로지스서귀포": "이천센터",
    "서귀포센터": "이천센터",
    "제주센터": "이천센터",
    "제주주류": "이천센터",
    "지에스리테일신제주애월상온": "이천센터",
}

log = io.StringIO()
def L(*a):
    print(*a, file=log)
    # 콘솔이 cp949라 한글 외 기호에서 죽는 경우가 있어 안전 출력
    try:
        print(*a)
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "utf-8"
        print(" ".join(str(x) for x in a).encode(enc, "replace").decode(enc))


def sheet(path, name):
    wb = openpyxl.load_workbook(path, data_only=True)
    return wb[name]


# ---------- 1) 센터마스터: 점포마스터 D(센터명) → G(지방거점) ----------
center = {}
for _, fn, sh in FORMS:
    ws = sheet(os.path.join(SRC, fn), sh)
    n = 0
    for r in range(2, ws.max_row + 1):
        nm = sc.norm(ws.cell(r, 4).value)
        hub = sc.norm(ws.cell(r, 7).value)
        if nm and hub and nm not in center:   # 시트에 중복 행이 있어 first-wins
            center[nm] = hub
            n += 1
    L("점포마스터 %-30s → 신규 센터 %d개" % (sh + "(" + fn + ")", n))
center.update(CENTER_OVERRIDES)
L("센터마스터: %d개 · 이고센터 %s" % (len(center), sorted(set(center.values()))))

# ---------- 2) 화주상품명 → (한익스상품명, 구분) : GS 파일 '상품마스터(대표)' ----------
pn2rep = {}
ws = sheet(os.path.join(SRC, "GS 스마트오더 출력양식.xlsx"), "상품마스터(대표)")
for r in range(2, ws.max_row + 1):
    name = sc.norm(ws.cell(r, 1).value)   # 제품(화주상품명)
    rep = sc.norm(ws.cell(r, 2).value)    # 대표(한익스상품명)
    cat = sc.norm(ws.cell(r, 3).value)
    cat = cat if cat in sc.PRODUCT_CATS else sc.DEFAULT_CAT
    if name and rep and name not in pn2rep:
        pn2rep[name] = {"rep": rep, "cat": cat}
L("화주상품명 → 한익스상품명: %d개" % len(pn2rep))

# ---------- 3) 상품마스터: 상품코드 → {rep, cat} ----------
# 같은 상품이라도 채널마다 코드가 달라, 여러 코드가 같은 rep을 가리킨다.
product = {}
unmapped = set()
sources = [(ch, os.path.join(SRC, fn)) for ch, fn, _ in FORMS]
sources += [(ch, os.path.join(SRC, fn)) for ch, fn in RAWS]
for ch, path in sources:
    if not os.path.exists(path):
        L("  (건너뜀, 파일 없음) %s" % os.path.basename(path))
        continue
    try:
        recs = sc.parse_raw(path, ch)
    except Exception as e:
        L("  (건너뜀, 파싱 실패) %s: %s" % (os.path.basename(path), e))
        continue
    added = 0
    for rec in recs:
        code, name = rec["prod_code"], rec["prod_name"]
        if not code or code in product:
            continue
        v = pn2rep.get(name)
        if v:
            product[code] = dict(v)
            added += 1
        else:
            unmapped.add((code, name))
    L("  %-30s 레코드 %4d → 신규 상품코드 %d" % (os.path.basename(path), len(recs), added))

reps = {v["rep"] for v in product.values()}
L("상품마스터: 코드 %d개 / 한익스상품명 %d개 (같은 상품의 채널별 코드가 묶임)" % (len(product), len(reps)))
if unmapped:
    L("한익스상품명을 못 찾은 코드 %d개 (화면에서 등록 필요):" % len(unmapped))
    for code, name in sorted(unmapped)[:20]:
        L("   %s  %s" % (code, name))

# ---------- 저장 ----------
# 재실행해도 화면에서 사람이 등록/수정한 매핑이 날아가지 않도록 '기존 값 우선' 병합.
# (빈 마스터면 전체 생성, 이미 있으면 빠진 항목만 채움. 단 CENTER_OVERRIDES는 항상 적용)
def merge_save(fname, fresh, force=None):
    cur = sc.load_master(fname)
    added = 0
    for k, v in fresh.items():
        if k not in cur:
            cur[k] = v
            added += 1
    if force:
        cur.update(force)
    sc.save_master(fname, cur)
    L("  %-20s 기존 %d개 유지 + 신규 %d개 → 총 %d개" % (fname, len(cur) - added, added, len(cur)))

L("저장(기존 값 우선 병합) → %s" % sc.MASTER_DIR)
merge_save(sc.CENTER_MASTER, center, force=CENTER_OVERRIDES)
merge_save("product.json", product)
merge_save("product_names.json", pn2rep)

with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "bootstrap_log.txt"), "w", encoding="utf-8") as f:
    f.write(log.getvalue())
