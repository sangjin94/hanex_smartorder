# -*- coding: utf-8 -*-
"""기존 OB스마트오더 자동화 파일에서 코드 기반 마스터를 부트스트랩한다.
 - masters/center_cu.json   : {센터코드: 거점센터}
 - masters/center_gs.json   : {센터코드: 거점센터}   (GS25 VA코드 + GS슈퍼 숫자코드 통합)
 - masters/center_e24.json  : {입고센터명: 거점센터}  (E24는 센터코드가 없어 명칭키)
 - masters/product.json     : {상품코드: 대표상품명}   (전 채널 통합)
 - masters/center_name_hint_*.json : {센터코드: 센터명}  참고용(관리 UI 표시)
"""
import openpyxl, xlrd, os, json, io

SRC = r"C:\Users\HanEx\Desktop\OB스마트오더 자동화"
OUT = os.path.join(os.path.dirname(__file__), "masters")
os.makedirs(OUT, exist_ok=True)
log = io.StringIO()
def L(*a): print(*a, file=log)

def norm(v):
    if v is None: return ""
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit(): s = s[:-2]
    return s

def load_xlsx(path, sheet):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows

# ---------- 1) 점포마스터: 리테일러 센터명 -> 거점센터 ----------
def name2hub_from_master(path, sheet, name_col=3, hub_col=6):
    rows = load_xlsx(path, sheet)
    m = {}
    for r in rows[1:]:
        if len(r) <= hub_col: continue
        name = norm(r[name_col]); hub = norm(r[hub_col])
        if name and hub and name not in m:
            m[name] = hub
    return m

cu_n2h = name2hub_from_master(os.path.join(SRC,"(CU)스마트오더 출력양식.xlsx"), "점포마스터")
gs_n2h = name2hub_from_master(os.path.join(SRC,"GS 스마트오더 출력양식.xlsx"), "점포마스터")
e24_n2h = name2hub_from_master(os.path.join(SRC,"이마트24 스마트오더 출력양식.xlsx"), "E24 점포마스터")
L("점포마스터 name->hub  CU:%d  GS:%d  E24:%d" % (len(cu_n2h), len(gs_n2h), len(e24_n2h)))

# ---------- 2) 상품마스터: 상품명 -> (대표, 구분) ----------
def prodname2rep(path, sheet="상품마스터(대표)"):
    rows = load_xlsx(path, sheet)
    m = {}
    for r in rows[1:]:
        if len(r) < 2: continue
        name = norm(r[0]); rep = norm(r[1])
        cat = norm(r[2]) if len(r) > 2 else ""
        cat = cat if cat in ("대월", "프리미엄") else "대월"
        if name and rep and name not in m:
            m[name] = (rep, cat)
    return m
pn2rep = prodname2rep(os.path.join(SRC,"GS 스마트오더 출력양식.xlsx"))
from collections import Counter as _C
L("상품마스터 name->(rep,cat) : %d  구분분포=%s" % (len(pn2rep), dict(_C(c for _,c in pn2rep.values()))))

# ---------- 3) CU: 센터코드->센터명 (발주확인 시트) ----------
def cols_by_header(rows, header_row, wanted):
    """헤더행에서 원하는 라벨의 컬럼 index를 부분일치로 찾는다."""
    hdr = [norm(x) for x in rows[header_row]]
    idx = {}
    for key, cands in wanted.items():
        for i, h in enumerate(hdr):
            if any(c in h for c in cands):
                idx[key] = i; break
    return idx

cu_rows = load_xlsx(os.path.join(SRC,"CU 스마트오더.xlsx"), "발주확인 및 확정내역(스티커제작용)")
cu_idx = cols_by_header(cu_rows, 0, {"code":["센터 코드","센터코드"], "name":["센터명"], "pcode":["상품 코드","상품코드"], "pname":["상품명"]})
L("CU 헤더 idx: %s" % cu_idx)
cu_code2hub = {}; cu_code2name = {}; cu_pcode2rep = {}; cu_unmapped_center=set(); cu_unmapped_prod=set()
for r in cu_rows[2:]:
    if len(r) <= max(cu_idx.values()): continue
    code = norm(r[cu_idx["name"]] if False else r[cu_idx["code"]]); name = norm(r[cu_idx["name"]])
    if code and name:
        cu_code2name.setdefault(code, name)
        hub = cu_n2h.get(name)
        if hub: cu_code2hub[code] = hub
        else: cu_unmapped_center.add((code,name))
    pcode = norm(r[cu_idx["pcode"]]); pname = norm(r[cu_idx["pname"]])
    if pcode and pname:
        rep = pn2rep.get(pname)
        if rep: cu_pcode2rep[pcode] = rep
        else: cu_unmapped_prod.add((pcode,pname))
L("CU 센터코드->hub: %d (미매핑 %d)  상품코드->대표: %d (미매핑 %d)" %
  (len(cu_code2hub), len(cu_unmapped_center), len(cu_pcode2rep), len(cu_unmapped_prod)))

# ---------- 4) GS슈퍼 lowdata: 센터코드->센터 ----------
gs_code2hub = {}; gs_code2name = {}; gs_pcode2rep = {}; gs_unmapped_center=set(); gs_unmapped_prod=set()
gsl = load_xlsx(os.path.join(SRC,"GS 스마트오더 출력양식.xlsx"), "lowdata(라벨출력)")
gs_idx = cols_by_header(gsl, 0, {"code":["센터코드"], "name":["센터"], "pcode":["상품코드"], "pname":["상품명"]})
# '센터' 부분일치가 '센터코드'를 먼저 잡을 수 있어 재조정
def find_exact(rows, label):
    hdr=[norm(x) for x in rows[0]]
    for i,h in enumerate(hdr):
        if h==label: return i
    return None
gs_idx["code"]=find_exact(gsl,"센터코드"); gs_idx["name"]=find_exact(gsl,"센터")
gs_idx["pcode"]=find_exact(gsl,"상품코드"); gs_idx["pname"]=find_exact(gsl,"상품명")
L("GS슈퍼 헤더 idx: %s" % gs_idx)
for r in gsl[1:]:
    code=norm(r[gs_idx["code"]]); name=norm(r[gs_idx["name"]])
    if code and name:
        gs_code2name.setdefault(code,name)
        hub=gs_n2h.get(name)
        if hub: gs_code2hub[code]=hub
        else: gs_unmapped_center.add((code,name))
    pcode=norm(r[gs_idx["pcode"]]); pname=norm(r[gs_idx["pname"]])
    if pcode and pname:
        rep=pn2rep.get(pname)
        if rep: gs_pcode2rep[pcode]=rep
        else: gs_unmapped_prod.add((pcode,pname))

# ---------- 5) GS25 .xls: 센터코드(VA..)->센터 ----------
def load_xls(path):
    wb=xlrd.open_workbook(path); ws=wb.sheet_by_index(0)
    return [[ws.cell_value(i,j) for j in range(ws.ncols)] for i in range(ws.nrows)]
g25=load_xls(os.path.join(SRC,"GS 스마트오더.xls"))
h=[norm(x) for x in g25[0]]
def gi(label):
    for i,x in enumerate(h):
        if x==label: return i
    return None
c_code=gi("센터코드"); c_name=gi("센터"); c_pcode=gi("상품코드"); c_pname=gi("상품명")
L("GS25 헤더 idx: code=%s name=%s pcode=%s pname=%s" % (c_code,c_name,c_pcode,c_pname))
for r in g25[1:]:
    code=norm(r[c_code]); name=norm(r[c_name])
    if code and name:
        gs_code2name.setdefault(code,name)
        hub=gs_n2h.get(name)
        if hub: gs_code2hub[code]=hub
        else: gs_unmapped_center.add((code,name))
    pcode=norm(r[c_pcode]); pname=norm(r[c_pname])
    if pcode and pname:
        rep=pn2rep.get(pname)
        if rep: gs_pcode2rep[pcode]=rep
        else: gs_unmapped_prod.add((pcode,pname))
L("GS 센터코드->hub: %d (미매핑 %d)  상품코드->대표: %d (미매핑 %d)" %
  (len(gs_code2hub), len(gs_unmapped_center), len(gs_pcode2rep), len(gs_unmapped_prod)))

# ---------- 6) E24: 입고센터명->hub (점포마스터 GS점포명 컬럼이 곧 입고센터명) ----------
e24_center = dict(e24_n2h)   # {'양산(상온)': '양산센터', ...}
# E24 상품코드->대표 (lowdata)
e24_pcode2rep={}; e24_unmapped_prod=set()
e24l=load_xlsx(os.path.join(SRC,"이마트24 스마트오더 출력양식.xlsx"), "lowdata(라벨출력)")
eh=[norm(x) for x in e24l[0]]
def ei(*labels):
    for i,x in enumerate(eh):
        if x in labels: return i
    return None
ep=ei("상품코드"); en=ei("앱 상품명","앱상품명"); erep=ei("상품명(대표)")
for r in e24l[1:]:
    if ep is None: break
    pcode=norm(r[ep]) if ep<len(r) else ""
    if not pcode: continue
    appname = norm(r[en]) if (en is not None and en<len(r)) else ""
    pr = pn2rep.get(appname)                        # (rep, cat)
    if pr is None:
        rep = norm(r[erep]) if (erep is not None and erep<len(r)) else ""
        if rep: pr = (rep, "대월")
    if pr: e24_pcode2rep[pcode]=pr
L("E24 입고센터->hub: %d   상품코드->대표: %d" % (len(e24_center), len(e24_pcode2rep)))

# ---------- 통합 상품마스터: {코드: {rep, cat}} ----------
product = {}
for d in (cu_pcode2rep, gs_pcode2rep, e24_pcode2rep):
    for code,(rep,cat) in d.items():
        product[code] = {"rep": rep, "cat": cat}
from collections import Counter as _C2
L("통합 상품코드->{대표,구분}: %d  구분분포=%s" % (len(product), dict(_C2(v['cat'] for v in product.values()))))

def dump(name, obj):
    with open(os.path.join(OUT,name),"w",encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1, sort_keys=True)

dump("center_cu.json", cu_code2hub)
dump("center_gs.json", gs_code2hub)
dump("center_e24.json", e24_center)
dump("product.json", product)
# 보조 시드: 화주상품명 -> {대표, 구분}. 코드 미등록 시 이름으로 폴백(특히 프리미엄 자동분류)
dump("product_names.json", {name: {"rep": rep, "cat": cat} for name, (rep, cat) in pn2rep.items()})
dump("center_cu_names.json", cu_code2name)
dump("center_gs_names.json", gs_code2name)
# 미매핑 목록(참고)
dump("_unmapped.json", {
    "cu_center":[list(x) for x in sorted(cu_unmapped_center)],
    "gs_center":[list(x) for x in sorted(gs_unmapped_center)],
    "cu_prod":[list(x) for x in sorted(cu_unmapped_prod)][:50],
    "gs_prod":[list(x) for x in sorted(gs_unmapped_prod)][:50],
})

with open(os.path.join(os.path.dirname(__file__),"bootstrap_log.txt"),"w",encoding="utf-8") as f:
    f.write(log.getvalue())
print("OK")
