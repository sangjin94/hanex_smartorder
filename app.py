# -*- coding: utf-8 -*-
"""한익스프레스 스마트오더 라벨 생성기 (웹)

기능
 - 판매채널(CU/GS/E24)별 RAW 업로드 → 코드 기반 매핑 → 부착양식(A4) 라벨 xlsx 생성
 - 업로드 중 미매핑 센터/상품 코드를 화면에서 즉시 마스터에 등록
 - 마스터 관리(센터명→이고센터[전 채널 공통], 상품코드→한익스상품명) 조회/추가/수정/삭제
"""
import os, io, json, re, uuid, datetime
from flask import (Flask, request, render_template, send_file, redirect,
                   url_for, flash, jsonify, abort)
import smartorder_core as sc

app = Flask(__name__)
app.secret_key = "hanex-smartorder-label-2026"
app.config["MAX_CONTENT_LENGTH"] = 40 * 1024 * 1024  # 40MB

# nginx 서브경로(/smartorder/) 배포 시 url_for가 접두어를 붙이도록 처리
class PrefixMiddleware:
    def __init__(self, wsgi_app):
        self.app = wsgi_app
    def __call__(self, environ, start_response):
        prefix = environ.get("HTTP_X_FORWARDED_PREFIX", "")
        if prefix:
            environ["SCRIPT_NAME"] = prefix.rstrip("/")
        return self.app(environ, start_response)
app.wsgi_app = PrefixMiddleware(app.wsgi_app)

sc.ensure_masters_seeded()

CH_ORDER = ["cu", "gs", "e24"]

# ---------------- 업로드 아카이브(영구 누적) ----------------
# 배포 시 git 밖 영구 디렉터리로 분리(SMARTORDER_ARCHIVE). 기본은 앱 폴더 archive/.
ARCHIVE_DIR = os.environ.get("SMARTORDER_ARCHIVE") or os.path.join(sc.BASE, "archive")
INDEX_PATH = os.path.join(ARCHIVE_DIR, "index.json")
os.makedirs(ARCHIVE_DIR, exist_ok=True)

def _load_index():
    if not os.path.exists(INDEX_PATH):
        return []
    try:
        with open(INDEX_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def _save_index(items):
    with open(INDEX_PATH, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=1)

def _safe_name(name):
    base = os.path.basename(name)
    return re.sub(r'[\\/:*?"<>|]+', "_", base).strip() or "raw"

def _rec_by_id(uid):
    for r in _load_index():
        if r["id"] == uid:
            return r
    return None

def _rec_path(rec):
    return os.path.join(ARCHIVE_DIR, rec["ch"], rec["fname"])


@app.route("/")
def index():
    centers = len(sc.load_master(sc.CENTER_MASTER))
    products = len(sc.load_master("product.json"))
    return render_template("index.html", channels=CH_ORDER, cfg=sc.CHANNELS,
                           centers=centers, products=products)


@app.route("/u/<ch>", methods=["GET", "POST"])
def upload(ch):
    if ch not in sc.CHANNELS:
        abort(404)
    cfg = sc.CHANNELS[ch]
    if request.method == "GET":
        recent = [r for r in reversed(_load_index()) if r["ch"] == ch][:8]
        return render_template("upload.html", ch=ch, cfg=cfg, recent=recent)

    f = request.files.get("file")
    if not f or not f.filename:
        flash("파일을 선택해 주세요.", "error")
        return redirect(url_for("upload", ch=ch))
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".xlsx", ".xls"):
        flash("xlsx 또는 xls 파일만 업로드할 수 있습니다.", "error")
        return redirect(url_for("upload", ch=ch))

    # 채널 폴더에 영구 저장(누적)
    now = datetime.datetime.now()
    uid = "%s_%s_%s" % (ch, now.strftime("%Y%m%d_%H%M%S"), uuid.uuid4().hex[:4])
    fname = uid + "__" + _safe_name(f.filename)
    os.makedirs(os.path.join(ARCHIVE_DIR, ch), exist_ok=True)
    path = os.path.join(ARCHIVE_DIR, ch, fname)
    f.save(path)

    # 통계 미리 계산(이력 표시용) — 실패해도 파일은 보관
    nrec = nrows = total_qty = None
    try:
        records = sc.parse_raw(path, ch)
        rows, uc, up = sc.process(records, ch)
        nrec, nrows = len(records), len(rows)
        total_qty = sum(r["수량"] for r in rows if isinstance(r["수량"], (int, float)))
    except Exception:
        pass

    idx = _load_index()
    idx.append({"id": uid, "ch": ch, "orig": f.filename, "fname": fname,
                "uploaded_at": now.strftime("%Y-%m-%d %H:%M:%S"),
                "nrec": nrec, "nrows": nrows, "total_qty": total_qty})
    _save_index(idx)
    return redirect(url_for("result", ch=ch, token=uid))


def _load_job(ch, token):
    rec = _rec_by_id(token)
    if not rec or rec["ch"] != ch:
        return None
    path = _rec_path(rec)
    if not os.path.exists(path):
        return None
    return {"channel": ch, "path": path, "filename": rec["orig"]}


@app.route("/r/<ch>/<token>")
def result(ch, token):
    if ch not in sc.CHANNELS:
        abort(404)
    job = _load_job(ch, token)
    if not job:
        flash("업로드 세션이 만료되었습니다. 다시 업로드해 주세요.", "error")
        return redirect(url_for("upload", ch=ch))
    cfg = sc.CHANNELS[ch]
    try:
        records = sc.parse_raw(job["path"], ch)
    except Exception as e:
        flash("RAW 파일 분석 실패: %s" % e, "error")
        return redirect(url_for("upload", ch=ch))
    rows, uc, up = sc.process(records, ch)
    rows = sc.sort_rows(rows)
    total_qty = sum(r["수량"] for r in rows if isinstance(r["수량"], (int, float)))
    prod_list, hub_list = sc.compute_totals(rows)
    hubs = sc.all_hubs()
    # 기존 한익스상품명 목록(+구분) — 미매핑 코드를 기존 상품에 붙일 때 오타로 새 상품이 생기는 걸 방지
    rep_cats = sc.all_reps()
    return render_template("result.html", ch=ch, cfg=cfg, token=token,
                           filename=job["filename"], nrec=len(records),
                           nrows=len(rows), total_qty=total_qty,
                           unmapped_centers=uc, unmapped_products=up,
                           hubs=hubs, rep_cats=rep_cats,
                           prod_list=prod_list, hub_list=hub_list,
                           cover_title=sc.cover_title(cfg, rows),
                           product_cats=sc.PRODUCT_CATS)


def _block_if_missing_reps(rows, ch, token):
    """한익스상품명(대표)이 없는 상품이 하나라도 있으면 출력(다운로드·인쇄)을 막는다.
    표지에는 항상 한익스상품명만 들어가야 하므로, 화주상품명으로 대체 출력하지 않는다."""
    miss = sc.missing_reps(rows)
    if not miss:
        return None
    flash("한익스상품명이 등록되지 않은 상품이 있어 출력할 수 없습니다: %s — 아래에서 등록 후 다시 시도해 주세요."
          % ", ".join("%s(%s)" % (c, n) for c, n in list(miss.items())[:5]), "error")
    return redirect(url_for("result", ch=ch, token=token))


@app.route("/print/<ch>/<token>")
def print_labels(ch, token):
    if ch not in sc.CHANNELS:
        abort(404)
    job = _load_job(ch, token)
    if not job:
        flash("업로드 세션이 만료되었습니다.", "error")
        return redirect(url_for("upload", ch=ch))
    cfg = sc.CHANNELS[ch]
    records = sc.parse_raw(job["path"], ch)
    rows, uc, up = sc.process(records, ch)
    blocked = _block_if_missing_reps(rows, ch, token)
    if blocked:
        return blocked
    rows = sc.sort_rows(rows)
    for r in rows:
        r["title"] = sc.make_title(cfg, r.get("구분"))
    return render_template("print.html", ch=ch, cfg=cfg, rows=rows, token=token)


@app.route("/assign/<ch>/<token>", methods=["POST"])
def assign(ch, token):
    if ch not in sc.CHANNELS:
        abort(404)
    # 센터 매핑 저장(전 채널 공통 마스터, 키=센터명)
    cm = sc.load_master(sc.CENTER_MASTER)
    for key in request.form:
        if key.startswith("center::"):
            name = key[len("center::"):]
            val = request.form.get(key, "").strip()
            if val:
                cm[name] = val
    sc.save_master(sc.CENTER_MASTER, cm)
    # 상품 매핑 저장 (공통 마스터) — {대표(한익스), 구분}
    pm = sc.load_master("product.json")
    for key in request.form:
        if key.startswith("prod::"):
            code = key[len("prod::"):]
            rep = request.form.get(key, "").strip()
            cat = request.form.get("pcat::" + code, sc.DEFAULT_CAT).strip() or sc.DEFAULT_CAT
            if rep:
                pm[code] = {"rep": rep, "cat": cat}
    sc.save_master("product.json", pm)
    flash("매핑을 저장했습니다.", "ok")
    return redirect(url_for("result", ch=ch, token=token))


@app.route("/download/<ch>/<token>")
def download(ch, token):
    if ch not in sc.CHANNELS:
        abort(404)
    job = _load_job(ch, token)
    if not job:
        flash("업로드 세션이 만료되었습니다.", "error")
        return redirect(url_for("upload", ch=ch))
    cfg = sc.CHANNELS[ch]
    records = sc.parse_raw(job["path"], ch)
    rows, uc, up = sc.process(records, ch)
    blocked = _block_if_missing_reps(rows, ch, token)
    if blocked:
        return blocked
    bio = sc.generate_workbook(rows, ch, job["path"])
    today = datetime.datetime.now().strftime("%Y%m%d")
    fname = "%s_스마트오더_부착양식_%s.xlsx" % (cfg["name"], today)
    return send_file(bio, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/labelout/<ch>/<token>")
def labelout(ch, token):
    """폼텍 디자인프로 데이터 원본용: RAW 원본 컬럼 + 배송센터(이고) + 상품명(대표) 시트만 담은 xlsx."""
    if ch not in sc.CHANNELS:
        abort(404)
    job = _load_job(ch, token)
    if not job:
        flash("업로드 세션이 만료되었습니다.", "error")
        return redirect(url_for("upload", ch=ch))
    cfg = sc.CHANNELS[ch]
    records = sc.parse_raw(job["path"], ch)
    rows, uc, up = sc.process(records, ch)
    blocked = _block_if_missing_reps(rows, ch, token)
    if blocked:
        return blocked
    bio = sc.generate_labelout_workbook(job["path"], ch)
    today = datetime.datetime.now().strftime("%Y%m%d")
    fname = "%s_스마트오더_라벨출력(폼텍)_%s.xlsx" % (cfg["name"], today)
    return send_file(bio, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------- 업로드 이력(누적 파일) ----------------
@app.route("/history")
def history():
    ch = request.args.get("ch", "")
    items = list(reversed(_load_index()))
    if ch in sc.CHANNELS:
        items = [r for r in items if r["ch"] == ch]
    counts = {}
    for r in _load_index():
        counts[r["ch"]] = counts.get(r["ch"], 0) + 1
    return render_template("history.html", items=items, cfg=sc.CHANNELS,
                           channels=CH_ORDER, sel=ch, counts=counts)


@app.route("/history/raw/<uid>")
def history_raw(uid):
    rec = _rec_by_id(uid)
    if not rec:
        abort(404)
    path = _rec_path(rec)
    if not os.path.exists(path):
        abort(404)
    return send_file(path, as_attachment=True, download_name=rec["orig"])


@app.route("/history/delete/<uid>", methods=["POST"])
def history_delete(uid):
    idx = _load_index()
    rec = next((r for r in idx if r["id"] == uid), None)
    if rec:
        try:
            os.remove(_rec_path(rec))
        except Exception:
            pass
        idx = [r for r in idx if r["id"] != uid]
        _save_index(idx)
        flash("삭제했습니다: %s" % rec["orig"], "ok")
    return redirect(url_for("history", ch=request.form.get("ch", "")))


# ---------------- 마스터 관리 ----------------
MASTER_DEFS = {
    "center":  {"file": sc.CENTER_MASTER, "title": "센터 (센터명→이고센터) · 전 채널 공통", "kind": "center"},
    "product": {"file": "product.json",   "title": "상품 (한익스상품명 ← 채널별 상품코드)", "kind": "product"},
}

@app.route("/masters")
def masters():
    counts = {k: len(sc.load_master(v["file"])) for k, v in MASTER_DEFS.items()}
    return render_template("masters.html", defs=MASTER_DEFS, counts=counts)


@app.route("/masters/<which>", methods=["GET"])
def master_edit(which):
    if which not in MASTER_DEFS:
        abort(404)
    d = MASTER_DEFS[which]
    data = sc.load_master(d["file"])
    q = request.args.get("q", "").strip()
    reps = []
    # 상품 마스터: 같은 상품이 채널마다 코드가 달라 → 한익스상품명으로 묶어서 보여준다
    if d["kind"] == "product":
        groups = {}
        for code in sorted(data):
            rep, cat = sc.product_rep_cat(data, code)
            g = groups.setdefault(rep or "", {"rep": rep or "", "cat": cat or sc.DEFAULT_CAT, "codes": []})
            g["codes"].append(code)
        reps = sorted(groups)
        items = sorted(groups.values(), key=lambda g: g["rep"])
        if q:
            ql = q.lower()
            items = [g for g in items
                     if ql in g["rep"].lower() or any(ql in c.lower() for c in g["codes"])]
    else:
        items = sorted(data.items())
        if q:
            items = [(k, v) for k, v in items if q.lower() in k.lower() or q.lower() in str(v).lower()]
    # 한 화면에 다 뿌리면 찾기 힘들어서 페이지로 나눈다
    per, found = 25, len(items)
    pages = max(1, (found + per - 1) // per)
    page = min(max(request.args.get("page", 1, type=int), 1), pages)
    items = items[(page - 1) * per:page * per]
    hubs = sc.all_hubs()
    return render_template("master_edit.html", which=which, d=d, items=items, reps=reps,
                           q=q, total=len(data), found=found, page=page, pages=pages,
                           hubs=hubs, product_cats=sc.PRODUCT_CATS)


@app.route("/masters/<which>/save", methods=["POST"])
def master_save(which):
    if which not in MASTER_DEFS:
        abort(404)
    d = MASTER_DEFS[which]
    data = sc.load_master(d["file"])
    is_prod = d["kind"] == "product"
    act = request.form.get("action")
    if act == "add":
        k = request.form.get("key", "").strip()
        v = request.form.get("value", "").strip()
        if k and v:
            if is_prod:
                cat = request.form.get("cat", sc.DEFAULT_CAT).strip() or sc.DEFAULT_CAT
                data[k] = {"rep": v, "cat": cat}
                flash("추가: %s → %s (%s)" % (k, v, cat), "ok")
            else:
                data[k] = v
                flash("추가: %s → %s" % (k, v), "ok")
    elif act == "update":
        k = request.form.get("key", "").strip()
        v = request.form.get("value", "").strip()
        if k in data and v:
            if is_prod:
                cat = request.form.get("cat", sc.DEFAULT_CAT).strip() or sc.DEFAULT_CAT
                data[k] = {"rep": v, "cat": cat}
                flash("수정: %s → %s (%s)" % (k, v, cat), "ok")
            else:
                data[k] = v
                flash("수정: %s → %s" % (k, v), "ok")
    elif act == "update_group" and is_prod:
        # 한익스상품명/구분 일괄 수정 — 그 상품의 채널별 코드 전부에 적용
        old = request.form.get("old_rep", "").strip()
        v = request.form.get("value", "").strip()
        cat = request.form.get("cat", sc.DEFAULT_CAT).strip() or sc.DEFAULT_CAT
        codes = [c for c in data if (sc.product_rep_cat(data, c)[0] or "") == old]
        if v and codes:
            for c in codes:
                data[c] = {"rep": v, "cat": cat}
            flash("수정: %s → %s (%s) · 코드 %d개 적용" % (old, v, cat, len(codes)), "ok")
    elif act == "delete_group" and is_prod:
        old = request.form.get("old_rep", "").strip()
        codes = [c for c in data if (sc.product_rep_cat(data, c)[0] or "") == old]
        for c in codes:
            data.pop(c)
        if codes:
            flash("삭제: %s (코드 %d개)" % (old, len(codes)), "ok")
    elif act == "delete":
        k = request.form.get("key", "").strip()
        if k in data:
            data.pop(k)
            flash("삭제: %s" % k, "ok")
    sc.save_master(d["file"], data)
    return redirect(url_for("master_edit", which=which, q=request.form.get("q", "")))


# ---------- 마스터 엑셀 (다운로드 → 엑셀에서 수정 → 업로드) ----------
MASTER_TMP = os.path.join(ARCHIVE_DIR, "master_upload")
MASTER_BAK = os.path.join(ARCHIVE_DIR, "master_backup")


@app.route("/masters/<which>/xlsx")
def master_xlsx(which):
    if which not in sc.MASTER_KINDS:
        abort(404)
    bio = sc.build_master_xlsx(which)
    today = datetime.datetime.now().strftime("%Y%m%d")
    nm = {"center": "센터마스터", "product": "상품마스터"}[which]
    return send_file(bio, as_attachment=True, download_name="%s_%s.xlsx" % (nm, today),
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/masters/<which>/import", methods=["POST"])
def master_import(which):
    """엑셀 업로드 → 무엇이 바뀌는지 먼저 보여준다(아직 반영 안 함)."""
    if which not in sc.MASTER_KINDS:
        abort(404)
    f = request.files.get("file")
    if not f or not f.filename:
        flash("엑셀 파일을 선택해 주세요.", "error")
        return redirect(url_for("master_edit", which=which))
    if os.path.splitext(f.filename)[1].lower() not in (".xlsx", ".xlsm"):
        flash("xlsx 파일만 업로드할 수 있습니다.", "error")
        return redirect(url_for("master_edit", which=which))
    os.makedirs(MASTER_TMP, exist_ok=True)
    token = "%s_%s" % (which, uuid.uuid4().hex[:8])
    path = os.path.join(MASTER_TMP, token + ".xlsx")
    f.save(path)
    try:
        new, errors = sc.parse_master_xlsx(path, which)
    except Exception as e:
        os.remove(path)
        flash("엑셀을 읽지 못했습니다: %s" % e, "error")
        return redirect(url_for("master_edit", which=which))
    if not new and not errors:
        os.remove(path)
        flash("내용이 없는 엑셀입니다.", "error")
        return redirect(url_for("master_edit", which=which))
    diff = sc.diff_master(which, new)
    return render_template("master_import.html", which=which, d=MASTER_DEFS[which],
                           token=token, filename=f.filename, diff=diff,
                           errors=errors, total=len(new))


@app.route("/masters/<which>/import/apply", methods=["POST"])
def master_import_apply(which):
    """확인 후 실제 반영. 반영 직전 현재 마스터를 백업한다."""
    if which not in sc.MASTER_KINDS:
        abort(404)
    token = request.form.get("token", "")
    path = os.path.join(MASTER_TMP, token + ".xlsx")
    if not token.startswith(which + "_") or not os.path.exists(path):
        flash("업로드가 만료되었습니다. 엑셀을 다시 올려주세요.", "error")
        return redirect(url_for("master_edit", which=which))
    new, errors = sc.parse_master_xlsx(path, which)
    fname = sc.MASTER_KINDS[which]["file"]
    # 백업 후 교체
    os.makedirs(MASTER_BAK, exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    with open(os.path.join(MASTER_BAK, "%s.%s.json" % (fname, ts)), "w", encoding="utf-8") as bf:
        json.dump(sc.load_master(fname), bf, ensure_ascii=False, indent=1, sort_keys=True)
    diff = sc.diff_master(which, new)
    sc.save_master(fname, new)
    os.remove(path)
    flash("엑셀을 반영했습니다 — 추가 %d · 수정 %d · 삭제 %d (전체 %d개). 이전 마스터는 백업해 두었습니다."
          % (len(diff["added"]), len(diff["changed"]), len(diff["removed"]), len(new)), "ok")
    return redirect(url_for("master_edit", which=which))


@app.route("/health")
def health():
    return jsonify(ok=True)


if __name__ == "__main__":
    import webbrowser, threading
    port = int(os.environ.get("PORT", 5057))
    if os.environ.get("NO_BROWSER") != "1":
        threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:%d/" % port)).start()
    app.run(host="0.0.0.0", port=port, debug=False)
