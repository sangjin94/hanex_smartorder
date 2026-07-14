# -*- coding: utf-8 -*-
"""한익스프레스 스마트오더 라벨 생성기 (웹)

기능
 - 판매채널(CU/GS/E24)별 RAW 업로드 → 코드 기반 매핑 → 부착양식(A4) 라벨 xlsx 생성
 - 업로드 중 미매핑 센터/상품 코드를 화면에서 즉시 마스터에 등록
 - 마스터 관리(센터코드→거점센터, 상품코드→대표상품명) 조회/추가/수정/삭제
"""
import os, io, uuid, datetime, tempfile
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

UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "hanex_smartorder_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)
JOBS = {}   # token -> {channel, path, filename, created}

CH_ORDER = ["cu", "gs", "e24"]

def _cleanup():
    now = datetime.datetime.now()
    for tok, job in list(JOBS.items()):
        if (now - job["created"]).total_seconds() > 6 * 3600:
            try: os.remove(job["path"])
            except Exception: pass
            JOBS.pop(tok, None)


@app.route("/")
def index():
    stats = {}
    for ch in CH_ORDER:
        cfg = sc.CHANNELS[ch]
        stats[ch] = {
            "name": cfg["name"],
            "centers": len(sc.load_master(cfg["center_master"])),
        }
    products = len(sc.load_master("product.json"))
    return render_template("index.html", channels=CH_ORDER, cfg=sc.CHANNELS,
                           stats=stats, products=products)


@app.route("/u/<ch>", methods=["GET", "POST"])
def upload(ch):
    if ch not in sc.CHANNELS:
        abort(404)
    cfg = sc.CHANNELS[ch]
    if request.method == "GET":
        return render_template("upload.html", ch=ch, cfg=cfg)

    f = request.files.get("file")
    if not f or not f.filename:
        flash("파일을 선택해 주세요.", "error")
        return redirect(url_for("upload", ch=ch))
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".xlsx", ".xls"):
        flash("xlsx 또는 xls 파일만 업로드할 수 있습니다.", "error")
        return redirect(url_for("upload", ch=ch))

    _cleanup()
    token = uuid.uuid4().hex
    path = os.path.join(UPLOAD_DIR, token + ext)
    f.save(path)
    JOBS[token] = {"channel": ch, "path": path, "filename": f.filename,
                   "created": datetime.datetime.now()}
    return redirect(url_for("result", ch=ch, token=token))


def _load_job(ch, token):
    job = JOBS.get(token)
    if not job or job["channel"] != ch or not os.path.exists(job["path"]):
        return None
    return job


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
    return render_template("result.html", ch=ch, cfg=cfg, token=token,
                           filename=job["filename"], nrec=len(records),
                           nrows=len(rows), total_qty=total_qty,
                           unmapped_centers=uc, unmapped_products=up,
                           hubs=hubs, center_key=cfg["center_key"],
                           prod_list=prod_list, hub_list=hub_list)


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
    rows = sc.sort_rows(rows)
    return render_template("print.html", ch=ch, cfg=cfg, rows=rows, token=token)


@app.route("/assign/<ch>/<token>", methods=["POST"])
def assign(ch, token):
    if ch not in sc.CHANNELS:
        abort(404)
    cfg = sc.CHANNELS[ch]
    # 센터 매핑 저장
    cm = sc.load_master(cfg["center_master"])
    for key in request.form:
        if key.startswith("center::"):
            code = key[len("center::"):]
            val = request.form.get(key, "").strip()
            if val:
                cm[code] = val
    sc.save_master(cfg["center_master"], cm)
    # 상품 매핑 저장 (공통 마스터)
    pm = sc.load_master("product.json")
    for key in request.form:
        if key.startswith("prod::"):
            code = key[len("prod::"):]
            val = request.form.get(key, "").strip()
            if val:
                pm[code] = val
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
    bio = sc.generate_workbook(rows, ch)
    today = datetime.datetime.now().strftime("%Y%m%d")
    fname = "%s_스마트오더_부착양식_%s.xlsx" % (cfg["name"], today)
    return send_file(bio, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------- 마스터 관리 ----------------
MASTER_DEFS = {
    "center_cu":  {"file": "center_cu.json",  "title": "CU 센터 (센터코드→거점센터)", "kind": "center"},
    "center_gs":  {"file": "center_gs.json",  "title": "GS 센터 (센터코드→거점센터)", "kind": "center"},
    "center_e24": {"file": "center_e24.json", "title": "이마트24 센터 (입고센터명→거점센터)", "kind": "center"},
    "product":    {"file": "product.json",    "title": "상품 (상품코드→대표=한익스상품명)", "kind": "product"},
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
    items = sorted(data.items())
    if q:
        items = [(k, v) for k, v in items if q.lower() in k.lower() or q.lower() in str(v).lower()]
    hubs = sc.all_hubs()
    return render_template("master_edit.html", which=which, d=d, items=items,
                           q=q, total=len(data), hubs=hubs)


@app.route("/masters/<which>/save", methods=["POST"])
def master_save(which):
    if which not in MASTER_DEFS:
        abort(404)
    d = MASTER_DEFS[which]
    data = sc.load_master(d["file"])
    act = request.form.get("action")
    if act == "add":
        k = request.form.get("key", "").strip()
        v = request.form.get("value", "").strip()
        if k and v:
            data[k] = v
            flash("추가: %s → %s" % (k, v), "ok")
    elif act == "update":
        k = request.form.get("key", "").strip()
        v = request.form.get("value", "").strip()
        if k in data and v:
            data[k] = v
            flash("수정: %s → %s" % (k, v), "ok")
    elif act == "delete":
        k = request.form.get("key", "").strip()
        if k in data:
            data.pop(k)
            flash("삭제: %s" % k, "ok")
    sc.save_master(d["file"], data)
    return redirect(url_for("master_edit", which=which, q=request.form.get("q", "")))


@app.route("/health")
def health():
    return jsonify(ok=True)


if __name__ == "__main__":
    import webbrowser, threading
    port = int(os.environ.get("PORT", 5057))
    if os.environ.get("NO_BROWSER") != "1":
        threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:%d/" % port)).start()
    app.run(host="0.0.0.0", port=port, debug=False)
