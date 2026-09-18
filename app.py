"""DocMeta Web - view and carefully edit PDF / Word document metadata.

Single-file Flask app. PDF metadata is handled with the bundled ExifTool,
Word (.docx) metadata with python-docx. Built into one Windows exe with
PyInstaller (see .github/workflows/build-windows.yml).
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import uuid
from datetime import datetime

from flask import Flask, request, redirect, url_for, render_template, send_file, jsonify

try:
    import version
    APP_VERSION = version.APP_VERSION
except Exception:
    APP_VERSION = "dev"

app = Flask(__name__)

UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "docmeta_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

ALLOWED = {".pdf", ".docx"}


# ----------------------------------------------------------------------------
# helpers
# ----------------------------------------------------------------------------
def resource_path(rel):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def find_exiftool():
    """Locate the exiftool binary: bundled next to the frozen exe, else PATH."""
    candidates = [
        os.path.join(resource_path("."), "exiftool.exe"),
        os.path.join(os.path.dirname(sys.executable), "exiftool.exe"),
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    found = shutil.which("exiftool")
    if found:
        return found
    return "exiftool"


def run_exiftool(args, cwd=None):
    exe = find_exiftool()
    # The Windows exiftool.exe needs its exiftool_files/ sibling visible
    # from the working directory.
    workdir = cwd
    if workdir is None and getattr(sys, "frozen", False):
        staged = resource_path("exiftool_files")
        if os.path.isdir(staged):
            workdir = resource_path(".")
    p = subprocess.run(
        [exe] + args,
        capture_output=True,
        text=True,
        cwd=workdir,
    )
    return p


def doc_dir(uid):
    return os.path.join(UPLOAD_DIR, uid)


def doc_path(uid):
    d = doc_dir(uid)
    if not os.path.isdir(d):
        return None
    for f in os.listdir(d):
        return os.path.join(d, f)
    return None


def file_kind(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return "pdf"
    if ext == ".docx":
        return "docx"
    return None


# ----------------------------------------------------------------------------
# PDF metadata via ExifTool
# ----------------------------------------------------------------------------
# (friendly label, PDF Info tag, XMP tag, kind)
PDF_FIELDS = [
    ("Author", "PDF:Author", "XMP-dc:Creator", "text"),
    ("Creator (app)", "PDF:Creator", "XMP-xmp:CreatorTool", "text"),
    ("Producer", "PDF:Producer", "XMP-pdf:Producer", "text"),
    ("Created", "PDF:CreateDate", "XMP-xmp:CreateDate", "date"),
    ("Last modified", "PDF:ModDate", "XMP-xmp:ModifyDate", "date"),
    ("Title", "PDF:Title", "XMP-dc:Title", "text"),
    ("Subject", "PDF:Subject", "XMP-dc:Description", "text"),
    ("Keywords", "PDF:Keywords", "XMP-dc:Subject", "text"),
]


def read_pdf_meta(path):
    p = run_exiftool(["-j", "-G", "-a", path])
    if p.returncode != 0:
        return None, p.stderr.strip() or "ExifTool failed to read the file."
    try:
        info = json.loads(p.stdout)[0]
    except Exception:
        return None, "Could not parse ExifTool output."
    fields = []
    for label, pdf_tag, xmp_tag, kind in PDF_FIELDS:
        val = info.get(pdf_tag)
        if val is None:
            val = info.get(xmp_tag)
        if isinstance(val, dict):  # lang-alt, e.g. {'x-default': '...'}
            val = val.get("x-default") or next(iter(val.values()), "")
        if isinstance(val, list):
            val = ", ".join(str(v) for v in val)
        fields.append({"label": label, "key": pdf_tag, "kind": kind,
                       "value": "" if val is None else str(val)})
    # raw table: every tag for the curious
    raw = []
    for k in sorted(info.keys()):
        if k in ("SourceFile", "ExifToolVersion"):
            continue
        v = info[k]
        if isinstance(v, (dict, list)):
            v = json.dumps(v, ensure_ascii=False)
        raw.append((k, str(v)))
    return {"fields": fields, "raw": raw}, None


def exiftool_date(dt_local):
    """'2020-01-15T10:30' -> '2020:01:15 10:30:00' for ExifTool."""
    dt = datetime.strptime(dt_local, "%Y-%m-%dT%H:%M")
    return dt.strftime("%Y:%m:%d %H:%M:%S")


def write_pdf_meta(path, form):
    args = ["-overwrite_original"]
    for label, pdf_tag, xmp_tag, kind in PDF_FIELDS:
        if kind == "date":
            raw = form.get("date:" + pdf_tag, "").strip()
            if not raw:
                continue  # empty date = leave unchanged
            val = exiftool_date(raw)
        else:
            val = form.get("text:" + pdf_tag, "")
            if val is None:
                continue
            val = val.strip()
            # empty text = delete the tag (both Info dict and XMP copies)
            if val == "":
                args.append(f"-{pdf_tag}=")
                args.append(f"-{xmp_tag}=")
                continue
        args.append(f"-{pdf_tag}={val}")
        args.append(f"-{xmp_tag}={val}")
    args.append(path)
    p = run_exiftool(args)
    if p.returncode != 0:
        return p.stderr.strip() or "ExifTool failed to write."
    return None


def clear_pdf_meta(path):
    args = ["-overwrite_original"]
    for _label, pdf_tag, xmp_tag, kind in PDF_FIELDS:
        if kind == "date":
            continue  # keep dates; they are edited individually
        args.append(f"-{pdf_tag}=")
        args.append(f"-{xmp_tag}=")
    args.append(path)
    p = run_exiftool(args)
    if p.returncode != 0:
        return p.stderr.strip() or "ExifTool failed."
    return None


# ----------------------------------------------------------------------------
# DOCX metadata via python-docx
# ----------------------------------------------------------------------------
# (friendly label, core_properties attr, kind)
DOCX_FIELDS = [
    ("Author", "author", "text"),
    ("Last modified by", "last_modified_by", "text"),
    ("Created", "created", "date"),
    ("Modified", "modified", "date"),
    ("Title", "title", "text"),
    ("Subject", "subject", "text"),
    ("Keywords", "keywords", "text"),
    ("Category", "category", "text"),
    ("Comments", "comments", "text"),
]


def _docx():
    from docx import Document
    return Document


def read_docx_meta(path):
    try:
        doc = _docx()(path)
    except Exception as e:
        return None, f"Could not open Word document: {e}"
    cp = doc.core_properties
    fields = []
    for label, attr, kind in DOCX_FIELDS:
        val = getattr(cp, attr, None)
        if isinstance(val, datetime):
            val = val.strftime("%Y-%m-%dT%H:%M")
        fields.append({"label": label, "key": attr, "kind": kind,
                       "value": "" if val is None else str(val)})
    return {"fields": fields, "raw": []}, None


def _parse_local_dt(raw):
    dt = datetime.strptime(raw, "%Y-%m-%dT%H:%M")
    # treat the entered time as this computer's local time
    return dt.astimezone()


def write_docx_meta(path, form):
    try:
        doc = _docx()(path)
    except Exception as e:
        return f"Could not open Word document: {e}"
    cp = doc.core_properties
    try:
        for label, attr, kind in DOCX_FIELDS:
            if kind == "date":
                raw = form.get("date:" + attr, "").strip()
                if not raw:
                    continue  # empty date = leave unchanged
                setattr(cp, attr, _parse_local_dt(raw))
            else:
                val = form.get("text:" + attr, "")
                if val is None:
                    continue
                setattr(cp, attr, val.strip())
        doc.save(path)
    except Exception as e:
        return f"Could not save Word document: {e}"
    return None


def clear_docx_meta(path):
    try:
        doc = _docx()(path)
    except Exception as e:
        return f"Could not open Word document: {e}"
    cp = doc.core_properties
    try:
        for _label, attr, kind in DOCX_FIELDS:
            if kind == "date":
                continue
            setattr(cp, attr, "")
        doc.save(path)
    except Exception as e:
        return f"Could not save Word document: {e}"
    return None


# ----------------------------------------------------------------------------
# routes
# ----------------------------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html", app_version=APP_VERSION)


@app.route("/upload", methods=["POST"])
def upload():
    f = request.files.get("file")
    if not f or not f.filename:
        return redirect(url_for("index"))
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ALLOWED:
        return render_template("index.html", app_version=APP_VERSION,
                               error="Only .pdf and .docx files are supported."), 400
    uid = uuid.uuid4().hex[:12]
    os.makedirs(doc_dir(uid), exist_ok=True)
    # keep the original filename for the download at the end
    safe_name = os.path.basename(f.filename)
    f.save(os.path.join(doc_dir(uid), safe_name))
    return redirect(url_for("doc", uid=uid))


def _read_meta(path, kind):
    if kind == "pdf":
        return read_pdf_meta(path)
    return read_docx_meta(path)


@app.route("/doc/<uid>")
def doc(uid):
    path = doc_path(uid)
    if not path:
        return redirect(url_for("index"))
    kind = file_kind(path)
    meta, err = _read_meta(path, kind)
    if err:
        return render_template("index.html", app_version=APP_VERSION, error=err), 400
    return render_template(
        "doc.html",
        app_version=APP_VERSION,
        uid=uid,
        filename=os.path.basename(path),
        kind=kind,
        fields=meta["fields"],
        raw=meta["raw"],
        saved=request.args.get("saved"),
        cleared=request.args.get("cleared"),
        error=request.args.get("error"),
    )


@app.route("/doc/<uid>/save", methods=["POST"])
def save(uid):
    path = doc_path(uid)
    if not path:
        return redirect(url_for("index"))
    kind = file_kind(path)
    err = write_pdf_meta(path, request.form) if kind == "pdf" else write_docx_meta(path, request.form)
    if err:
        return redirect(url_for("doc", uid=uid, error=err))
    return redirect(url_for("doc", uid=uid, saved=1))


@app.route("/doc/<uid>/clear", methods=["POST"])
def clear(uid):
    path = doc_path(uid)
    if not path:
        return redirect(url_for("index"))
    kind = file_kind(path)
    err = clear_pdf_meta(path) if kind == "pdf" else clear_docx_meta(path)
    if err:
        return redirect(url_for("doc", uid=uid, error=err))
    return redirect(url_for("doc", uid=uid, cleared=1))


@app.route("/doc/<uid>/download")
def download(uid):
    path = doc_path(uid)
    if not path:
        return redirect(url_for("index"))
    return send_file(path, as_attachment=True,
                     download_name=os.path.basename(path))


@app.route("/api/health")
def health():
    return jsonify({"ok": True, "version": APP_VERSION,
                    "exiftool": find_exiftool()})


def main():
    import webbrowser
    from threading import Timer
    if getattr(sys, "frozen", False):
        Timer(1.0, lambda: webbrowser.open("http://127.0.0.1:5000")).start()
    app.run(host="127.0.0.1", port=5000, debug=False)


if __name__ == "__main__":
    main()
