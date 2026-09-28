"""DocMeta Web - view and carefully edit PDF / Word document metadata.

Single-file Flask app. PDF metadata is handled with the bundled ExifTool,
Word (.docx) metadata with python-docx. Built into one Windows exe with
PyInstaller (see .github/workflows/build-windows.yml).
"""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import uuid
from datetime import datetime, timezone

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
# pdf_tag may be None when only an XMP tag exists (e.g. Device).
PDF_FIELDS = [
    ("Author", "PDF:Author", "XMP-dc:Creator", "text"),
    ("Creator (app)", "PDF:Creator", "XMP-xmp:CreatorTool", "text"),
    ("Producer", "PDF:Producer", "XMP-pdf:Producer", "text"),
    ("Device", None, "XMP-dc:Source", "text"),
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
        key = pdf_tag or xmp_tag
        val = info.get(pdf_tag) if pdf_tag else None
        if val is None and xmp_tag:
            val = info.get(xmp_tag)
        if isinstance(val, dict):  # lang-alt, e.g. {'x-default': '...'}
            val = val.get("x-default") or next(iter(val.values()), "")
        if isinstance(val, list):
            # dates: show the first; text: join like before
            val = val[0] if (kind == "date" and val) else ", ".join(str(v) for v in val)
        if kind == "date":
            # Render in datetime-local format so the edit box shows the
            # current value (ExifTool's "YYYY:MM:DD HH:MM:SS" is invalid
            # for that input and would display as empty).
            val = exiftool_to_input(val)
        fields.append({"label": label, "key": key, "kind": kind,
                       "value": "" if val is None else str(val)})
    # raw table: every tag for the curious
    raw = []
    for k in sorted(info.keys()):
        if k in ("SourceFile", "ExifTool:ExifToolVersion"):
            continue
        v = info[k]
        if isinstance(v, (dict, list)):
            v = json.dumps(v, ensure_ascii=False)
        raw.append((k, str(v)))
    return {"fields": fields, "raw": raw}, None


def exiftool_to_input(val):
    """'2024:01:15 10:30:00' (or with a timezone suffix) -> '2024-01-15T10:30'
    for <input type="datetime-local">. Returns '' if unparseable."""
    if not val:
        return ""
    m = re.match(r"(\d{4}):(\d{2}):(\d{2})[ T](\d{2}):(\d{2})", str(val))
    if not m:
        return ""
    return f"{m.group(1)}-{m.group(2)}-{m.group(3)}T{m.group(4)}:{m.group(5)}"


def parse_date_input(raw, field_label):
    """Parse a datetime-local style value flexibly.

    Accepts 'YYYY-MM-DDTHH:MM', 'YYYY-MM-DDTHH:MM:SS', or a bare
    'YYYY-MM-DD' (midnight). Raises ValueError with a friendly message
    naming the field when nothing matches.
    """
    raw = (raw or "").strip()
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            pass
    raise ValueError(
        f"Couldn't understand the date for '{field_label}'. "
        "Pick it from the calendar popup (date and time)."
    )


def exiftool_date(dt):
    """datetime -> 'YYYY:MM:DD HH:MM:SS' for ExifTool."""
    return dt.strftime("%Y:%m:%d %H:%M:%S")


def write_pdf_meta(path, form):
    before = _read_producer_tags(path)
    args = ["-overwrite_original"]
    for label, pdf_tag, xmp_tag, kind in PDF_FIELDS:
        key = pdf_tag or xmp_tag
        tags = [t for t in (pdf_tag, xmp_tag) if t]
        if kind == "date":
            raw = form.get("date:" + key, "").strip()
            if not raw:
                continue  # empty date = leave unchanged
            try:
                val = exiftool_date(parse_date_input(raw, label))
            except ValueError as e:
                return str(e)
        else:
            val = form.get("text:" + key, "")
            if val is None:
                continue
            val = val.strip()
            # empty text = delete the tag (both Info dict and XMP copies)
        for t in tags:
            args.append(f"-{t}={val}")
    args.append(path)
    p = run_exiftool(args)
    if p.returncode != 0:
        return p.stderr.strip() or "ExifTool failed to write."
    _scrub_tool_signature(path, form, before)
    return None


_TOOL_SIG = re.compile(r"image::exiftool", re.IGNORECASE)
_PRODUCER_TAGS = ("PDF:Producer", "XMP-pdf:Producer")


def _read_producer_tags(path):
    """Current {tag: value} for the Producer tags ('' when absent)."""
    p = run_exiftool(["-j", "-G", *_PRODUCER_TAGS, path])
    try:
        info = json.loads(p.stdout)[0]
    except Exception:
        return {t: "" for t in _PRODUCER_TAGS}
    out = {}
    for t in _PRODUCER_TAGS:
        v = info.get(t)
        if isinstance(v, list):
            v = v[0] if v else ""
        out[t] = "" if v is None else str(v)
    return out


def _scrub_tool_signature(path, form, before):
    """Keep ExifTool's own Producer signature out of the file.

    If a Producer tag now contains ExifTool's signature, restore the value
    the user typed (when ExifTool clobbered it) or delete the tag (when the
    signature was already there or the field was cleared). The only
    exception: the user deliberately typed the signature itself.
    """
    user_val = (form.get("text:PDF:Producer") or "").strip()
    user_typed_sig = bool(user_val and _TOOL_SIG.search(user_val)
                          and before.get("PDF:Producer") != user_val)
    if user_typed_sig:
        return
    cur = _read_producer_tags(path)
    fix = []
    for tag in _PRODUCER_TAGS:
        if cur.get(tag) and _TOOL_SIG.search(cur[tag]):
            if user_val and not _TOOL_SIG.search(user_val):
                fix.append(f"-{tag}={user_val}")
            else:
                fix.append(f"-{tag}=")
    if fix:
        run_exiftool(["-overwrite_original", *fix, path])


def clear_pdf_meta(path):
    args = ["-overwrite_original"]
    for _label, pdf_tag, xmp_tag, kind in PDF_FIELDS:
        if kind == "date":
            continue  # keep dates; they are edited individually
        if pdf_tag:
            args.append(f"-{pdf_tag}=")
        if xmp_tag:
            args.append(f"-{xmp_tag}=")
    args.append(path)
    p = run_exiftool(args)
    if p.returncode != 0:
        return p.stderr.strip() or "ExifTool failed."
    _scrub_tool_signature(path, {}, {})
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


def _parse_local_dt(raw, label):
    # treat the entered time as this computer's local time
    return parse_date_input(raw, label).astimezone()


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
                try:
                    setattr(cp, attr, _parse_local_dt(raw, label))
                except ValueError as e:
                    return str(e)
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
# "Save to PC": stamp Windows filesystem dates
# ----------------------------------------------------------------------------
# Explorer's Properties reads Created/Modified from the file itself, not from
# embedded metadata — so a browser download can never carry her dates over
# (the browser always stamps "now"). Writing the finished file to her
# Downloads folder ourselves lets us stamp the dates she chose.
def _dt_to_filetime(dt):
    """datetime -> Windows FILETIME (100ns ticks since 1601-01-01 UTC)."""
    aware = dt.astimezone() if dt.tzinfo is None else dt
    utc = aware.astimezone(timezone.utc)
    epoch = datetime(1601, 1, 1, tzinfo=timezone.utc)
    ft = int((utc - epoch).total_seconds() * 10_000_000)
    return ft


def set_windows_file_times(path, created=None, modified=None):
    """Stamp a file's Created / Last-modified dates (Windows only).

    created/modified: datetimes (naive = this computer's local time).
    Her date always wins: callers pass her chosen date, never the original.
    Returns an error string, or None on success / non-Windows.
    """
    if os.name != "nt":
        return None  # not Windows: nothing to stamp
    if created is None and modified is None:
        return None
    try:
        import ctypes
        from ctypes import wintypes

        kernel32 = ctypes.windll.kernel32
        FILE_WRITE_ATTRIBUTES = 0x100
        OPEN_EXISTING = 3
        h = kernel32.CreateFileW(os.fspath(path), FILE_WRITE_ATTRIBUTES, 0,
                                 None, OPEN_EXISTING, 0, None)
        if h == wintypes.HANDLE(-1).value:  # INVALID_HANDLE_VALUE
            return "Could not open the file to stamp its dates."
        try:
            c = wintypes.FILETIME.from_buffer_copy(
                _dt_to_filetime(created).to_bytes(8, "little")) if created else None
            m = wintypes.FILETIME.from_buffer_copy(
                _dt_to_filetime(modified).to_bytes(8, "little")) if modified else None
            ok = kernel32.SetFileTime(h,
                                      ctypes.byref(c) if c else None,
                                      None,  # leave last-access alone
                                      ctypes.byref(m) if m else None)
            if not ok:
                return "Windows refused to set the file dates."
        finally:
            kernel32.CloseHandle(h)
    except Exception as e:
        return f"Could not stamp the Windows file dates: {e}"
    return None


def unique_download_path(filename):
    """A non-clobbering path inside the user's Downloads folder."""
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(downloads, exist_ok=True)
    base, ext = os.path.splitext(filename)
    candidate = os.path.join(downloads, filename)
    n = 1
    while os.path.exists(candidate):
        n += 1
        candidate = os.path.join(downloads, f"{base} ({n}){ext}")
    return candidate


def _optional_dt(raw, label):
    """datetime-local value -> aware local datetime, or None if empty."""
    raw = (raw or "").strip()
    if not raw:
        return None
    try:
        return parse_date_input(raw, label).astimezone()
    except ValueError:
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
        saved_to=request.args.get("saved_to"),
        error=request.args.get("error"),
    )


def _write_meta(path, kind, form):
    if kind == "pdf":
        return write_pdf_meta(path, form)
    return write_docx_meta(path, form)


@app.route("/doc/<uid>/save_to_pc", methods=["POST"])
def save_to_pc(uid):
    """Save metadata, then write the finished file straight to Downloads
    with her dates stamped on it, so Explorer Properties shows them.

    A browser download always stamps "now" as the file's Created date;
    writing the file ourselves is the only way her chosen date survives.
    Her date always wins: we stamp what she typed, never the original.
    """
    path = doc_path(uid)
    if not path:
        return redirect(url_for("index"))
    kind = file_kind(path)
    form = request.form
    err = _write_meta(path, kind, form)
    if err:
        return redirect(url_for("doc", uid=uid, error=err))
    now = datetime.now().astimezone()
    if kind == "pdf":
        created = _optional_dt(form.get("date:PDF:CreateDate", ""), "Created") or now
        modified = _optional_dt(form.get("date:PDF:ModDate", ""), "Last modified") or created
    else:
        created = _optional_dt(form.get("date:created", ""), "Created") or now
        modified = _optional_dt(form.get("date:modified", ""), "Modified") or created
    dest = unique_download_path(os.path.basename(path))
    try:
        shutil.copy2(path, dest)
    except Exception as e:
        return redirect(url_for("doc", uid=uid,
                                error=f"Could not save to Downloads: {e}"))
    err = set_windows_file_times(dest, created, modified)
    if err:
        return redirect(url_for("doc", uid=uid, error=err))
    msg = (f"Saved to Downloads as {os.path.basename(dest)} — "
           f"Properties will show Created: {created.strftime('%Y-%m-%d %H:%M')}")
    return redirect(url_for("doc", uid=uid, saved_to=msg))


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
