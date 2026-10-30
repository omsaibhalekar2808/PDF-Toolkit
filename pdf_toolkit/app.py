import io
import os
import shutil
import tempfile
import zipfile
import logging
import warnings
from pathlib import Path
from flask import Flask, jsonify, render_template, request, send_file
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from werkzeug.utils import secure_filename
_DLL_HANDLE = None

if os.name == "nt":
    tesseract_dir = Path(r"C:\Program Files\Tesseract-OCR")

    if tesseract_dir.is_dir():
        _DLL_HANDLE = os.add_dll_directory(str(tesseract_dir))
import pdf_tools as T

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_CONTENT_LENGTH", 25 * 1024 * 1024))
app.config["MAX_FORM_MEMORY_SIZE"] = 2 * 1024 * 1024
app.config["MAX_FORM_PARTS"] = 100
app.config["JSON_SORT_KEYS"] = False
warnings.filterwarnings("error", category=Warning, module="PIL")
logging.basicConfig(level=os.environ.get("LOG_LEVEL", "INFO"))
logger = logging.getLogger("pdf_toolkit")

limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=[os.environ.get("RATE_LIMIT_DEFAULT", "120 per hour")],
    storage_uri=os.environ.get("RATELIMIT_STORAGE_URI", "memory://"),
    headers_enabled=True,
)

PDF = ".pdf"
IMG = ".jpg,.jpeg,.png,.webp,.bmp,.tif,.tiff"

class Ctx:
    def __init__(self, workdir, files, extra, form):
        self.dir = workdir
        self.files = files      
        self.extra = extra      
        self.form = form

    @property
    def first(self):
        return self.files[0]

    @property
    def stem(self):
        return Path(self.first).stem

    def out(self, name):
        return os.path.join(self.dir, name)

    def named(self, suffix, ext=".pdf"):
        return self.out(f"{self.stem}_{suffix}{ext}")

    def subdir(self, name):
        p = os.path.join(self.dir, name)
        os.makedirs(p, exist_ok=True)
        return p

    def s(self, key, default=""):
        v = self.form.get(key, "")
        return v.strip() if v and v.strip() else default

    def i(self, key, default=0):
        try:
            return int(float(self.s(key, default)))
        except ValueError:
            return default

    def f(self, key, default=0.0):
        try:
            return float(self.s(key, default))
        except ValueError:
            return default

    def b(self, key):
        return self.form.get(key) in ("on", "true", "1", "yes")

def h_merge(c):        return T.merge_pdf(c.files, c.out("merged.pdf"))
def h_split(c):        return T.split_pdf(c.first, c.subdir("parts"), c.s("ranges") or None)
def h_remove(c):       return T.remove_pages(c.first, c.named("removed"), c.s("pages"))
def h_extract(c):      return T.extract_pages(c.first, c.named("extracted"), c.s("pages"))
def h_organize(c):     return T.organize_pdf(c.first, c.named("organized"), c.s("order"))
def h_scan(c):         return T.scan_to_pdf(c.files, c.out("scan.pdf"), c.b("enhance"))
def h_compress(c):     return T.compress_pdf(c.first, c.named("compressed"), c.s("quality", "ebook"))
def h_repair(c):       return T.repair_pdf(c.first, c.named("repaired"))
def h_ocr(c):          return T.ocr_pdf(c.first, c.named("ocr"), c.s("language", "eng"))
def h_jpg2pdf(c):      return T.jpg_to_pdf(c.files, c.out("images.pdf"))
def h_word2pdf(c):     return T.word_to_pdf(c.first, c.named("converted"))
def h_ppt2pdf(c):      return T.powerpoint_to_pdf(c.first, c.named("converted"))
def h_excel2pdf(c):    return T.excel_to_pdf(c.first, c.named("converted"))


def h_html2pdf(c):
    # Remote URL rendering is intentionally disabled for the public service: fetching
    # arbitrary pages can enable SSRF against internal services and cloud metadata.
    if c.s("url"):
        raise ValueError("Remote URL conversion is disabled for security. Upload an HTML file instead.")
    if not c.files:
        raise ValueError("Upload an HTML file to convert it.")
    return T.html_to_pdf(c.first, c.named("converted"))


def h_pdf2jpg(c):      return T.pdf_to_jpg(c.first, c.subdir("pages"), c.i("dpi", 150), c.i("quality", 90))
def h_pdf2word(c):     return T.pdf_to_word(c.first, c.named("converted", ".docx"))
def h_pdf2ppt(c):      return T.pdf_to_powerpoint(c.first, c.named("converted", ".pptx"), c.i("dpi", 150))
def h_pdf2excel(c):    return T.pdf_to_excel(c.first, c.named("converted", ".xlsx"))
def h_pdf2pdfa(c):     return T.pdf_to_pdfa(c.first, c.named("pdfa"))
def h_rotate(c):       return T.rotate_pdf(c.first, c.named("rotated"), c.i("angle", 90), c.s("pages") or None)


def h_numbers(c):
    return T.add_page_numbers(c.first, c.named("numbered"), c.s("position", "bottom-center"),
                              c.s("fmt", "Page {n} of {total}"), c.i("start", 1), c.i("fontsize", 10))


def h_watermark(c):
    return T.add_watermark(c.first, c.named("watermarked"), c.s("text", "CONFIDENTIAL"),
                           c.i("fontsize", 60), opacity=c.f("opacity", 0.3),
                           angle=c.i("angle", 45), image=c.extra.get("image"))


def h_crop(c):
    return T.crop_pdf(c.first, c.named("cropped"), c.f("left"), c.f("top"), c.f("right"), c.f("bottom"))


def h_edit(c):
    replacements, add_text, add_image = {}, [], []
    if c.s("find"):
        replacements[c.s("find")] = c.s("replace")
    if c.s("text"):
        add_text.append({"page": c.i("page", 1), "x": c.f("x", 72), "y": c.f("y", 72),
                         "text": c.s("text"), "size": c.i("size", 12)})
    if c.extra.get("image"):
        x, y = c.f("x", 72), c.f("y", 72)
        add_image.append({"page": c.i("page", 1), "path": c.extra["image"],
                          "rect": [x, y, x + c.f("img_w", 150), y + c.f("img_h", 100)]})
    if not (replacements or add_text or add_image):
        raise ValueError("Enter text to find/replace, text to add, or an image to place.")
    return T.edit_pdf(c.first, c.named("edited"), replacements, add_text, add_image)


def h_forms(c):
    raw = c.s("values")
    if not raw:
        fields = T.list_form_fields(c.first)
        if not fields:
            return "No fillable form fields found in this PDF."
        return "Form fields (name | type | value):\n" + "\n".join(
            f"{k} | {t} | {v}" for k, (t, v) in fields.items())
    values = {}
    for line in raw.splitlines():
        if "=" in line:
            k, v = line.split("=", 1)
            v = v.strip()
            values[k.strip()] = True if v.lower() == "true" else False if v.lower() == "false" else v
    return T.fill_form(c.first, c.named("filled"), values)


def h_unlock(c):       return T.unlock_pdf(c.first, c.named("unlocked"), c.s("password"))


def h_protect(c):
    if not c.s("password"):
        raise ValueError("Enter a password.")
    return T.protect_pdf(c.first, c.named("protected"), c.s("password"))


def h_sign(c):
    x, y = c.f("x", 350), c.f("y", 700)
    return T.sign_pdf(c.first, c.named("signed"), c.extra["image"], c.i("page", 1),
                      (x, y, x + c.f("width", 180), y + c.f("height", 70)))


def h_redact(c):
    terms = [t.strip() for t in c.s("terms").replace(",", "\n").splitlines() if t.strip()]
    if not terms:
        raise ValueError("Enter at least one word or phrase to redact.")
    return T.redact_pdf(c.first, c.named("redacted"), terms)


def h_compare(c):
    if len(c.files) != 2:
        raise ValueError("Select exactly two PDF files.")
    return T.compare_pdf(c.files[0], c.files[1], c.out("comparison.html"))


def h_summary(c):      return T.ai_summarizer(c.first, c.s("length", "medium"))
def h_translate(c):    return T.translate_pdf(c.first, c.named("translated"), c.s("language", "Hindi"))
def h_markdown(c):     return T.pdf_to_markdown(c.first, c.named("converted", ".md"))


def P(name, label, kind="text", **kw):
    return {"name": name, "label": label, "type": kind, **kw}


def tool(id_, cat, label, desc, handler, accept=PDF, multiple=False, files_label=None,
         params=None, extra=None, files_required=True):
    return {"id": id_, "category": cat, "label": label, "desc": desc, "handler": handler,
            "accept": accept, "multiple": multiple, "files_label": files_label,
            "params": params or [], "extra": extra or [], "files_required": files_required}


ORG, OPT, TOPDF, FROMPDF = "Organize PDF", "Optimize PDF", "Convert to PDF", "Convert from PDF"
EDIT, SEC, AI = "Edit PDF", "PDF Security", "PDF Intelligence"

TOOL_LIST = [
    tool("merge", ORG, "Merge PDF", "Combine several PDFs into one, in the order you pick.",
         h_merge, multiple=True),
    tool("split", ORG, "Split PDF", "Split into single pages, or into custom page ranges.",
         h_split, params=[P("ranges", "Ranges (optional)", placeholder="e.g. 1-3,4-6  (empty = every page)")]),
    tool("remove", ORG, "Remove pages", "Delete the pages you don't need.",
         h_remove, params=[P("pages", "Pages to remove", placeholder="e.g. 2,4-6", required=True)]),
    tool("extract", ORG, "Extract pages", "Keep only the pages you choose.",
         h_extract, params=[P("pages", "Pages to keep", placeholder="e.g. 1,3-5", required=True)]),
    tool("organize", ORG, "Organize PDF", "Reorder, duplicate or drop pages.",
         h_organize, params=[P("order", "New page order", placeholder="e.g. 3,1,2", required=True)]),
    tool("scan", ORG, "Scan to PDF", "Turn photos of documents into a clean scanned PDF.",
         h_scan, accept=IMG, multiple=True, files_label="Photos / images",
         params=[P("enhance", "Enhance like a scanner (black & white)", "checkbox", default=True)]),

    tool("compress", OPT, "Compress PDF", "Reduce file size while keeping quality.",
         h_compress, params=[P("quality", "Compression level", "select", default="ebook", options=[
             ["screen", "Maximum (smallest file)"], ["ebook", "Recommended"],
             ["printer", "Low (better quality)"], ["prepress", "Minimal (best quality)"]])]),
    tool("repair", OPT, "Repair PDF", "Try to recover a damaged PDF.", h_repair),
    tool("ocr", OPT, "OCR PDF", "Make scanned PDFs searchable and selectable.",
         h_ocr, params=[P("language", "OCR language code", default="eng", placeholder="eng, hin, deu ...")]),

    tool("jpg2pdf", TOPDF, "JPG to PDF", "Convert images to a single PDF.",
         h_jpg2pdf, accept=IMG, multiple=True, files_label="Images"),
    tool("word2pdf", TOPDF, "WORD to PDF", "Convert DOC / DOCX to PDF.", h_word2pdf, accept=".doc,.docx,.odt,.rtf"),
    tool("ppt2pdf", TOPDF, "POWERPOINT to PDF", "Convert PPT / PPTX to PDF.", h_ppt2pdf, accept=".ppt,.pptx,.odp"),
    tool("excel2pdf", TOPDF, "EXCEL to PDF", "Convert XLS / XLSX to PDF.", h_excel2pdf, accept=".xls,.xlsx,.ods,.csv"),
    tool("html2pdf", TOPDF, "HTML to PDF", "Turn a web page or HTML file into a PDF.",
         h_html2pdf, accept=".html,.htm", files_required=False, files_label="HTML file (optional)",
         params=[P("url", "…or a web page URL", placeholder="https://example.com")]),

    tool("pdf2jpg", FROMPDF, "PDF to JPG", "Save every page as an image (ZIP).",
         h_pdf2jpg, params=[P("dpi", "Resolution (DPI)", "number", default=150, min=50, max=400),
                            P("quality", "JPG quality", "number", default=90, min=30, max=100)]),
    tool("pdf2word", FROMPDF, "PDF to WORD", "Convert to an editable Word document.", h_pdf2word),
    tool("pdf2ppt", FROMPDF, "PDF to POWERPOINT", "Each page becomes a slide.",
         h_pdf2ppt, params=[P("dpi", "Image quality (DPI)", "number", default=150, min=72, max=300)]),
    tool("pdf2excel", FROMPDF, "PDF to EXCEL", "Extract tables into a spreadsheet.", h_pdf2excel),
    tool("pdf2pdfa", FROMPDF, "PDF to PDF/A", "Convert to the archival PDF/A format.", h_pdf2pdfa),

    tool("rotate", EDIT, "Rotate PDF", "Rotate all pages, or only some.",
         h_rotate, params=[P("angle", "Rotate", "select", default="90", options=[
             ["90", "90° clockwise"], ["180", "180°"], ["270", "90° counter-clockwise"]]),
             P("pages", "Pages (optional)", placeholder="e.g. 1,3-4  (empty = all)")]),
    tool("numbers", EDIT, "Add page numbers", "Stamp page numbers on every page.",
         h_numbers, params=[
             P("position", "Position", "select", default="bottom-center", options=[
                 [v, v.replace("-", " ").title()] for v in
                 ["bottom-center", "bottom-right", "bottom-left", "top-center", "top-right", "top-left"]]),
             P("fmt", "Format", default="Page {n} of {total}", help="Use {n} and {total}"),
             P("start", "First number", "number", default=1),
             P("fontsize", "Font size", "number", default=10)]),
    tool("watermark", EDIT, "Add watermark", "Stamp text or an image over your PDF.",
         h_watermark, params=[
             P("text", "Watermark text", default="CONFIDENTIAL"),
             P("fontsize", "Font size", "number", default=60),
             P("opacity", "Opacity (0.1 – 1)", "number", default=0.3, step=0.1, min=0.1, max=1),
             P("angle", "Angle", "number", default=45)],
         extra=[{"name": "image", "label": "Image watermark (optional, replaces text)", "accept": IMG, "required": False}]),
    tool("crop", EDIT, "Crop PDF", "Trim margins (in points, 72 = 1 inch).",
         h_crop, params=[P("left", "Left", "number", default=0), P("top", "Top", "number", default=0),
                         P("right", "Right", "number", default=0), P("bottom", "Bottom", "number", default=0)]),
    tool("edit", EDIT, "Edit PDF", "Replace text, add text, or place an image.",
         h_edit, params=[
             P("find", "Find text", group="Replace text"), P("replace", "Replace with", group="Replace text"),
             P("text", "Text to add", group="Add content"),
             P("page", "Page", "number", default=1, group="Add content"),
             P("x", "X position", "number", default=72, group="Add content"),
             P("y", "Y position", "number", default=72, group="Add content"),
             P("size", "Font size", "number", default=12, group="Add content"),
             P("img_w", "Image width", "number", default=150, group="Add content"),
             P("img_h", "Image height", "number", default=100, group="Add content")],
         extra=[{"name": "image", "label": "Image to place (optional)", "accept": IMG, "required": False}]),
    tool("forms", EDIT, "PDF Forms", "List form fields, or fill them in.",
         h_forms, params=[P("values", "Values to fill (one per line, name=value)", "textarea",
                            placeholder="Leave empty to just list the form fields.\nName=John\nAgree=true")]),

    tool("unlock", SEC, "Unlock PDF", "Remove password protection (you need the password).",
         h_unlock, params=[P("password", "Current password", "password", required=True)]),
    tool("protect", SEC, "Protect PDF", "Encrypt with a password (AES-256).",
         h_protect, params=[P("password", "New password", "password", required=True)]),
    tool("sign", SEC, "Sign PDF", "Place your signature image on a page.",
         h_sign, params=[P("page", "Page", "number", default=1), P("x", "X", "number", default=350),
                         P("y", "Y", "number", default=700), P("width", "Width", "number", default=180),
                         P("height", "Height", "number", default=70)],
         extra=[{"name": "image", "label": "Signature image (PNG/JPG)", "accept": IMG, "required": True}]),
    tool("redact", SEC, "Redact PDF", "Permanently black out sensitive text.",
         h_redact, params=[P("terms", "Words / phrases to redact (comma or new line)", "textarea", required=True)]),
    tool("compare", SEC, "Compare PDF", "Side-by-side text differences of two PDFs.",
         h_compare, multiple=True, files_label="Select exactly 2 PDFs"),

    tool("summary", AI, "AI Summarizer", "Get a clear summary of a long PDF.",
         h_summary, params=[P("length", "Length", "select", default="medium", options=[
             ["short", "Short"], ["medium", "Medium"], ["detailed", "Detailed"]])]),
    tool("translate", AI, "Translate PDF", "Translate the whole PDF into another language.",
         h_translate, params=[P("language", "Translate to", default="Hindi")]),
    tool("markdown", AI, "PDF to Markdown", "Extract the PDF as clean Markdown.", h_markdown),
]

TOOLS = {t["id"]: t for t in TOOL_LIST}
CATEGORIES = [ORG, OPT, TOPDF, FROMPDF, EDIT, SEC, AI]


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/tools")
def api_tools():
    public = [{k: v for k, v in t.items() if k != "handler"} for t in TOOL_LIST]
    return jsonify({"categories": CATEGORIES, "tools": public})


def _save(storage, folder, idx):
    name = secure_filename(storage.filename or "") or f"file{idx}.bin"
    sub = os.path.join(folder, str(idx))      
    os.makedirs(sub, exist_ok=True)
    path = os.path.join(sub, name)
    storage.save(path)
    return path


def _zip_dir(paths):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            z.write(p, os.path.basename(p))
    buf.seek(0)
    return buf


@app.route("/api/run/<tool_id>", methods=["POST"])
@limiter.limit(os.environ.get("RATE_LIMIT_RUN", "10 per minute"))
def api_run(tool_id):
    spec = TOOLS.get(tool_id)
    if not spec:
        return jsonify(error="Unknown tool."), 404

    workdir = tempfile.mkdtemp(prefix="pdftool_")
    try:
        indir = os.path.join(workdir, "in")
        os.makedirs(indir)

        uploads = [f for f in request.files.getlist("files") if f and f.filename]
        if spec["files_required"] and not uploads:
            return jsonify(error="Please choose a file first."), 400
        if not spec["multiple"]:
            uploads = uploads[:1]
        # Bound file count and enforce the extension allowlist advertised by this tool.
        if len(uploads) > 10:
            return jsonify(error="At most 10 files may be uploaded per operation."), 400
        allowed = {x.strip().lower() for x in spec["accept"].split(",") if x.strip()}
        for storage in uploads:
            suffix = Path(storage.filename or "").suffix.lower()
            if suffix not in allowed:
                return jsonify(error=f"Unsupported file type for this tool. Allowed extensions: {', '.join(sorted(allowed))}."), 400
        files = [_save(f, indir, i) for i, f in enumerate(uploads)]
        for path in files:
            if Path(path).suffix.lower() == ".pdf":
                try:
                    from pypdf import PdfReader
                    reader = PdfReader(path, strict=False)
                    if len(reader.pages) > int(os.environ.get("MAX_PDF_PAGES", "500")):
                        return jsonify(error="PDF exceeds the maximum page count."), 413
                except (OSError, ValueError) as exc:
                    raise ValueError("The uploaded PDF could not be read. Check that it is a valid PDF.") from exc

        extra = {}
        for e in spec["extra"]:
            f = request.files.get(e["name"])
            if f and f.filename:
                extra_allowed = {x.strip().lower() for x in e.get("accept", "").split(",") if x.strip()}
                if extra_allowed and Path(f.filename).suffix.lower() not in extra_allowed:
                    return jsonify(error=f"Unsupported file type for {e['label']}."), 400
                extra[e["name"]] = _save(f, indir, 99)
            elif e.get("required"):
                return jsonify(error=f"Missing: {e['label']}"), 400

        for p in spec["params"]:
            if p.get("required") and p["type"] != "checkbox" and not request.form.get(p["name"], "").strip():
                return jsonify(error=f"Please fill in: {p['label']}"), 400

        ctx = Ctx(workdir, files, extra, request.form)
        result = spec["handler"](ctx)

        # Bound generated artifacts as well as incoming requests.
        max_output = int(os.environ.get("MAX_OUTPUT_BYTES", 100 * 1024 * 1024))
        output_paths = result if isinstance(result, (list, tuple)) else ([result] if isinstance(result, str) and os.path.isfile(result) else [])
        if any(os.path.getsize(path) > max_output for path in output_paths):
            return jsonify(error="Generated output exceeds the configured size limit."), 413

        if isinstance(result, str) and not os.path.exists(result):
            return jsonify(text=result)

        if isinstance(result, (list, tuple)):
            data = _zip_dir(result)
            name = f"{Path(files[0]).stem}_{tool_id}.zip" if files else f"{tool_id}.zip"
            return send_file(data, as_attachment=True, download_name=name, mimetype="application/zip")

        with open(result, "rb") as fh:
            data = io.BytesIO(fh.read())
        return send_file(data, as_attachment=True, download_name=Path(result).name)

    except ValueError as exc:  # bad input (wrong password, empty selection, ...)
        return jsonify(error=str(exc)), 400
    except Exception:
        # Detailed exception information belongs in server logs, not public responses.
        logger.exception("PDF tool failed: tool_id=%s", tool_id)
        return jsonify(error="Processing failed. Check the file and parameters, or try again later."), 500
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


@app.errorhandler(413)
def too_big(_):
    return jsonify(error="File too large (limit 200 MB)."), 413


@app.after_request
def security_headers(response):
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "DENY")
    response.headers.setdefault("Referrer-Policy", "no-referrer")
    response.headers.setdefault("Permissions-Policy", "camera=(), microphone=(), geolocation=()")
    response.headers.setdefault("Content-Security-Policy", "default-src 'self'; img-src 'self' data: blob:; style-src 'self' 'unsafe-inline'; script-src 'self' 'unsafe-inline'; base-uri 'self'; frame-ancestors 'none'")
    return response


if __name__ == "__main__":
    app.run(host=os.environ.get("HOST", "127.0.0.1"), port=int(os.environ.get("PORT", "5000")), debug=False)