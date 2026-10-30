import os
import json
import csv
import io
import re
import sqlite3
from datetime import datetime, timedelta
from dotenv import load_dotenv
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, session
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from fpdf import FPDF
from google import genai
from google.genai import types
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

load_dotenv()

app = Flask(__name__, template_folder="template")
app.secret_key = os.environ.get("SECRET_KEY", "cpit251-sbs-secret-key-2026")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.environ.get("SESSION_COOKIE_SECURE", "true").lower() == "true"

# Folders for storing uploaded images and structured .txt files
UPLOAD_FOLDER = "uploads"
TXT_FOLDER = "records_txt"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(TXT_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "webp", "txt", "mp3", "wav", "m4a", "webm", "ogg"}
IMAGE_EXTENSIONS = {"png", "jpg", "jpeg", "webp"}
AUDIO_EXTENSIONS = {"mp3", "wav", "m4a", "webm", "ogg"}

# Default Admin User (Hashed password for Rubric Section 2c Security)
ADMIN_USER = "wajid"
ADMIN_HASH = generate_password_hash("kau2026")

def get_db():
    conn = sqlite3.connect("sbs_records.db")
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                record_type TEXT,
                name TEXT,
                shop TEXT,
                doc_date TEXT,
                items TEXT,
                total_amount TEXT,
                plain_text TEXT,
                raw_txt_content TEXT,
                image_filename TEXT,
                audio_filename TEXT,
                txt_filename TEXT,
                created_at TEXT,
                deleted_at TEXT
            )
        """)
        columns = {row["name"] for row in conn.execute("PRAGMA table_info(records)").fetchall()}
        if "audio_filename" not in columns:
            conn.execute("ALTER TABLE records ADD COLUMN audio_filename TEXT")
        if "deleted_at" not in columns:
            conn.execute("ALTER TABLE records ADD COLUMN deleted_at TEXT")
        user = conn.execute("SELECT id FROM users WHERE email = ?", ("wajid@cydevlab.stream",)).fetchone()
        if not user:
            cur = conn.execute(
                "INSERT INTO users (email, password_hash, created_at) VALUES (?, ?, ?)",
                ("wajid@cydevlab.stream", generate_password_hash("kau2026"),
                 datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            )
            user_id = cur.lastrowid
        else:
            user_id = user["id"]
        record_columns = {row["name"] for row in conn.execute("PRAGMA table_info(records)").fetchall()}
        if "user_id" not in record_columns:
            conn.execute("ALTER TABLE records ADD COLUMN user_id INTEGER")
        conn.execute("UPDATE records SET user_id = ? WHERE user_id IS NULL", (user_id,))
init_db()

def current_user_id():
    return session.get("user_id")

def require_login():
    if not current_user_id():
        flash("Please log in to access your records.", "warning")
        return redirect(url_for("login", next=request.path))
    return None

def purge_expired_trash():
    cutoff = (datetime.now() - timedelta(days=10)).strftime("%Y-%m-%d %H:%M:%S")
    with get_db() as conn:
        expired = conn.execute(
            "SELECT image_filename, audio_filename, txt_filename FROM records "
            "WHERE deleted_at IS NOT NULL AND deleted_at < ?", (cutoff,)
        ).fetchall()
        conn.execute(
            "DELETE FROM records WHERE deleted_at IS NOT NULL AND deleted_at < ?", (cutoff,)
        )
    for record in expired:
        for folder, filename in ((UPLOAD_FOLDER, record["image_filename"]),
                                 (UPLOAD_FOLDER, record["audio_filename"]),
                                 (TXT_FOLDER, record["txt_filename"])):
            if filename:
                path = os.path.join(folder, filename)
                if os.path.exists(path):
                    os.remove(path)

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def filename_stem(filename):
    return os.path.splitext(secure_filename(filename))[0]

def format_structured_txt(data):
    """Formats a record as a professional plain-text report."""
    lines = [
        "SBS SMART CAPTURE",
        "Professional Structured Record",
        "=" * 68,
        f"Record title : {data.get('name', 'Untitled')}",
        f"Record type  : {data.get('record_type', 'General Document')}",
        f"Record date  : {data.get('doc_date', datetime.now().strftime('%Y-%m-%d'))}",
        f"Merchant     : {data.get('shop', 'N/A')}",
        f"Total amount : {data.get('total_amount', 'N/A')}",
        "",
        "PURCHASED ITEMS",
        "-" * 68,
        f"{data.get('items', 'None')}",
        "",
        "DOCUMENT NOTES",
        "-" * 68,
        f"{data.get('plain_text', '')}",
        "",
        "=" * 68,
        "Generated by SBS Smart Capture | Organized records, simplified."
    ]
    return "\n".join(lines)

def export_stem(record):
    return secure_filename(record["name"] or f"record_{record['id']}") or f"record_{record['id']}"

def parse_pdf_items(value):
    """Convert stored item text into clean item, quantity, unit, and price rows."""
    rows = []
    for raw_line in str(value or "").splitlines():
        line = raw_line.strip()
        if not line or line.upper() in {"N/A", "NONE", "-"}:
            continue

        parts = [part.strip() for part in re.split(r"\s*\|\s*", line)]
        item = parts[0]
        quantity = "1"
        unit = "unit"
        price = "N/A"

        if len(parts) >= 3:
            quantity_text, price = parts[-2], parts[-1]
            item = " | ".join(parts[:-2]).strip()
        elif len(parts) == 2:
            if re.search(r"->|→", parts[0]):
                item, price = parts[0], parts[1]
                quantity_text = ""
            else:
                quantity_text = parts[1]
        else:
            quantity_text = ""

        quantity_match = re.search(
            r"(?<![\w.])(\d+(?:[.,]\d+)?)\s*(pc|pcs|piece|pieces|qty|unit|units|kg|g|mg|lb|oz|ml|l|qt|box|boxes|pack|packs)?\b",
            quantity_text,
            flags=re.IGNORECASE,
        )
        if quantity_match:
            quantity = quantity_match.group(1).replace(",", ".")
            unit = (quantity_match.group(2) or "unit").lower()
            unit = {
                "pcs": "pc", "piece": "pc", "pieces": "pc",
                "qty": "unit", "units": "unit", "boxes": "box",
                "packs": "pack",
            }.get(unit, unit)
        elif quantity_text:
            quantity = quantity_text

        arrow_match = re.search(
            r"(.+?)\s*(?:->|→|[-:])\s*(\d+(?:[.,]\d+)?)\s*(pc|pcs|piece|pieces|qty|unit|units|kg|g|mg|lb|oz|ml|l|qt|box|boxes|pack|packs)?\b",
            item,
            flags=re.IGNORECASE,
        )
        if arrow_match:
            item = arrow_match.group(1).strip()
            quantity = arrow_match.group(2).replace(",", ".")
            unit = (arrow_match.group(3) or "unit").lower()
            unit = {"pcs": "pc", "piece": "pc", "pieces": "pc", "boxes": "box", "packs": "pack"}.get(unit, unit)

        rows.append((item or "Unspecified item", quantity, unit, price or "N/A"))

    return rows or [("No purchased items listed", "—", "—", "N/A")]

def build_pdf(record):
    pdf = FPDF()
    pdf.set_margins(18, 18, 18)
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.add_page()
    pdf.set_fill_color(36, 59, 115)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 20)
    pdf.cell(0, 14, "SBS SMART CAPTURE", fill=True, ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 8, "Professional Structured Record", fill=True, ln=True)
    pdf.ln(8)
    pdf.set_text_color(33, 41, 66)
    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 9, "Record summary", ln=True)
    pdf.set_font("Helvetica", "", 10)
    fields = [
        ("Record title", record["name"]), ("Record type", record["record_type"]),
        ("Record date", record["doc_date"]), ("Merchant", record["shop"]),
        ("Total amount", record["total_amount"]),
    ]
    for label, value in fields:
        pdf.set_fill_color(239, 244, 255)
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(42, 8, label, border=1, fill=True)
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 8, str(value or "N/A"), border=1, ln=True)
    pdf.ln(7)
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(33, 41, 66)
    pdf.cell(0, 9, "Purchased items", ln=True)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(36, 59, 115)
    pdf.set_text_color(255, 255, 255)
    column_widths = [78, 25, 25, 46]
    headers = ["Item", "Quantity", "Unit", "Price"]
    for width, header in zip(column_widths, headers):
        pdf.cell(width, 9, header, border=1, fill=True, align="C")
    pdf.ln()

    for row_index, row in enumerate(parse_pdf_items(record["items"])):
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(33, 41, 66)
        pdf.set_fill_color(248, 251, 255) if row_index % 2 == 0 else pdf.set_fill_color(255, 255, 255)
        values = [str(value).encode("latin-1", "replace").decode("latin-1") for value in row]
        for width, value in zip(column_widths, values):
            pdf.cell(width, 8, value[:55], border=1, fill=True, align="C" if width != 78 else "L")
        pdf.ln()

    pdf.ln(7)
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(33, 41, 66)
    pdf.cell(0, 9, "Document notes", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_fill_color(248, 251, 255)
    safe_content = str(record["plain_text"] or "N/A").encode("latin-1", "replace").decode("latin-1")
    pdf.multi_cell(0, 7, safe_content, border=1, fill=True)
    pdf.ln(8)
    pdf.set_font("Helvetica", "I", 9)
    pdf.set_text_color(100, 110, 130)
    pdf.cell(0, 7, "Generated by SBS Smart Capture", align="C")
    return pdf.output()

def analyze_with_gemini(image_path=None, voice_text=None):
    """Calls Gemini API to categorize into Receipt/Purchase OR General Document."""
    api_key = os.environ.get("GEMINI_API_KEY", "")
    today = datetime.now().strftime("%Y-%m-%d")

    prompt = f"""
    Analyze this input carefully.
    CRITERIA:
    1. If it relates to a purchase, receipt, invoice, or expense:
       Set "record_type" to "Purchase/Receipt". Extract the "name" (summary title), "shop" (store/merchant name), "doc_date" (YYYY-MM-DD, default {today}), "items" (formatted as 'Item | Qty | Price' per line), "total_amount", and "plain_text" (any extra notes).
    2. If it is NOT related to a purchase/shop:
       Set "record_type" to "General Document". Set "shop" to "N/A", "items" to "N/A", "total_amount" to "N/A", and extract "name" (title/subject), "doc_date" ({today}), and put all full document information into "plain_text".

    Return ONLY valid JSON with keys:
    "record_type", "name", "shop", "doc_date", "items", "total_amount", "plain_text"
    """

    if not api_key:
        # Automatic offline fallback so your live class demo NEVER crashes
        return {
            "record_type": "Purchase/Receipt" if not voice_text else "Voice Record",
            "name": "Sample Extracted Record (Add GEMINI_API_KEY for Live AI)",
            "shop": "Albaik / Local Store",
            "doc_date": today,
            "items": "Meal Combo | 1 | 22.00 SAR\nDrink | 1 | 5.00 SAR",
            "total_amount": "27.00 SAR",
            "plain_text": voice_text if voice_text else "OCR Extracted Receipt Text. Edit any field below."
        }

    try:
        client = genai.Client(api_key=api_key)
        contents = [prompt]
        if image_path and os.path.exists(image_path):
            with open(image_path, "rb") as f:
                img_bytes = f.read()
            mime = "image/png" if image_path.endswith(".png") else "image/jpeg"
            contents.append(types.Part.from_bytes(data=img_bytes, mime_type=mime))
        if voice_text:
            contents.append(f"Input Text/Voice Transcript: {voice_text}")

        response = client.models.generate_content(
            model="gemini-3.6-flash",
            contents=contents,
            config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        return json.loads(response.text)
    except Exception as e:
        app.logger.warning("Gemini extraction unavailable: %s", e)
        return {
            "record_type": "General Document",
            "name": "Manual Review Needed",
            "shop": "N/A",
            "doc_date": today,
            "items": "N/A",
            "total_amount": "N/A",
            "plain_text": (
                f"Raw Input: {voice_text or 'Image uploaded'}. "
                "Automatic extraction was temporarily unavailable; please review and complete the fields manually."
            )
        }

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password")
        with get_db() as conn:
            user = conn.execute("SELECT * FROM users WHERE email = ?", (email,)).fetchone()
        if user and check_password_hash(user["password_hash"], password):
            session.clear()
            session["user_id"] = user["id"]
            session["user_email"] = user["email"]
            return redirect(url_for("index"))
        flash("Invalid email or password.", "danger")
    return render_template("index.html", page="login")

@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        confirm_password = request.form.get("confirm_password", "")
        if "@" not in email or "." not in email.split("@")[-1]:
            flash("Enter a valid email address.", "danger")
        elif len(password) < 8:
            flash("Password must be at least 8 characters.", "danger")
        elif password != confirm_password:
            flash("Passwords do not match.", "danger")
        else:
            try:
                with get_db() as conn:
                    conn.execute(
                        "INSERT INTO users (email, password_hash, created_at) VALUES (?, ?, ?)",
                        (email, generate_password_hash(password),
                         datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
                    )
                flash("Account created. You can now log in.", "success")
                return redirect(url_for("login"))
            except sqlite3.IntegrityError:
                flash("An account with that email already exists.", "warning")
    return render_template("index.html", page="register")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))

@app.route("/")
@app.route("/records/search")
def index():
    auth_redirect = require_login()
    if auth_redirect:
        return auth_redirect
    purge_expired_trash()
    user_id = current_user_id()
    with get_db() as conn:
        total_records = conn.execute("SELECT COUNT(*) FROM records WHERE user_id = ? AND deleted_at IS NULL", (user_id,)).fetchone()[0]
        purchase_records = conn.execute(
            "SELECT COUNT(*) FROM records WHERE user_id = ? AND deleted_at IS NULL AND record_type = ?", (user_id, "Purchase/Receipt")
        ).fetchone()[0]
        general_records = conn.execute(
            "SELECT COUNT(*) FROM records WHERE user_id = ? AND deleted_at IS NULL AND record_type = ?", (user_id, "General Document")
        ).fetchone()[0]
        recent_records = conn.execute(
            "SELECT * FROM records WHERE user_id = ? AND deleted_at IS NULL ORDER BY id DESC LIMIT 7", (user_id,)
        ).fetchall()
    return render_template(
        "index.html",
        page="dashboard",
        records=recent_records,
        total_records=total_records,
        purchase_records=purchase_records,
        general_records=general_records,
    )

@app.route("/records")
def records_page():
    auth_redirect = require_login()
    if auth_redirect:
        return auth_redirect
    purge_expired_trash()
    query = request.args.get("q", "").strip()
    rtype = request.args.get("type", "").strip()
    with get_db() as conn:
        sql = "SELECT * FROM records WHERE user_id = ? AND deleted_at IS NULL"
        params = [current_user_id()]
        if query:
            sql += " AND (name LIKE ? OR shop LIKE ? OR plain_text LIKE ? OR items LIKE ?)"
            params.extend([f"%{query}%"] * 4)
        if rtype:
            sql += " AND record_type = ?"
            params.append(rtype)
        sql += " ORDER BY id DESC"
        records = conn.execute(sql, params).fetchall()
    return render_template("index.html", page="records", records=records, query=query, rtype=rtype)

@app.route("/about")
def about():
    return render_template("index.html", page="about")

@app.route("/help")
def help_page():
    return render_template("index.html", page="help")

@app.route("/trash")
def trash_page():
    auth_redirect = require_login()
    if auth_redirect:
        return auth_redirect
    purge_expired_trash()
    with get_db() as conn:
        records = conn.execute(
            "SELECT *, julianday(?) - julianday(deleted_at) AS age_days "
            "FROM records WHERE user_id = ? AND deleted_at IS NOT NULL ORDER BY deleted_at DESC",
            (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), current_user_id()),
        ).fetchall()
    return render_template("index.html", page="trash", records=records)

@app.route("/records/export-excel", methods=["POST"])
def export_excel():
    auth_redirect = require_login()
    if auth_redirect:
        return auth_redirect
    purge_expired_trash()
    selected_ids = request.form.getlist("record_ids")
    export_all = request.form.get("export_all") == "1"

    with get_db() as conn:
        if export_all:
            records = conn.execute("SELECT * FROM records WHERE user_id = ? AND deleted_at IS NULL ORDER BY id DESC", (current_user_id(),)).fetchall()
        elif selected_ids:
            placeholders = ",".join("?" for _ in selected_ids)
            records = conn.execute(
                f"SELECT * FROM records WHERE user_id = ? AND deleted_at IS NULL AND id IN ({placeholders}) ORDER BY id DESC",
                [current_user_id()] + [int(record_id) for record_id in selected_ids],
            ).fetchall()
        else:
            flash("Select at least one record to export.", "warning")
            return redirect(url_for("records_page"))

    if not records:
        flash("There are no records available for export.", "warning")
        return redirect(url_for("records_page"))

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "SBS Records"
    headers = ["ID", "Record Type", "Name", "Merchant", "Date", "Total Amount", "Items", "Description"]
    sheet.append(["SBS SMART CAPTURE - RECORD EXPORT"])
    sheet.append([f"Exported: {datetime.now().strftime('%Y-%m-%d %H:%M')}"])
    sheet.append(headers)

    header_fill = PatternFill("solid", fgColor="DCEEFF")
    header_font = Font(color="17365D", bold=True)
    title_font = Font(color="17365D", bold=True, size=16)
    title_fill = PatternFill("solid", fgColor="B9D9F5")
    subtitle_fill = PatternFill("solid", fgColor="EEF7FF")
    thin_border = Border(
        left=Side(style="thin", color="B7CDE3"),
        right=Side(style="thin", color="B7CDE3"),
        top=Side(style="thin", color="B7CDE3"),
        bottom=Side(style="thin", color="B7CDE3"),
    )

    for cell in sheet[1]:
        cell.fill = title_fill
        cell.font = title_font
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = thin_border
    for cell in sheet[2]:
        cell.fill = subtitle_fill
        cell.font = Font(color="52677D", italic=True)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = thin_border
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    sheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))

    for cell in sheet[3]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for record in records:
        sheet.append([
            record["id"], record["record_type"], record["name"], record["shop"],
            record["doc_date"], record["total_amount"], record["items"], record["plain_text"],
        ])

    for row in sheet.iter_rows(min_row=4, max_row=sheet.max_row, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if cell.row % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F7FBFF")

    widths = [9, 22, 30, 24, 15, 18, 38, 55]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.row_dimensions[1].height = 30
    sheet.row_dimensions[2].height = 22
    sheet.freeze_panes = "A4"
    sheet.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{sheet.max_row}"
    sheet.print_title_rows = "1:3"
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_margins.left = 0.25
    sheet.page_margins.right = 0.25
    sheet.page_margins.top = 0.5
    sheet.page_margins.bottom = 0.5

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name=f"sbs_records_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

@app.route("/records/add", methods=["POST"])
def add_record():
    auth_redirect = require_login()
    if auth_redirect:
        return auth_redirect
    voice_text = request.form.get("voice_text", "").strip()
    requested_name = request.form.get("record_name", "").strip()
    requested_type = request.form.get("record_type", "").strip()
    file = request.files.get("file")
    audio_file = request.files.get("audio_file")
    img_filename = ""
    img_path = None
    audio_filename = ""
    uploaded_name = ""

    if file and file.filename and allowed_file(file.filename):
        uploaded_name = filename_stem(file.filename)
        ext = file.filename.rsplit(".", 1)[1].lower()
        img_filename = secure_filename(f"scan_{int(datetime.now().timestamp())}.{ext}")
        img_path = os.path.join(UPLOAD_FOLDER, img_filename)
        file.save(img_path)
        if ext == "txt":
            with open(img_path, "r", encoding="utf-8", errors="ignore") as f:
                voice_text = f.read()
            img_path = None
        elif ext not in IMAGE_EXTENSIONS:
            os.remove(img_path)
            img_filename = ""
            img_path = None

    if audio_file and audio_file.filename:
        audio_ext = audio_file.filename.rsplit(".", 1)[-1].lower()
        if audio_ext not in AUDIO_EXTENSIONS:
            flash("Audio must be MP3, WAV, M4A, WEBM, or OGG.", "danger")
            return redirect(url_for("index"))
        try:
            audio_duration = float(request.form.get("audio_duration", "0") or 0)
        except ValueError:
            audio_duration = 0
        if audio_duration > 300:
            flash("Audio recordings must be five minutes or shorter.", "danger")
            return redirect(url_for("index"))
        audio_filename = secure_filename(f"audio_{int(datetime.now().timestamp())}.{audio_ext}")
        audio_file.save(os.path.join(UPLOAD_FOLDER, audio_filename))

    extracted = analyze_with_gemini(image_path=img_path, voice_text=voice_text)
    if audio_filename and not img_filename:
        extracted["record_type"] = requested_type or "Audio"
        if not voice_text:
            extracted["plain_text"] = "Audio note attached to this record."
    record_name = requested_name or uploaded_name or extracted.get("name", "Untitled")
    extracted["name"] = record_name
    raw_txt = format_structured_txt(extracted)
    txt_stem = secure_filename(record_name) or f"record_{int(datetime.now().timestamp())}"
    txt_filename = f"{txt_stem} (SBS).txt"
    txt_path = os.path.join(TXT_FOLDER, txt_filename)
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(raw_txt)

    with get_db() as conn:
        cur = conn.execute("""
            INSERT INTO records (user_id, record_type, name, shop, doc_date, items, total_amount,
                                 plain_text, raw_txt_content, image_filename, audio_filename, txt_filename, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            current_user_id(),
            extracted.get("record_type", "General Document"),
            extracted.get("name", "Untitled"),
            extracted.get("shop", "N/A"),
            extracted.get("doc_date", datetime.now().strftime("%Y-%m-%d")),
            extracted.get("items", "N/A"),
            extracted.get("total_amount", "N/A"),
            extracted.get("plain_text", ""),
            raw_txt,
            img_filename,
            audio_filename,
            txt_filename,
            datetime.now().strftime("%Y-%m-%d %H:%M")
        ))
        new_id = cur.lastrowid

    flash("Record processed & saved to structured .TXT file! You can review or edit below.", "success")
    return redirect(url_for("view_record", id=new_id))

@app.route("/records/view")
def view_record():
    auth_redirect = require_login()
    if auth_redirect:
        return auth_redirect
    purge_expired_trash()
    rec_id = request.args.get("id", type=int)
    with get_db() as conn:
        record = conn.execute("SELECT * FROM records WHERE user_id = ? AND id = ? AND deleted_at IS NULL", (current_user_id(), rec_id)).fetchone()
    if not record:
        flash("Record not found.", "warning")
        return redirect(url_for("index"))
    return render_template("index.html", page="view", record=record)

@app.route("/records/edit", methods=["POST"])
def edit_record():
    auth_redirect = require_login()
    if auth_redirect:
        return auth_redirect
    rec_id = request.form.get("id", type=int)
    data = {
        "record_type": request.form.get("record_type"),
        "name": request.form.get("name"),
        "shop": request.form.get("shop"),
        "doc_date": request.form.get("doc_date"),
        "items": request.form.get("items"),
        "total_amount": request.form.get("total_amount"),
        "plain_text": request.form.get("plain_text")
    }
    raw_txt = format_structured_txt(data)

    with get_db() as conn:
        rec = conn.execute(
            "SELECT txt_filename FROM records WHERE user_id = ? AND id = ? AND deleted_at IS NULL", (current_user_id(), rec_id)
        ).fetchone()
        if rec and rec["txt_filename"]:
            with open(os.path.join(TXT_FOLDER, rec["txt_filename"]), "w", encoding="utf-8") as f:
                f.write(raw_txt)
        conn.execute("""
            UPDATE records SET record_type=?, name=?, shop=?, doc_date=?, items=?,
            total_amount=?, plain_text=?, raw_txt_content=? WHERE user_id=? AND id=?
        """, (data["record_type"], data["name"], data["shop"], data["doc_date"],
              data["items"], data["total_amount"], data["plain_text"], raw_txt,
              current_user_id(), rec_id))

    flash("Record and structured .TXT file updated!", "success")
    return redirect(url_for("records_page"))

@app.route("/records/delete", methods=["POST"])
def delete_record():
    auth_redirect = require_login()
    if auth_redirect:
        return auth_redirect
    rec_id = request.form.get("id", type=int)
    with get_db() as conn:
        conn.execute(
            "UPDATE records SET deleted_at = ? WHERE user_id = ? AND id = ? AND deleted_at IS NULL",
            (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), current_user_id(), rec_id),
        )
    flash("Record moved to the recycle bin for 10 days.", "info")
    return redirect(url_for("index"))

@app.route("/records/trash", methods=["POST"])
def trash_records():
    auth_redirect = require_login()
    if auth_redirect:
        return auth_redirect
    selected_ids = request.form.getlist("record_ids")
    if not selected_ids:
        flash("Select at least one record to move to the recycle bin.", "warning")
        return redirect(url_for("records_page"))
    placeholders = ",".join("?" for _ in selected_ids)
    with get_db() as conn:
        conn.execute(
            f"UPDATE records SET deleted_at = ? WHERE user_id = ? AND deleted_at IS NULL AND id IN ({placeholders})",
            [datetime.now().strftime("%Y-%m-%d %H:%M:%S"), current_user_id()] + [int(record_id) for record_id in selected_ids],
        )
    flash(f"{len(selected_ids)} record(s) moved to the recycle bin for 10 days.", "info")
    return redirect(url_for("records_page"))

@app.route("/trash/restore", methods=["POST"])
def restore_record():
    auth_redirect = require_login()
    if auth_redirect:
        return auth_redirect
    rec_id = request.form.get("id", type=int)
    with get_db() as conn:
        conn.execute("UPDATE records SET deleted_at = NULL WHERE user_id = ? AND id = ?", (current_user_id(), rec_id))
    flash("Record restored to the active library.", "success")
    return redirect(url_for("trash_page"))

@app.route("/trash/delete", methods=["POST"])
def permanently_delete_record():
    auth_redirect = require_login()
    if auth_redirect:
        return auth_redirect
    rec_id = request.form.get("id", type=int)
    with get_db() as conn:
        record = conn.execute(
            "SELECT image_filename, audio_filename, txt_filename FROM records "
            "WHERE user_id = ? AND id = ? AND deleted_at IS NOT NULL", (current_user_id(), rec_id)
        ).fetchone()
        conn.execute("DELETE FROM records WHERE user_id = ? AND id = ? AND deleted_at IS NOT NULL", (current_user_id(), rec_id))
    if record:
        for folder, filename in ((UPLOAD_FOLDER, record["image_filename"]),
                                 (UPLOAD_FOLDER, record["audio_filename"]),
                                 (TXT_FOLDER, record["txt_filename"])):
            if filename:
                path = os.path.join(folder, filename)
                if os.path.exists(path):
                    os.remove(path)
    flash("Record permanently deleted.", "warning")
    return redirect(url_for("trash_page"))

@app.route("/records/download")
def download_record():
    auth_redirect = require_login()
    if auth_redirect:
        return auth_redirect
    rec_id = request.args.get("id", type=int)
    fmt = request.args.get("format", "txt")
    with get_db() as conn:
        rec = conn.execute("SELECT * FROM records WHERE user_id = ? AND id = ? AND deleted_at IS NULL", (current_user_id(), rec_id)).fetchone()
    if not rec:
        return redirect(url_for("index"))

    stem = export_stem(rec)
    if fmt == "txt":
        return send_file(io.BytesIO(rec["raw_txt_content"].encode("utf-8")),
                         as_attachment=True, download_name=f"{stem} (SBS).txt", mimetype="text/plain")
    elif fmt == "pdf":
        pdf_bytes = build_pdf(rec)
        return send_file(io.BytesIO(pdf_bytes),
                         as_attachment=True, download_name=f"{stem} (SBS).pdf", mimetype="application/pdf")

@app.route("/uploads/<filename>")
def uploaded_file(filename):
    auth_redirect = require_login()
    if auth_redirect:
        return auth_redirect
    safe_name = secure_filename(filename)
    with get_db() as conn:
        owns_file = conn.execute(
            "SELECT 1 FROM records WHERE user_id = ? AND (image_filename = ? OR audio_filename = ?)",
            (current_user_id(), safe_name, safe_name),
        ).fetchone()
    if not owns_file:
        return redirect(url_for("index"))
    return send_file(os.path.join(UPLOAD_FOLDER, safe_name))

@app.route("/records/export-csv", methods=["POST"])
def export_csv():
    auth_redirect = require_login()
    if auth_redirect:
        return auth_redirect
    selected_ids = request.form.getlist("record_ids")
    if not selected_ids:
        flash("Select at least one record for CSV export.", "warning")
        return redirect(url_for("records_page"))
    with get_db() as conn:
        placeholders = ",".join("?" for _ in selected_ids)
        records = conn.execute(
            f"SELECT * FROM records WHERE user_id = ? AND deleted_at IS NULL AND id IN ({placeholders}) ORDER BY id DESC",
            [current_user_id()] + [int(record_id) for record_id in selected_ids],
        ).fetchall()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["SBS SMART CAPTURE - SELECTED RECORDS"])
    writer.writerow(["Exported", datetime.now().strftime("%Y-%m-%d %H:%M")])
    writer.writerow([])
    writer.writerow(["ID", "Record Type", "Record Name", "Merchant", "Date", "Total Amount", "Purchased Items", "Document Notes"])
    for record in records:
        writer.writerow([
            record["id"], record["record_type"], record["name"], record["shop"],
            record["doc_date"], record["total_amount"], record["items"], record["plain_text"],
        ])
    return send_file(
        io.BytesIO(output.getvalue().encode("utf-8-sig")),
        as_attachment=True,
        download_name="SBS Selected Records.csv",
        mimetype="text/csv",
    )

if __name__ == "__main__":
    app.run(debug=True, port=5000)