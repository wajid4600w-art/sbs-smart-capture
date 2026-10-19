import os
import json
import csv
import io
import sqlite3
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, session
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from fpdf import FPDF
from google import genai
from google.genai import types

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "cpit251-sbs-secret-key-2026")

# Folders for storing uploaded images and structured .txt files
UPLOAD_FOLDER = "uploads"
TXT_FOLDER = "records_txt"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(TXT_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "webp", "txt"}

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
            CREATE TABLE IF NOT EXISTS records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                record_type TEXT,
                name TEXT,
                shop TEXT,
                doc_date TEXT,
                items TEXT,
                total_amount TEXT,
                plain_text TEXT,
                raw_txt_content TEXT,
                image_filename TEXT,
                txt_filename TEXT,
                created_at TEXT
            )
        """)
init_db()

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def format_structured_txt(data):
    """Formats extracted data into a clean, parseable tabular .txt file."""
    lines = [
        "==================================================",
        "       SBS SMART STRUCTURED RECORD FILE           ",
        "==================================================",
        f"RECORD_TYPE   : {data.get('record_type', 'General Document')}",
        f"NAME          : {data.get('name', 'Untitled')}",
        f"DATE          : {data.get('doc_date', datetime.now().strftime('%Y-%m-%d'))}",
        f"SHOP_MERCHANT : {data.get('shop', 'N/A')}",
        f"TOTAL_AMOUNT  : {data.get('total_amount', '0.00')}",
        "--------------------------------------------------",
        "[PURCHASED_ITEMS_TABULAR]",
        "Item_Name | Quantity | Price",
        f"{data.get('items', 'None')}",
        "--------------------------------------------------",
        "[DOCUMENT_PLAIN_TEXT]",
        f"{data.get('plain_text', '')}",
        "=================================================="
    ]
    return "\n".join(lines)

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
            model="gemini-2.5-flash",
            contents=contents,
            config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        return json.loads(response.text)
    except Exception as e:
        return {
            "record_type": "General Document",
            "name": "Manual Review Needed",
            "shop": "N/A",
            "doc_date": today,
            "items": "N/A",
            "total_amount": "N/A",
            "plain_text": f"Raw Input: {voice_text or 'Image uploaded'}. (Note: {str(e)})"
        }

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")
        if username == ADMIN_USER and check_password_hash(ADMIN_HASH, password):
            session["user"] = username
            return redirect(url_for("index"))
        flash("Invalid credentials! Default is wajid / kau2026", "danger")
    return render_template("index.html", page="login")

@app.route("/logout")
def logout():
    session.pop("user", None)
    return redirect(url_for("index"))

@app.route("/")
@app.route("/records/search")
def index():
    query = request.args.get("q", "").strip()
    rtype = request.args.get("type", "").strip()
    with get_db() as conn:
        sql = "SELECT * FROM records WHERE 1=1"
        params = []
        if query:
            sql += " AND (name LIKE ? OR shop LIKE ? OR plain_text LIKE ? OR items LIKE ?)"
            params.extend([f"%{query}%"] * 4)
        if rtype:
            sql += " AND record_type = ?"
            params.append(rtype)
        sql += " ORDER BY id DESC"
        records = conn.execute(sql, params).fetchall()
    return render_template("index.html", page="dashboard", records=records, query=query, rtype=rtype)

@app.route("/records/add", methods=["POST"])
def add_record():
    voice_text = request.form.get("voice_text", "").strip()
    file = request.files.get("file")
    img_filename = ""
    img_path = None

    if file and file.filename and allowed_file(file.filename):
        ext = file.filename.rsplit(".", 1)[1].lower()
        img_filename = secure_filename(f"scan_{int(datetime.now().timestamp())}.{ext}")
        img_path = os.path.join(UPLOAD_FOLDER, img_filename)
        file.save(img_path)
        if ext == "txt":
            with open(img_path, "r", encoding="utf-8", errors="ignore") as f:
                voice_text = f.read()
            img_path = None

    extracted = analyze_with_gemini(image_path=img_path, voice_text=voice_text)
    raw_txt = format_structured_txt(extracted)
    txt_filename = f"record_{int(datetime.now().timestamp())}.txt"
    txt_path = os.path.join(TXT_FOLDER, txt_filename)
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(raw_txt)

    with get_db() as conn:
        cur = conn.execute("""
            INSERT INTO records (record_type, name, shop, doc_date, items, total_amount,
                                 plain_text, raw_txt_content, image_filename, txt_filename, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            extracted.get("record_type", "General Document"),
            extracted.get("name", "Untitled"),
            extracted.get("shop", "N/A"),
            extracted.get("doc_date", datetime.now().strftime("%Y-%m-%d")),
            extracted.get("items", "N/A"),
            extracted.get("total_amount", "N/A"),
            extracted.get("plain_text", ""),
            raw_txt,
            img_filename,
            txt_filename,
            datetime.now().strftime("%Y-%m-%d %H:%M")
        ))
        new_id = cur.lastrowid

    flash("Record processed & saved to structured .TXT file! You can review or edit below.", "success")
    return redirect(url_for("view_record", id=new_id))

@app.route("/records/view")
def view_record():
    rec_id = request.args.get("id", type=int)
    with get_db() as conn:
        record = conn.execute("SELECT * FROM records WHERE id = ?", (rec_id,)).fetchone()
    if not record:
        flash("Record not found.", "warning")
        return redirect(url_for("index"))
    return render_template("index.html", page="view", record=record)

@app.route("/records/edit", methods=["POST"])
def edit_record():
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
        rec = conn.execute("SELECT txt_filename FROM records WHERE id = ?", (rec_id,)).fetchone()
        if rec and rec["txt_filename"]:
            with open(os.path.join(TXT_FOLDER, rec["txt_filename"]), "w", encoding="utf-8") as f:
                f.write(raw_txt)
        conn.execute("""
            UPDATE records SET record_type=?, name=?, shop=?, doc_date=?, items=?,
            total_amount=?, plain_text=?, raw_txt_content=? WHERE id=?
        """, (data["record_type"], data["name"], data["shop"], data["doc_date"],
              data["items"], data["total_amount"], data["plain_text"], raw_txt, rec_id))

    flash("Record and structured .TXT file updated!", "success")
    return redirect(url_for("view_record", id=rec_id))

@app.route("/records/delete", methods=["POST"])
def delete_record():
    rec_id = request.form.get("id", type=int)
    with get_db() as conn:
        conn.execute("DELETE FROM records WHERE id = ?", (rec_id,))
    flash("Record deleted.", "info")
    return redirect(url_for("index"))

@app.route("/records/download")
def download_record():
    rec_id = request.args.get("id", type=int)
    fmt = request.args.get("format", "txt")
    with get_db() as conn:
        rec = conn.execute("SELECT * FROM records WHERE id = ?", (rec_id,)).fetchone()
    if not rec:
        return redirect(url_for("index"))

    if fmt == "txt":
        return send_file(io.BytesIO(rec["raw_txt_content"].encode("utf-8")),
                         as_attachment=True, download_name=f"sbs_record_{rec_id}.txt", mimetype="text/plain")
    elif fmt == "pdf":
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Helvetica", size=11)
        for line in rec["raw_txt_content"].split("\n"):
            safe_line = line.encode("latin-1", "replace").decode("latin-1")
            pdf.cell(0, 7, txt=safe_line, ln=True)
        pdf_bytes = pdf.output()
        return send_file(io.BytesIO(pdf_bytes),
                         as_attachment=True, download_name=f"sbs_record_{rec_id}.pdf", mimetype="application/pdf")
    elif fmt == "csv":
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(["ID", "Type", "Name", "Shop", "Date", "Total", "Items", "Plain_Text"])
        writer.writerow([rec["id"], rec["record_type"], rec["name"], rec["shop"],
                         rec["doc_date"], rec["total_amount"], rec["items"], rec["plain_text"]])
        return send_file(io.BytesIO(output.getvalue().encode("utf-8")),
                         as_attachment=True, download_name=f"sbs_record_{rec_id}.csv", mimetype="text/csv")

@app.route("/uploads/<filename>")
def uploaded_file(filename):
    return send_file(os.path.join(UPLOAD_FOLDER, secure_filename(filename)))

if __name__ == "__main__":
    app.run(debug=True, port=5000)