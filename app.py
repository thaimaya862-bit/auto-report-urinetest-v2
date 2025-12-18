import os
import io
import re
import datetime
import subprocess
from pathlib import Path

from flask import Flask, render_template, request, send_from_directory
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import pdfplumber
from PIL import Image

try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

# Optional renderer (more reliable for "PDF -> image")
try:
    import pypdfium2 as pdfium
except Exception:
    pdfium = None

app = Flask(__name__)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
TEMPLATE_DIR = os.path.join(BASE_DIR, "word_templates")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

THAI_MONTHS = [
    "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]

def format_thai_date_full(d: datetime.date) -> str:
    # รูปแบบตามที่สั่ง: 16 ธันวาคม พ.ศ.2568
    return f"{d.day} {THAI_MONTHS[d.month]} พ.ศ.{d.year + 543}"

def guess_gender_from_fullname(fullname):
    if not fullname:
        return ""
    prefix = fullname.split()[0]
    male_prefixes = ["นาย", "ด.ช.", "เด็กชาย"]
    female_prefixes = ["นาง", "นางสาว", "ด.ญ.", "เด็กหญิง", "น.ส."]
    if prefix in male_prefixes:
        return "ชาย"
    if prefix in female_prefixes:
        return "หญิง"
    police_prefixes = [
        "ร.ต.อ.", "ร.ต.ท.", "ร.ต.ต.", "ด.ต.", "ส.ต.อ.", "ส.ต.ท.", "ส.ต.ต.",
        "พ.ต.อ.", "พ.ต.ท.", "พ.ต.ต.", "พ.ต.", "พล.ต.ต.", "พล.ต.อ.", "จ.ส.ต."
    ]
    if any(prefix.startswith(p) for p in police_prefixes):
        return "ชาย"
    return ""

TEAMS = {
    "1": {
        "name": "ชุดที่ 1",
        "leader": "ร.ต.อ.พิชิต  พัฒนาศูร",
        "leader_phone": "062-108-4116",
        "members": [
            "ร.ต.ต.ณรงค์  บุตรพรม",
            "ด.ต.อดุลย์  ธงศรี",
            "ส.ต.ท.ชนาธิป  ประหา",
        ],
    },
    "2": {
        "name": "ชุดที่ 2",
        "leader": "ร.ต.อ.สัญปกรณ์  นครเพชร",
        "leader_phone": "085-123-3219",
        "members": [
            "ร.ต.อ.สายสิทธิ์  มีศักดิ์",
            "ด.ต.วุฒินันต์  ประเสริฐสังข์",
            "ด.ต.จักรพันธ์  โพธิ์ศรีศาสตร์",
        ],
    },
    "3": {
        "name": "ชุดที่ 3",
        "leader": "ร.ต.อ.ปัญญา  วรรณชาติ",
        "leader_phone": "094-157-4741",
        "members": [
            "ร.ต.ท.ศักดิ์ศรี  สรรพวุธ",
            "ร.ต.ต.ปราศภัยพาล  แก้วทรายขาว",
            "ส.ต.ท.อนิรุทธ์  ทุหา",
        ],
    },
}


def parse_pdf_register(file_stream):
    result = {
        "FULLNAME": "",
        "CID": "",
        "DOB": "",
        "AGE": "",
        "HOUSE_NO": "",
        "MOO": "",
        "TAMBON": "",
        "AMPHUR": "",
        "PROVINCE": "",
        "MOVEIN_DATE": "",
    }
    with pdfplumber.open(file_stream) as pdf:
        text = ""
        for page in pdf.pages:
            t = page.extract_text() or ""
            text += t + "\n"
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    for line in lines:
        if "เลขประจำตัวประชาชน" in line:
            parts = line.replace("เลขประจำตัวประชาชน", " ").split()
            for p in parts:
                if any(ch.isdigit() for ch in p):
                    result["CID"] = p.strip()
                    break
    for line in lines:
        if "ชื่อ-ชื่อสกุล" in line:
            try:
                after = line.split("ชื่อ-ชื่อสกุล", 1)[1].strip()
                for cut in ["เพศ", "วันเดือนปีเกิด", "อายุ"]:
                    if cut in after:
                        after = after.split(cut, 1)[0].strip()
                result["FULLNAME"] = after
            except Exception:
                pass
    for line in lines:
        if "วันเดือนปีเกิด" in line:
            try:
                part = line.split("วันเดือนปีเกิด", 1)[1].strip()
                if "อายุ" in part:
                    dob_text, age_part = part.split("อายุ", 1)
                    result["DOB"] = dob_text.strip()
                    age_num = "".join(ch for ch in age_part if ch.isdigit())
                    result["AGE"] = age_num
            except Exception:
                pass
    for line in lines:
        if "บ้านเลขที่" in line:
            parts = line.split()
            for i, p in enumerate(parts):
                if p.startswith("บ้านเลขที่") and i + 1 < len(parts):
                    result["HOUSE_NO"] = parts[i + 1]
        if "หมู่" in line and "บ้านเลขที่" in line:
            parts = line.split()
            for i, p in enumerate(parts):
                if p.startswith("หมู่") and i + 1 < len(parts):
                    result["MOO"] = parts[i + 1]
    for line in lines:
        if "ตำบล" in line and not result["TAMBON"]:
            parts = line.split()
            for i, p in enumerate(parts):
                if p.startswith("ตำบล") and i + 1 < len(parts):
                    result["TAMBON"] = parts[i + 1]
        if "อำเภอ" in line and not result["AMPHUR"]:
            parts = line.split()
            for i, p in enumerate(parts):
                if p.startswith("อำเภอ") and i + 1 < len(parts):
                    result["AMPHUR"] = parts[i + 1]
        if "จังหวัด" in line and not result["PROVINCE"]:
            parts = line.split()
            for i, p in enumerate(parts):
                if p.startswith("จังหวัด") and i + 1 < len(parts):
                    result["PROVINCE"] = parts[i + 1]
    for line in lines:
        if "วันที่ย้ายเข้า" in line:
            try:
                result["MOVEIN_DATE"] = line.split("วันที่ย้ายเข้า", 1)[1].strip()
            except Exception:
                pass
    addr_parts = []
    if result.get("HOUSE_NO"):
        addr_parts.append(result["HOUSE_NO"])
    if result.get("MOO"):
        addr_parts.append("หมู่ " + result["MOO"])
    if result.get("TAMBON"):
        addr_parts.append("ตำบล " + result["TAMBON"])
    if result.get("AMPHUR"):
        addr_parts.append("อำเภอ " + result["AMPHUR"])
    if result.get("PROVINCE"):
        addr_parts.append("จังหวัด " + result["PROVINCE"])
    result["ADDRESS_FULL"] = " ".join(addr_parts)
    result["GENDER"] = guess_gender_from_fullname(result.get("FULLNAME", ""))
    return result


def safe_filename(prefix: str, ext: str) -> str:
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{prefix}_{ts}.{ext}"

def pdf_first_page_to_pil(file_stream) -> Image.Image:
    # Return PIL image (RGB) of first page.
    # 1) Try pdfplumber's to_image (needs imagemagick sometimes)
    try:
        file_stream.seek(0)
        with pdfplumber.open(file_stream) as pdf:
            if not pdf.pages:
                raise ValueError("Empty PDF")
            page = pdf.pages[0]
            try:
                im = page.to_image(resolution=160).original  # PIL
                return im.convert("RGB")
            except Exception:
                # fallback: try rendering via pdfium
                pass
    except Exception:
        pass

    # 2) Try pdfium (preferred on Linux if installed)
    if pdfium is None:
        raise RuntimeError("PDF->Image renderer not available (pypdfium2 not installed)")
    file_stream.seek(0)
    data = file_stream.read()
    pdf = pdfium.PdfDocument(data)
    if len(pdf) < 1:
        raise ValueError("Empty PDF")
    page = pdf[0]
    # scale for readability
    bitmap = page.render(scale=2.0)
    pil = bitmap.to_pil()
    return pil.convert("RGB")

def inline_image_from_pil(doc: DocxTemplate, img: Image.Image, width_mm: int = 150) -> InlineImage:
    bio = io.BytesIO()
    img.save(bio, format="PNG", optimize=True)
    bio.seek(0)
    return InlineImage(doc, bio, width=Mm(width_mm))

def convert_docx_to_pdf(docx_path: str, out_dir: str) -> str | None:
    # Returns pdf filepath if successful else None
    try:
        if docx2pdf_convert is not None:
            # docx2pdf on Windows/mac expects output dir or file; handle both
            docx2pdf_convert(docx_path, out_dir)
            pdf_name = Path(docx_path).with_suffix(".pdf").name
            pdf_path = os.path.join(out_dir, pdf_name)
            if os.path.exists(pdf_path):
                return pdf_path
    except Exception:
        pass

    # Linux fallback: libreoffice (if available)
    try:
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        pdf_name = Path(docx_path).with_suffix(".pdf").name
        pdf_path = os.path.join(out_dir, pdf_name)
        if os.path.exists(pdf_path):
            return pdf_path
    except Exception:
        return None
    return None

def template_for_count(n: int) -> str:
    n = max(1, min(6, n))
    return os.path.join(TEMPLATE_DIR, f"main_template_urinetest{n}.docx")

def suffix_letter(i: int) -> str:
    # 0->A ... 5->F
    return chr(ord("A") + i)

def build_context_for_person(person: dict, letter: str) -> dict:
    # Map parsed fields -> suffixed variables (A..F)
    ctx = {}
    # base keys we support (ยึด v2)
    keys = [
        "FULLNAME", "CID", "DOB", "AGE",
        "HOUSE_NO", "MOO", "TAMBON", "AMPHUR", "PROVINCE",
        "ADDRESS_FULL", "MOVEIN_DATE"
    ]
    for k in keys:
        ctx[f"{k}_{letter}"] = person.get(k, "") or ""
    return ctx

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template("index.html", docx_filename=None, pdf_filename=None)

    # date
    doc_date_raw = (request.form.get("doc_date") or "").strip()
    if doc_date_raw:
        # expects YYYY-MM-DD
        try:
            y, m, d = [int(x) for x in doc_date_raw.split("-")]
            doc_date = datetime.date(y, m, d)
        except Exception:
            doc_date = datetime.date.today()
    else:
        doc_date = datetime.date.today()

    time_start = (request.form.get("time_start") or "").strip()

    # v2: รับไฟล์ทะเบียนราษฎร์แบบแยกช่อง 1-6 และเรียงตามลำดับช่อง (ข้ามช่องว่าง)
    pdf_files = []
    for i in range(1, 7):
        f = request.files.get(f"pdf_{i}")
        if f and f.filename:
            pdf_files.append(f)

    if not pdf_files:
        return render_template("index.html", docx_filename=None, pdf_filename=None)

    template_path = template_for_count(len(pdf_files))
    doc = DocxTemplate(template_path)

    # Build context
    ctx = {}
    ctx["DOC_DATE"] = format_thai_date_full(doc_date)
    ctx["TIME_START"] = time_start

    # For each person
    for idx, fs in enumerate(pdf_files):
        letter = suffix_letter(idx)
        # Parse text fields using v2 logic
        fs.stream.seek(0)
        person = parse_pdf_register(fs.stream)
        ctx.update(build_context_for_person(person, letter))

        # Insert house register image (PDF first page)
        try:
            fs.stream.seek(0)
            pil = pdf_first_page_to_pil(fs.stream)
            ctx[f"HOUSE_REG_IMAGE_{letter}"] = inline_image_from_pil(doc, pil, width_mm=150)
        except Exception:
            # If rendering fails, leave blank (do not crash)
            ctx[f"HOUSE_REG_IMAGE_{letter}"] = ""

    # Render
    out_docx = safe_filename("urinetest", "docx")
    out_docx_path = os.path.join(OUTPUT_DIR, out_docx)
    doc.render(ctx)
    doc.save(out_docx_path)

    # Convert to PDF (best effort)
    pdf_path = convert_docx_to_pdf(out_docx_path, OUTPUT_DIR)
    out_pdf = os.path.basename(pdf_path) if pdf_path else None

    return render_template("index.html", docx_filename=out_docx, pdf_filename=out_pdf)

@app.route("/download/<path:filename>")
def download_file(filename):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
