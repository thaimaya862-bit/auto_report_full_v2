
import os
import io
import datetime
from datetime import timedelta

from flask import Flask, render_template, request, send_from_directory, url_for
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import pdfplumber
from PIL import Image
from docx2pdf import convert as docx2pdf_convert

app = Flask(__name__)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "word_templates", "main_template.docx")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

THAI_DIGITS_MAP = str.maketrans("0123456789", "๐๑๒๓๔๕๖๗๘๙")

def to_thai_num(value):
    if value is None:
        return ""
    return str(value).translate(THAI_DIGITS_MAP)

def format_thai_date(date_obj):
    months = [
        "",
        "มกราคม",
        "กุมภาพันธ์",
        "มีนาคม",
        "เมษายน",
        "พฤษภาคม",
        "มิถุนายน",
        "กรกฎาคม",
        "สิงหาคม",
        "กันยายน",
        "ตุลาคม",
        "พฤศจิกายน",
        "ธันวาคม",
    ]
    year_be = date_obj.year + 543
    text = f"{date_obj.day} {months[date_obj.month]} พ.ศ. {year_be}"
    return to_thai_num(text)

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

def build_photo_grid(doc, image_files):
    imgs = []
    for fs in image_files:
        if fs and fs.filename:
            try:
                img = Image.open(fs.stream).convert("RGB")
                imgs.append(img)
            except Exception:
                continue
    if not imgs:
        return None
    while len(imgs) < 4:
        imgs.append(imgs[-1])
    thumb_w, thumb_h = 600, 800
    thumbs = []
    for img in imgs[:4]:
        im = img.copy()
        im.thumbnail((thumb_w, thumb_h))
        thumbs.append(im)
    grid_w = thumb_w * 2
    grid_h = thumb_h * 2
    grid = Image.new("RGB", (grid_w, grid_h), "white")
    positions = [(0, 0), (thumb_w, 0), (0, thumb_h), (thumb_w, thumb_h)]
    for pos, t in zip(positions, thumbs):
        grid.paste(t, pos)
    bio = io.BytesIO()
    grid.save(bio, format="JPEG")
    bio.seek(0)
    return InlineImage(doc, bio, width=Mm(120))

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template("index.html", teams=TEAMS, docx_filename=None, pdf_filename=None)
    doc_date_str = request.form.get("doc_date")
    time_start_str = request.form.get("time_start")
    team_id = request.form.get("team_id")
    pdf_file = request.files.get("pdf_register")
    photo1 = request.files.get("photo1")
    photo2 = request.files.get("photo2")
    photo3 = request.files.get("photo3")
    photo4 = request.files.get("photo4")
    if not doc_date_str:
        return "กรุณาเลือกวันที่", 400
    if not time_start_str:
        return "กรุณากรอกเวลาเริ่ม", 400
    if not team_id or team_id not in TEAMS:
        return "กรุณาเลือกชุดเจ้าหน้าที่", 400
    if not pdf_file or not pdf_file.filename.lower().endswith(".pdf"):
        return "กรุณาอัปโหลดไฟล์ทะเบียนราษฎร (PDF)", 400
    try:
        doc_date_obj = datetime.datetime.strptime(doc_date_str, "%Y-%m-%d").date()
    except Exception:
        return "รูปแบบวันที่ไม่ถูกต้อง", 400
    try:
        dt_base = datetime.datetime.combine(
            doc_date_obj,
            datetime.datetime.strptime(time_start_str, "%H:%M").time(),
        )
        dt_end = dt_base + timedelta(hours=1, minutes=15)
        time_start_output = dt_base.strftime("%H.%M")
        time_end_output = dt_end.strftime("%H.%M")
    except Exception:
        time_start_output = time_start_str.replace(":", ".")
        time_end_output = ""
    pdf_bytes = pdf_file.read()
    pdf_stream = io.BytesIO(pdf_bytes)
    person = parse_pdf_register(pdf_stream)
    team = TEAMS[team_id]
    members = team["members"]
    team_members_str = ", ".join(members)
    doc = DocxTemplate(TEMPLATE_PATH)
    photo_grid = build_photo_grid(doc, [photo1, photo2, photo3, photo4])
    gender = person.get("GENDER", "")
    context = {
        "DOC_DATE": format_thai_date(doc_date_obj),
        "TIME_START": to_thai_num(time_start_output),
        "TIME_END": to_thai_num(time_end_output),
        "FULLNAME": to_thai_num(person.get("FULLNAME", "")),
        "CID": to_thai_num(person.get("CID", "")),
        "DOB": to_thai_num(person.get("DOB", "")),
        "AGE": to_thai_num(person.get("AGE", "")),
        "HOUSE_NO": to_thai_num(person.get("HOUSE_NO", "")),
        "MOO": to_thai_num(person.get("MOO", "")),
        "TAMBON": person.get("TAMBON", ""),
        "AMPHUR": person.get("AMPHUR", ""),
        "PROVINCE": person.get("PROVINCE", ""),
        "MOVEIN_DATE": to_thai_num(person.get("MOVEIN_DATE", "")),
        "ADDRESS_FULL": to_thai_num(person.get("ADDRESS_FULL", "")),
        "GENDER": gender,
        "SEX": gender,
        "TEAM_NAME": team["name"],
        "TEAM_LEADER": team["leader"],
        "TEAM_LEADER_PHONE": to_thai_num(team["leader_phone"]),
        "TEAM_MEMBER1": members[0] if len(members) > 0 else "",
        "TEAM_MEMBER2": members[1] if len(members) > 1 else "",
        "TEAM_MEMBER3": members[2] if len(members) > 2 else "",
        "TEAM_MEMBERS": team_members_str,
        "PHOTO_GRID": photo_grid,
    }
    doc.render(context)
    fullname_for_file = person.get("FULLNAME", "").strip()
    if not fullname_for_file:
        fullname_for_file = "ไม่ทราบชื่อ"
    safe_name = "".join(ch if ch not in r'\/:*?"<>|' else "_" for ch in fullname_for_file)
    safe_name = safe_name.replace(" ", "_")
    filename_docx = f"บันทึกจับกุม_{safe_name}.docx"
    output_docx_path = os.path.join(OUTPUT_DIR, filename_docx)
    doc.save(output_docx_path)
    filename_pdf = None
    try:
        pdf_path = os.path.join(OUTPUT_DIR, filename_docx.replace(".docx", ".pdf"))
        docx2pdf_convert(output_docx_path, pdf_path)
        filename_pdf = os.path.basename(pdf_path)
    except Exception:
        filename_pdf = None
    return render_template(
        "index.html",
        teams=TEAMS,
        docx_filename=filename_docx,
        pdf_filename=filename_pdf,
    )

@app.route("/download/<path:filename>")
def download_file(filename):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
