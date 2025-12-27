import os
import pandas as pd
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import zipfile

app = Flask(__name__)

# ---------------- Folders ----------------
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

ALLOWED_EXTENSIONS = {'xls', 'xlsx', 'csv'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ---------------- Certificate Generator ----------------
def generate_certificate(student_name, college_name, event_name, filename):
    width, height = landscape(A4)
    file_path = os.path.join(OUTPUT_FOLDER, filename)
    c = canvas.Canvas(file_path, pagesize=landscape(A4))

    # ---------------- Background ----------------
    background_path = r"C:\Users\sakth\OneDrive\Desktop\sbc certificate\IMG-20250926-WA0005[1].jpg"
    if os.path.exists(background_path):
        bg_img = ImageReader(background_path)
        iw, ih = bg_img.getSize()
        scale = min(width / iw, height / ih)
        new_w, new_h = iw * scale, ih * scale
        x = (width - new_w) / 2
        y = (height - new_h) / 2
        c.drawImage(bg_img, x, y, width=new_w, height=new_h, mask="auto")
    else:
        c.setFillColorRGB(0.95, 0.95, 1)
        c.rect(0, 0, width, height, fill=1)

    # ---------------- Student Name ----------------
    font_size_name = 18
    max_width = width - 100
    while c.stringWidth(student_name, "Times-Bold", font_size_name) > max_width and font_size_name > 12:
        font_size_name -= 1

    c.setFont("Times-Bold", font_size_name)
    c.setFillColorRGB(0, 0, 0)
    c.drawCentredString(width / 1.75, 455/ 2 + 50, student_name)

    # ---------------- Event Name ----------------
    if event_name:
        font_size_event = 20
        max_width_event = width - 150
        while c.stringWidth(event_name, "Times-Bold", font_size_event) > max_width_event and font_size_event > 12:
            font_size_event -= 1
        c.setFont("Times-Bold", font_size_event)
        c.setFillColorRGB(0, 0, 0)
        c.drawCentredString(width / 4, 400 / 2 + 5, event_name)

    # ---------------- College Name (Right-Aligned) ----------------
    if college_name:
        font_size_college = 30
        max_width_college = width / 2 - 50
        while c.stringWidth(college_name, "Times-Bold", font_size_college) > max_width_college and font_size_college > 10:
            font_size_college -= 1

        c.setFont("Times-Bold", font_size_college)
        c.setFillColorRGB(0, 0, 0)
        x_pos = width - 350
        y_pos = height / 2 - 50
        c.drawRightString(x_pos, y_pos, f"{college_name}")

    # ---------------- Signatures ----------------
    sig_y = 100  # vertical position for signatures
    sig_width, sig_height = 150, 60  # adjust size

    try:
        # HOD Signature (left)
        hod_path = r""
        if os.path.exists(hod_path):
            c.drawImage(hod_path, width/4 - 145, sig_y - 30, width=140, height=20, mask='auto')
        c.setFont("Times-Roman", 12)
        c.drawCentredString(width/2, sig_y - 50, "")

        # Principal Signature (center)
        principal_path = r""
        if os.path.exists(principal_path):
            c.drawImage(principal_path, width/2 - 200/2, sig_y - 30, width=150, height=20, mask='auto')
        c.setFont("Times-Roman", 12)
        c.drawCentredString(width/2, sig_y - 20, "")

        # Director Signature (right)
        director_path = r""
        if os.path.exists(director_path):
            c.drawImage(director_path, 3*width/4 - 55, sig_y - 30, width=sig_width, height=30, mask='auto')
        c.setFont("Times-Roman", 12)
        c.drawCentredString(3*width/4, sig_y - 20, "")

    except Exception as e:
        print(f"Signature error: {e}")

    c.save()
    return file_path

# ---------------- Flask Route ----------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        college = request.form.get("college", "").strip()
        event = request.form.get("event", "").strip()
        uploaded_file = request.files.get("file")
        pdf_files = []

        try:
            # Single name + college + event
            if name:
                pdf_name = f"{name}_certificate.pdf"
                pdf_path = generate_certificate(name, college, event, pdf_name)
                pdf_files.append(pdf_path)

            # Excel/CSV upload
            elif uploaded_file and allowed_file(uploaded_file.filename):
                filename = secure_filename(uploaded_file.filename)
                filepath = os.path.join(UPLOAD_FOLDER, filename)
                uploaded_file.save(filepath)

                ext = filename.rsplit('.', 1)[1].lower()
                if ext in ['xls', 'xlsx']:
                    df = pd.read_excel(filepath)
                else:
                    df = pd.read_csv(filepath)

                if 'Name' not in df.columns:
                    return render_template("index.html", error="Excel/CSV must have a 'Name' column.")

                for idx, row in df.iterrows():
                    n_str = str(row['Name']).strip()
                    c_str = str(row['College']).strip() if 'College' in df.columns else college
                    e_str = str(row['Event']).strip() if 'Event' in df.columns else event
                    if n_str:
                        pdf_name = f"{n_str}_certificate.pdf"
                        pdf_path = generate_certificate(n_str, c_str, e_str, pdf_name)
                        pdf_files.append(pdf_path)
            else:
                return render_template("index.html", error="Please enter a name or upload a valid Excel/CSV file.")

            # Zip multiple PDFs
            if len(pdf_files) > 1:
                zip_path = os.path.join(OUTPUT_FOLDER, "certificates.zip")
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for pdf in pdf_files:
                        zipf.write(pdf, os.path.basename(pdf))
                return send_file(zip_path, as_attachment=True)

            elif len(pdf_files) == 1:
                return send_file(pdf_files[0], as_attachment=True)

            else:
                return render_template("index.html", error="No certificates generated.")

        except Exception as e:
            return render_template("index.html", error=f"Error: {e}")

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
