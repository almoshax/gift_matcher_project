import io
import json
from flask import Flask, request, send_file
import pandas as pd
from google.cloud import vision_v1
from google.oauth2 import service_account
import fitz

app = Flask(__name__)

# --------------------------
# Load reference
# --------------------------
with open("reference.json", "r", encoding="utf-8") as f:
    REF = json.load(f)

ref_dict = {r["code"].strip(): r for r in REF if r.get("code")}


# --------------------------
# Initialize Google Vision
# --------------------------
def init_vision():
    creds = service_account.Credentials.from_service_account_file(
        "gcloud_key.json"
    )
    client = vision_v1.ImageAnnotatorClient(credentials=creds)
    return client


# --------------------------
# OCR PDF using Google Vision (batchAnnotateFiles)
# --------------------------
def ocr_pdf(pdf_bytes):
    client = init_vision()

    mime_type = "application/pdf"

    file_input = vision_v1.types.InputConfig(
        content=pdf_bytes, mime_type=mime_type
    )

    feature = vision_v1.types.Feature(
        type=vision_v1.Feature.Type.DOCUMENT_TEXT_DETECTION
    )

    request = vision_v1.types.AsyncAnnotateFileRequest(
        features=[feature], input_config=file_input
    )

    operation = client.async_batch_annotate_files(requests=[request])
    response = operation.result(timeout=300)

    text = ""
    for r in response.responses:
        for page in r.responses:
            if page.full_text_annotation:
                text += page.full_text_annotation.text + "\n"

    return text


# --------------------------
# Extract codes & quantities
# --------------------------
def extract_codes(text):
    lines = text.split("\n")
    results = {}

    for i, line in enumerate(lines):
        line = line.strip()

        if line in ref_dict:
            qty = None

            m = fitz.re.search(r"\b(\d+)\b", line)
            if m:
                qty = int(m.group(1))

            if qty is None and i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                m2 = fitz.re.search(r"\b(\d+)\b", next_line)
                if m2:
                    qty = int(m2.group(1))

            if qty:
                results[line] = results.get(line, 0) + qty

    return results


# --------------------------
# Flask MAIN
# --------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":

        file = request.files.get("pdf")
        if not file:
            return "ارفع PDF"

        pdf_bytes = file.read()

        # OCR Google Vision
        text = ocr_pdf(pdf_bytes)

        # Extract codes + quantities
        extracted = extract_codes(text)

        rows = []
        for code, qty in extracted.items():

            item = ref_dict.get(code)
            if not item:
                continue

            pcs = int(item["pieces"])
            cartons = qty // pcs
            extra = qty - cartons * pcs
            gifts = cartons // 10

            rows.append({
                "كود الصنف": code,
                "اسم الصنف": item["name"],
                "الكمية المطلوبة": qty,
                "عدد القطع في الكرتونة": pcs,
                "عدد الكراتين": cartons,
                "الزيادة": extra,
                "عدد الهدايا": gifts,
                "الحالة": "مستحق",
            })

        df = pd.DataFrame(rows)
        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        return send_file(
            output,
            download_name="نتيجة_المطابقة.xlsx",
            as_attachment=True
        )

    return """
    <h1>Upload PDF</h1>
    <form method='POST' enctype='multipart/form-data'>
      <input type='file' name='pdf' accept='application/pdf'>
      <button type='submit'>Run</button>
    </form>
    """
