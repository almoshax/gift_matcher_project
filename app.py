import io
import json
import requests
from flask import Flask, request, send_file
import pandas as pd
import re

app = Flask(__name__)

# -------------------------
# Load reference
# -------------------------
with open("reference.json", "r", encoding="utf-8") as f:
    REF = json.load(f)

ref_dict = {r["code"].strip(): r for r in REF if r.get("code")}


# -------------------------
# OCR via free public API (no API key needed)
# -------------------------
def ocr_public(pdf_bytes):
    url = "https://api.ocr.space/parse/image"

    files = {
        "file": ("upload.pdf", pdf_bytes)
    }

    data = {
        "language": "eng",
        "isOverlayRequired": False,
        "OCREngine": 2,
    }

    r = requests.post(url, files=files, data=data, timeout=200)
    r.raise_for_status()
    j = r.json()

    if j.get("IsErroredOnProcessing"):
        raise RuntimeError("OCR error")

    text = ""
    for block in j.get("ParsedResults", []):
        text += block.get("ParsedText", "") + "\n"

    return text


# -------------------------
# Extract codes & quantities
# -------------------------
def extract_codes(text):
    lines = text.split("\n")
    results = {}

    for i, line in enumerate(lines):
        line = line.strip()

        if line in ref_dict:

            qty = None

            # search on same line
            m = re.search(r"\b(\d+)\b", line)
            if m:
                qty = int(m.group(1))

            # search on next line
            if qty is None and i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                m2 = re.search(r"\b(\d+)\b", next_line)
                if m2:
                    qty = int(m2.group(1))

            if qty:
                results[line] = results.get(line, 0) + qty

    return results


# -------------------------
# Main route
# -------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":

        file = request.files.get("pdf")
        if not file:
            return "ارفع PDF"

        pdf_bytes = file.read()

        # OCR
        text = ocr_public(pdf_bytes)

        # extract matched codes
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
                "الحالة": "مستحق"
            })

        df = pd.DataFrame(rows)
        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        return send_file(
            output,
            download_name="matching_result.xlsx",
            as_attachment=True
        )

    return """
    <h2>ارفع ملف PDF</h2>
    <form method='POST' enctype='multipart/form-data'>
        <input type='file' name='pdf' accept='application/pdf'>
        <button type='submit'>تشغيل OCR</button>
    </form>
    """


if __name__ == "__main__":
    app.run(debug=True)
