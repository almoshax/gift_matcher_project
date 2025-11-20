from flask import Flask, request, send_file
import pandas as pd
import fitz
import re
import io
import json

app = Flask(__name__)

# Load reference
with open("reference.json", "r", encoding="utf-8") as f:
    REF = json.load(f)

ref_dict = {r["code"].strip(): r for r in REF if r.get("code")}


def extract_codes_and_quantities(text):
    lines = text.split("\n")
    results = {}

    for i, line in enumerate(lines):
        code = line.strip()

        if code in ref_dict:

            qty = None

            # try same line
            m = re.search(r"\b(\d+)\b", line)
            if m:
                qty = int(m.group(1))

            # try next line
            if qty is None and i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                m2 = re.search(r"\b(\d+)\b", next_line)
                if m2:
                    qty = int(m2.group(1))

            if qty is not None:
                if code not in results:
                    results[code] = 0
                results[code] += qty

    return results


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":

        file = request.files.get("pdf")
        if not file:
            return "ارفع ملف PDF يا معلم"

        pdf_bytes = file.read()

        # Read PDF
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        text = ""
        for page in doc:
            text += page.get_text()

        extracted = extract_codes_and_quantities(text)

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
        <button type='submit'>تشغيل</button>
    </form>
    """


if __name__ == "__main__":
    app.run()
