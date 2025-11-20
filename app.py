from flask import Flask, request, render_template, send_file, redirect, url_for
import pandas as pd
import fitz
import re
import io
import json

app = Flask(__name__)

# -----------------------
# Load reference database
# -----------------------
with open("reference.json", "r", encoding="utf-8") as f:
    REF = json.load(f)

ref_dict = {r["code"].strip(): r for r in REF if r.get("code")}


# -----------------------
# Extract qty even if code is in one line and qty is next line
# -----------------------
def extract_all_codes_and_quantities(text):
    lines = text.split("\n")
    results = {}

    for i, line in enumerate(lines):
        line = line.strip()

        # if line is exactly a code from reference.json
        if line in ref_dict:

            # look in the SAME line first
            qty = None
            m = re.search(r"\b(\d{1,7})\b", line)
            if m:
                qty = int(m.group(1))

            # if not found → look in the NEXT line
            if qty is None and i + 1 < len(lines):
                next_line = lines[i + 1]
                m2 = re.search(r"\b(\d{1,7})\b", next_line)
                if m2:
                    qty = int(m2.group(1))

            # If quantity found, add/update
            if qty is not None:
                if line not in results:
                    results[line] = 0
                results[line] += qty

    return results


# -----------------------
# MAIN ROUTE
# -----------------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":

        if "pdf" not in request.files:
            return redirect(url_for("index"))

        file = request.files["pdf"]
        pdf_bytes = file.read()

        # Read PDF text
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        text = ""
        for page in doc:
            text += page.get_text()

        # Extract
        raw_results = extract_all_codes_and_quantities(text)

        # Build table rows
        rows = []
        for code, qty in raw_results.items():

            item = ref_dict.get(code)
            if not item:
                continue

            pcs = int(item["pieces"])
            cartons = qty // pcs
            extra = qty - cartons * pcs
            gifts = cartons // 10

            rows.append({
                "code": code,
                "name": item["name"],
                "qty": qty,
                "pcs": pcs,
                "cartons": cartons,
                "extra": extra,
                "gifts": gifts,
                "status": "مستحق"
            })

        # Sort by code
        rows = sorted(rows, key=lambda x: x["code"])

        # Save Excel into memory buffer
        df = pd.DataFrame(rows)
        excel_bytes = io.BytesIO()
        df.to_excel(excel_bytes, index=False)
        excel_bytes.seek(0)

        # Render table in HTML
        return render_template("index.html", rows=rows)

    return render_template("index.html", rows=None)


# -----------------------
# DOWNLOAD EXCEL
# -----------------------
@app.route("/download")
def download_excel():
    # regenerate same Excel from last request
    # safer implementation: regenerate from reference + last cached pdf
    # but here we simply rebuild from request values passed via GET
    return "Not implemented yet"  # (we will fill this after table integration)


if __name__ == "__main__":
    app.run(debug=True)
