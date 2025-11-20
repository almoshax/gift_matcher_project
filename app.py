from flask import Flask, request, render_template, send_file, redirect, url_for
import pandas as pd
import fitz
import re
import io
import json

app = Flask(__name__)

# -----------------------
# Load reference.json
# -----------------------
with open("reference.json", "r", encoding="utf-8") as f:
    REF = json.load(f)

ref_dict = {r["code"].strip(): r for r in REF if r.get("code")}

# -----------------------
# GLOBAL VARIABLE to store last results
# -----------------------
LAST_RESULTS = None


# -----------------------
# Extract qty even if code on one line and qty on next
# -----------------------
def extract_all_codes_and_quantities(text):
    lines = text.split("\n")
    results = {}

    for i, line in enumerate(lines):
        line = line.strip()

        # exactly matches code
        if line in ref_dict:

            qty = None

            # try same line
            m = re.search(r"\b(\d{1,7})\b", line)
            if m:
                qty = int(m.group(1))

            # try next line
            if qty is None and i + 1 < len(lines):
                next_line = lines[i + 1]
                m2 = re.search(r"\b(\d{1,7})\b", next_line)
                if m2:
                    qty = int(m2.group(1))

            if qty is not None:
                if line not in results:
                    results[line] = 0
                results[line] += qty

    return results


# -----------------------
# MAIN PAGE
# -----------------------
@app.route("/", methods=["GET", "POST"])
def index():
    global LAST_RESULTS

    if request.method == "POST":

        file = request.files.get("pdf")
        if not file:
            return redirect(url_for("index"))

        pdf_bytes = file.read()

        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        text = ""
        for page in doc:
            text += page.get_text()

        raw_results = extract_all_codes_and_quantities(text)

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

        rows = sorted(rows, key=lambda x: x["code"])

        # Save results for Excel
        LAST_RESULTS = rows

        return render_template("index.html", rows=rows)

    return render_template("index.html", rows=None)


# -----------------------
# DOWNLOAD EXCEL
# -----------------------
@app.route("/download")
def download_excel():
    global LAST_RESULTS

    if not LAST_RESULTS:
        return "لا توجد نتائج مطابقة بعد."

    df = pd.DataFrame(LAST_RESULTS)

    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return send_file(
        output,
        download_name="نتائج_المطابقة.xlsx",
        as_attachment=True
    )


if __name__ == "__main__":
    app.run(debug=True)
