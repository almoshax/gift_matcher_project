# app.py
from flask import Flask, request, send_file, Response
import pandas as pd
import fitz
import re
import io
import json
import requests
import os
import time

app = Flask(__name__)

# -----------------------
# OCR.space API key (from environment fallback to provided key)
# -----------------------
OCR_API_KEY = os.getenv("OCR_API_KEY", "K86982355688957")

# -----------------------
# Load reference.json
# -----------------------
with open("reference.json", "r", encoding="utf-8") as f:
    REF = json.load(f)

ref_dict = {r["code"].strip(): r for r in REF if r.get("code")}

# -----------------------
# Helper: call OCR.space with file bytes and return extracted text
# -----------------------
def ocr_space_from_bytes(file_bytes, filename="file.pdf", language="eng"):
    """
    Sends bytes to OCR.Space and returns the parsed text.
    Using multipart upload (file).
    """
    url = "https://api.ocr.space/parse/image"
    headers = {
        # no special headers required for OCR.space multipart
    }
    data = {
        "apikey": OCR_API_KEY,
        "language": language,
        "isOverlayRequired": False,
        "OCREngine": 2,  # use faster engine 2 if available
        # for PDFs you can add "filetype": "PDF" but not required when uploading file
    }
    files = {
        "file": (filename, file_bytes, "application/pdf")
    }

    resp = requests.post(url, data=data, files=files, timeout=120)
    resp.raise_for_status()
    j = resp.json()
    # OCR.space returns ParsedResults array
    parsed = []
    if j.get("IsErroredOnProcessing"):
        # raise a helpful error
        msg = j.get("ErrorMessage") or j.get("ErrorDetails") or "OCR error"
        raise RuntimeError(f"OCR.space error: {msg}")
    for pr in j.get("ParsedResults") or []:
        parsed_text = pr.get("ParsedText") or ""
        parsed.append(parsed_text)
    # join all parsed results
    return "\n".join(parsed)

# -----------------------
# Extract codes & quantities (search only codes present in reference)
# This function: split OCR text by lines (keeps order) and for each line that equals
# a reference code we look for a number in same line or the next line.
# Also supports when OCR puts code+qty in same line.
# -----------------------
def extract_codes_and_quantities_from_text(text):
    lines = text.splitlines()
    results = {}

    for i, raw_line in enumerate(lines):
        line = raw_line.strip()
        if not line:
            continue

        # If the line contains the code exactly (best)
        if line in ref_dict:
            qty = None
            # check same line for number
            m = re.search(r"\b(\d{1,9})\b", line)
            if m:
                qty = int(m.group(1))

            # else check next few lines (sometimes OCR splits)
            if qty is None:
                look_ahead = 3  # check next up to 3 lines for quantity
                for k in range(1, look_ahead + 1):
                    if i + k < len(lines):
                        nl = lines[i + k].strip()
                        m2 = re.search(r"\b(\d{1,9})\b", nl)
                        if m2:
                            qty = int(m2.group(1))
                            break

            if qty is not None:
                results.setdefault(line, 0)
                results[line] += qty
            continue

        # fallback: sometimes OCR merges code with other words, search for any known code as token
        # Tokenize current line and check tokens
        tokens = re.findall(r"[A-Z0-9\-]{4,15}", line)
        for tok in tokens:
            if tok in ref_dict:
                # try to find number in same line after token
                after_pattern = re.escape(tok) + r".{0,60}?(\d{1,9})"
                m_after = re.search(after_pattern, line)
                if m_after:
                    qty = int(m_after.group(1))
                    results.setdefault(tok, 0)
                    results[tok] += qty
                else:
                    # look ahead next lines
                    look_ahead = 3
                    qty = None
                    for k in range(1, look_ahead + 1):
                        if i + k < len(lines):
                            nl = lines[i + k].strip()
                            m2 = re.search(r"\b(\d{1,9})\b", nl)
                            if m2:
                                qty = int(m2.group(1))
                                break
                    if qty is not None:
                        results.setdefault(tok, 0)
                        results[tok] += qty
    return results


# -----------------------
# MAIN: simple upload → OCR → match → Excel download
# -----------------------
@app.route("/", methods=["GET", "POST"])
def upload_and_process():
    if request.method == "GET":
        # Simple HTML to upload (keeps it minimal)
        return """
        <h3>ارفع ملف PDF (ممسوح سكانر)</h3>
        <form method='POST' enctype='multipart/form-data'>
            <input type='file' name='pdf' accept='application/pdf' required>
            <button type='submit'>ارفع وحَمّل Excel</button>
        </form>
        """

    # POST: handle file
    file = request.files.get("pdf")
    if not file:
        return "لم يتم ارسال ملف PDF", 400

    pdf_bytes = file.read()

    # Step 1: try quick text extraction (in case PDF has embedded text)
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        quick_text = ""
        for p in doc:
            quick_text += p.get_text()
    except Exception:
        quick_text = ""

    # If quick_text too small, call OCR
    use_ocr = True
    if quick_text and len(re.sub(r"\s+", "", quick_text)) > 30:
        # we have usable text — still we might want to run OCR for safety, but prefer quick_text
        text = quick_text
        use_ocr = False
    else:
        # run OCR
        try:
            # language param could be 'eng' or 'ara' — codes are alphanumeric, use 'eng' for best results
            text = ocr_space_from_bytes(pdf_bytes, filename=file.filename or "file.pdf", language="eng")
        except Exception as e:
            return f"OCR failed: {str(e)}", 500

    # Now extract codes & quantities using reference-codes-only logic
    extracted = extract_codes_and_quantities_from_text(text)

    # Build rows
    rows = []
    for code, qty in extracted.items():
        item = ref_dict.get(code)
        if not item:
            continue
        try:
            pcs = int(item["pieces"])
        except Exception:
            pcs = None
        cartons = qty // pcs if pcs else None
        extra = qty - cartons * pcs if pcs else None
        gifts = cartons // 10 if cartons is not None else None

        rows.append({
            "كود الصنف": code,
            "اسم الصنف": item.get("name", ""),
            "الكمية المطلوبة": qty,
            "عدد القطع في الكرتونة": pcs,
            "عدد الكراتين": cartons,
            "الزيادة": extra,
            "عدد الهدايا": gifts,
            "الحالة": "مستحق" if pcs else "غير داخل العرض"
        })

    # Sort rows by code
    rows = sorted(rows, key=lambda r: r["كود الصنف"])

    # If no rows found: return helpful message + OCR text sample to debug
    if not rows:
        sample = text[:2000].replace("\n", "<br>")
        return Response(f"لم يتم العثور على أي بنود متطابقة. نص الاستخراج (عينة):<br><div style='white-space:pre-wrap'>{sample}</div>", mimetype="text/html")

    # Build Excel and return
    df = pd.DataFrame(rows)
    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    # send as attachment
    filename = f"نتائج_المطابقة_{int(time.time())}.xlsx"
    return send_file(output, download_name=filename, as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
