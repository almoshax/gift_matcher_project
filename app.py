from flask import Flask, request, render_template, send_file
import pandas as pd
import fitz, re, json, io

app = Flask(__name__)

with open("reference.json","r",encoding="utf-8") as f:
    REF=json.load(f)
ref_dict={r["code"]:r for r in REF}

def extract_qtys(text, code):
    out=[]
    for m in re.finditer(re.escape(code), text):
        after=text[m.end(): m.end()+120]
        m2=re.search(r"\b(\d{1,7})\b", after)
        if m2: out.append(int(m2.group(1)))
    return out

@app.route("/", methods=["GET","POST"])
def index():
    if request.method=="POST":
        file=request.files["pdf"]
        buf=file.read()
        doc=fitz.open(stream=buf, filetype="pdf")
        txt=""
        for p in doc:
            txt+=p.get_text()
        T=re.sub(r"\s+"," ",txt)
        rows=[]
        for code,item in ref_dict.items():
            qs=extract_qtys(T, code)
            if qs:
                qty=sum(qs)
                pcs=item["pieces"]
                cartons=qty//pcs
                extra=qty - cartons*pcs
                gifts=cartons//10
                rows.append(dict(code=code,name=item["name"],qty=qty,pcs=pcs,
                                 cartons=cartons,extra=extra,gifts=gifts,status="مستحق"))
        df=pd.DataFrame(rows)
        x=io.BytesIO()
        df.to_excel(x,index=False)
        x.seek(0)
        return send_file(x, download_name="نتائج_المطابقة.xlsx", as_attachment=True)
    return render_template("index.html")
