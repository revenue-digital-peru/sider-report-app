"""
Sider Express — Report Web App v4 (Memory Optimized)
"""
import os, io, json, re, gc
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/health")
def health():
    return jsonify({"status": "ok"})

@app.route("/modelos")
def listar_modelos():
    try:
        from google import genai
        client = genai.Client(api_key=GEMINI_API_KEY)
        nombres = [m.name for m in client.models.list()]
        return jsonify({"modelos": nombres})
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route("/generate", methods=["POST"])
def generate():
    try:
        if "xlsx_actual" not in request.files:
            return jsonify({"error": "Falta el XLSX de la semana actual."}), 400
        if "xlsx_anterior" not in request.files:
            return jsonify({"error": "Falta el XLSX de la semana anterior."}), 400
        if "pptx" not in request.files:
            return jsonify({"error": "Falta el archivo PPTX plantilla."}), 400

        config = {
            "mes":            request.form.get("mes", ""),
            "periodo_inicio": request.form.get("periodo_inicio", ""),
            "periodo_fin":    request.form.get("periodo_fin", ""),
            "semana":         request.form.get("semana", ""),
        }

        xlsx_actual_bytes   = request.files["xlsx_actual"].read()
        xlsx_anterior_bytes = request.files["xlsx_anterior"].read()
        pptx_bytes          = request.files["pptx"].read()
        gc.collect()

        data_actual = extract_minimal_data(xlsx_actual_bytes)
        del xlsx_actual_bytes; gc.collect()

        data_anterior = extract_minimal_data(xlsx_anterior_bytes)
        del xlsx_anterior_bytes; gc.collect()

        variaciones = calcular_variaciones(data_actual, data_anterior)
        del data_anterior; gc.collect()

        texts = generate_texts_with_gemini(data_actual, variaciones, config)
        gc.collect()

        replacements = build_replacements(data_actual, variaciones, texts, config)
        del data_actual, variaciones, texts; gc.collect()

        output_bytes = fill_pptx(pptx_bytes, replacements)
        del pptx_bytes, replacements; gc.collect()

        mes = config.get("mes", "reporte").upper()
        pi  = config.get("periodo_inicio", "").replace("/", "")
        pf  = config.get("periodo_fin", "").replace("/", "")

        return send_file(
            io.BytesIO(output_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=f"Sider_Reporte_{mes}_{pi}_{pf}.pptx"
        )
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


def extract_minimal_data(xlsx_bytes):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True, read_only=True)
    sheets = wb.sheetnames

    def read_sheet(name, max_row=25):
        if name not in sheets:
            return []
        rows = []
        for i, row in enumerate(wb[name].iter_rows(min_row=1, max_row=max_row, values_only=True)):
            if any(v is not None for v in row):
                rows.append(list(row))
        return rows

    def find_header(rows):
        for i, row in enumerate(rows):
            if row and any("plataforma" in str(v).lower() for v in row if v):
                return i
        return 0

    def to_dicts(rows, hi=0):
        if not rows or len(rows) <= hi:
            return []
        headers = [str(h).strip() if h is not None else f"c{i}" for i, h in enumerate(rows[hi])]
        return [dict(zip(headers, list(row) + [None]*(len(headers)-len(row))))
                for row in rows[hi+1:] if any(v is not None for v in row)]

    pres_r     = read_sheet("PRESUPUESTO", 12)
    total_r    = read_sheet("TOTAL", 25)
    reg_comp_r = read_sheet("CAMPAÑAS REGULARES COMPARATIVA ", 15)
    reg_met_r  = read_sheet("CAMPAÑAS REGULARES METRICAS SEC", 15)
    inc_comp_r = read_sheet("CAMPAÑAS INCENTIVOS COMPARATIVA", 10)
    inc_met_r  = read_sheet("CAMPAÑAS INCENTIVOS METRICAS SE", 10)

    wb.close()
    del wb; gc.collect()

    return {
        "pres":      to_dicts(pres_r,     find_header(pres_r)),
        "total_raw": total_r,
        "reg_comp":  to_dicts(reg_comp_r, find_header(reg_comp_r)),
        "reg_met":   to_dicts(reg_met_r,  find_header(reg_met_r)),
        "inc_comp":  to_dicts(inc_comp_r, find_header(inc_comp_r)),
        "inc_met":   to_dicts(inc_met_r,  find_header(inc_met_r)),
    }


def to_float(val):
    if val is None: return None
    try: return float(str(val).replace(",","").replace("$","").replace("%","").replace(" ","").strip())
    except: return None

def calc_var(a, b):
    fa, fb = to_float(a), to_float(b)
    if fa is None or fb is None or fb == 0: return "-"
    v = ((fa-fb)/abs(fb))*100
    return f"{'+' if v>=0 else ''}{v:.2f}%"

def fmt_cur(val):
    v = to_float(val)
    return f"${v:.2f}" if v is not None else str(val or "-")

def _row(rows, key, val):
    for r in rows:
        if val.lower() in str(r.get(key,"") or "").strip().lower(): return r
    return {}

def _raw(rows, label, col):
    for row in rows:
        if not row: continue
        for i in [0,1]:
            if len(row)>i and str(row[i] or "").strip().lower()==label.lower():
                return row[col] if len(row)>col else None
    return None

def _v(val):
    if val is None: return "-"
    if isinstance(val, float): return str(int(val)) if val==int(val) else f"{val:.2f}"
    return str(val)


def calcular_variaciones(actual, anterior):
    v = {}
    ra = list(_row(actual["reg_comp"],   "Plataforma", "Total").values())
    rb = list(_row(anterior["reg_comp"], "Plataforma", "Total").values())
    v["fb_inv_act"]   = fmt_cur(ra[4] if len(ra)>4 else None)
    v["fb_inv_ant"]   = fmt_cur(rb[4] if len(rb)>4 else None)
    v["fb_inv_var"]   = calc_var(ra[4] if len(ra)>4 else None, rb[4] if len(rb)>4 else None)
    v["fb_leads_act"] = str(int(to_float(ra[2]))) if len(ra)>2 and to_float(ra[2]) else "-"
    v["fb_leads_ant"] = str(int(to_float(rb[2]))) if len(rb)>2 and to_float(rb[2]) else "-"
    v["fb_leads_var"] = calc_var(ra[2] if len(ra)>2 else None, rb[2] if len(rb)>2 else None)
    v["fb_cpl_act"]   = fmt_cur(ra[3] if len(ra)>3 else None)
    v["fb_cpl_ant"]   = fmt_cur(rb[3] if len(rb)>3 else None)
    v["fb_cpl_var"]   = calc_var(ra[3] if len(ra)>3 else None, rb[3] if len(rb)>3 else None)
    ga = list(_row(actual["reg_comp"],   "Plataforma", "Google").values())
    gb = list(_row(anterior["reg_comp"], "Plataforma", "Google").values())
    v["g_inv_act"]   = fmt_cur(ga[4] if len(ga)>4 else None)
    v["g_inv_ant"]   = fmt_cur(gb[4] if len(gb)>4 else None)
    v["g_inv_var"]   = calc_var(ga[4] if len(ga)>4 else None, gb[4] if len(gb)>4 else None)
    v["g_leads_act"] = str(int(to_float(ga[2]))) if len(ga)>2 and to_float(ga[2]) else "-"
    v["g_leads_ant"] = str(int(to_float(gb[2]))) if len(gb)>2 and to_float(gb[2]) else "-"
    v["g_leads_var"] = calc_var(ga[2] if len(ga)>2 else None, gb[2] if len(gb)>2 else None)
    v["g_cpl_act"]   = fmt_cur(ga[3] if len(ga)>3 else None)
    v["g_cpl_ant"]   = fmt_cur(gb[3] if len(gb)>3 else None)
    v["g_cpl_var"]   = calc_var(ga[3] if len(ga)>3 else None, gb[3] if len(gb)>3 else None)
    return v


def generate_texts_with_gemini(data_actual, variaciones, config):
    from google import genai
    client = genai.Client(api_key=GEMINI_API_KEY)
    reg = [r for r in data_actual["reg_comp"] if r.get("Plataforma") or r.get("Campaña")]
    inc = [r for r in data_actual["inc_comp"] if r.get("Plataforma") or r.get("Campaña")]
    prompt = f"""Eres Programmatic Analyst Senior. Analiza datos de Sider Express para reporte semanal.
PERIODO: {config.get('periodo_inicio')} - {config.get('periodo_fin')} | MES: {config.get('mes')}
CAMPAÑAS: {json.dumps(reg, ensure_ascii=False)}
METRICAS SEC: {json.dumps(data_actual['reg_met'], ensure_ascii=False)}
INCENTIVOS: {json.dumps(inc, ensure_ascii=False)}
VARIACIONES: FB Inv:{variaciones['fb_inv_ant']}→{variaciones['fb_inv_act']}({variaciones['fb_inv_var']}) FB Leads:{variaciones['fb_leads_ant']}→{variaciones['fb_leads_act']}({variaciones['fb_leads_var']}) FB CPL:{variaciones['fb_cpl_ant']}→{variaciones['fb_cpl_act']}({variaciones['fb_cpl_var']}) G Inv:{variaciones['g_inv_ant']}→{variaciones['g_inv_act']}({variaciones['g_inv_var']}) G Leads:{variaciones['g_leads_ant']}→{variaciones['g_leads_act']}({variaciones['g_leads_var']}) G CPL:{variaciones['g_cpl_ant']}→{variaciones['g_cpl_act']}({variaciones['g_cpl_var']})
Responde SOLO JSON sin backticks: {{"comentario_overview":"...","comentario_metricas_sec":"...","comentario_incentivos":"...","comentario_inc_metricas":"...","fb_comentario_cpl":"CPL (Costo por lead): ...","fb_comentario_leads":"Leads: ...","g_comentario_cpl":"CPL (Costo por lead): ...","g_comentario_leads":"Leads: ...","next_steps_meta":"b1\\nb2\\nb3","next_steps_google":"b1\\nb2\\nb3","next_steps_seguimiento":"b1\\nb2\\nb3"}}"""
    response = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
    del client; gc.collect()
    raw = re.sub(r"^```(?:json)?\s*","",response.text.strip())
    raw = re.sub(r"\s*```$","",raw)
    del response; gc.collect()
    return json.loads(raw.strip())


def build_replacements(data_actual, variaciones, texts, config):
    rep = {}
    act = data_actual
    rep["{{MES}}"]            = config.get("mes","")
    rep["{{PERIODO_INICIO}}"] = config.get("periodo_inicio","")
    rep["{{PERIODO_FIN}}"]    = config.get("periodo_fin","")
    rep["{{SEMANA}}"]         = config.get("semana","")
    for label, prefix in [("Meta Perfo","PRES_META_PERFO"),("Meta Branding","PRES_META_BRAND"),
                           ("Google","PRES_GOOGLE"),("Tik Tok","PRES_TIKTOK"),("Total","PRES_TOTAL")]:
        r = _row(act["pres"],"Plataforma",label)
        rep[f"{{{{{prefix}_REAL}}}}"] = _v(r.get("Real","-"))
        rep[f"{{{{{prefix}_META}}}}"] = _v(r.get("Meta","-"))
        rep[f"{{{{{prefix}_LOG}}}}"]  = _v(r.get("Logrado","-"))
    t = act["total_raw"]
    rep["{{LEADS_FB_REAL}}"]      = _v(_raw(t,"Facebook",2))
    rep["{{LEADS_FB_META}}"]      = _v(_raw(t,"Facebook",3))
    rep["{{LEADS_FB_LOG}}"]       = _v(_raw(t,"Facebook",4))
    rep["{{LEADS_GOOGLE_REAL}}"]  = _v(_raw(t,"Google ",2))
    rep["{{LEADS_TOTAL_REAL}}"]   = _v(_raw(t,"Total",2))
    rep["{{LEADS_CAL_PCT}}"]      = _v(_raw(t,"Leads calificados",2))
    rep["{{LEADS_CAL_REAL}}"]     = _v(_raw(t,"Leads calificados",3))
    rep["{{LEADS_CAL_META}}"]     = _v(_raw(t,"Leads calificados",4))
    rep["{{LEADS_CAL_LOG}}"]      = _v(_raw(t,"Leads calificados",5))
    rep["{{FACT_REAL}}"]          = _v(_raw(t,"Facturación",8))
    rep["{{FACT_META}}"]          = _v(_raw(t,"Facturación",9))
    rep["{{FACT_LOG}}"]           = _v(_raw(t,"Facturación",10))
    rep["{{VENTAS_REAL}}"]        = _v(_raw(t,"VENTAS",2))
    rep["{{VENTAS_META}}"]        = _v(_raw(t,"VENTAS",3))
    rep["{{VENTAS_LOG}}"]         = _v(_raw(t,"VENTAS",4))
    for camp,prefix in [("Lima","LIMA"),("Trujillo B2C","TRUJ_B2C"),("Trujillo B2B","TRUJ_B2B")]:
        r = list(_row(act["reg_comp"],"Campaña",camp).values())
        rep[f"{{{{{prefix}_LEADS_ACT}}}}"] = _v(r[2]) if len(r)>2 else "-"
        rep[f"{{{{{prefix}_CPL_ACT}}}}"]   = _v(r[3]) if len(r)>3 else "-"
        rep[f"{{{{{prefix}_INV_ACT}}}}"]   = _v(r[4]) if len(r)>4 else "-"
        rep[f"{{{{{prefix}_LEADS_ANT}}}}"] = _v(r[5]) if len(r)>5 else "-"
        rep[f"{{{{{prefix}_CPL_ANT}}}}"]   = _v(r[6]) if len(r)>6 else "-"
        rep[f"{{{{{prefix}_INV_ANT}}}}"]   = _v(r[7]) if len(r)>7 else "-"
    rg = list(_row(act["reg_comp"],"Plataforma","Google").values())
    rep["{{GOOGLE_LEADS_ACT}}"] = _v(rg[2]) if len(rg)>2 else "-"
    rep["{{GOOGLE_CPL_ACT}}"]   = _v(rg[3]) if len(rg)>3 else "-"
    rep["{{GOOGLE_INV_ACT}}"]   = _v(rg[4]) if len(rg)>4 else "-"
    rep["{{GOOGLE_LEADS_ANT}}"] = _v(rg[5]) if len(rg)>5 else "-"
    rep["{{GOOGLE_CPL_ANT}}"]   = _v(rg[6]) if len(rg)>6 else "-"
    rep["{{GOOGLE_INV_ANT}}"]   = _v(rg[7]) if len(rg)>7 else "-"
    rt = list(_row(act["reg_comp"],"Plataforma","Total").values())
    rep["{{TOTAL_LEADS_ACT}}"] = _v(rt[2]) if len(rt)>2 else "-"
    rep["{{TOTAL_CPL_ACT}}"]   = _v(rt[3]) if len(rt)>3 else "-"
    rep["{{TOTAL_INV_ACT}}"]   = _v(rt[4]) if len(rt)>4 else "-"
    rep["{{TOTAL_LEADS_ANT}}"] = _v(rt[5]) if len(rt)>5 else "-"
    rep["{{TOTAL_CPL_ANT}}"]   = _v(rt[6]) if len(rt)>6 else "-"
    rep["{{TOTAL_INV_ANT}}"]   = _v(rt[7]) if len(rt)>7 else "-"
    for camp,prefix in [("Lima","LIMA"),("Trujillo B2C","TRUJ_B2C"),("Trujillo B2B","TRUJ_B2B")]:
        r = _row(act["reg_met"],"Campaña",camp)
        rep[f"{{{{{prefix}_CTR}}}}"]      = _v(r.get("CTR ",r.get("CTR","-")))
        rep[f"{{{{{prefix}_CTR_UNI}}}}"]  = _v(r.get("CTR unico","-"))
        rep[f"{{{{{prefix}_PCT_MSJS}}}}"] = _v(r.get("% Mensajes","-"))
        rep[f"{{{{{prefix}_CPM}}}}"]      = _v(r.get("CPM","-"))
        rep[f"{{{{{prefix}_CLICS}}}}"]    = _v(r.get("Clic enlace","-"))
        rep[f"{{{{{prefix}_ALCANCE}}}}"]  = _v(r.get("Alcance","-"))
    rtt = _row(act["reg_met"],"Plataforma","Total") or _row(act["reg_met"],"Campaña","Total")
    rep["{{TOTAL_CTR}}"]         = _v(rtt.get("CTR ",rtt.get("CTR","-")))
    rep["{{TOTAL_CTR_UNI}}"]     = _v(rtt.get("CTR unico","-"))
    rep["{{TOTAL_PCT_MSJS}}"]    = _v(rtt.get("% Mensajes","-"))
    rep["{{TOTAL_CPM}}"]         = _v(rtt.get("CPM","-"))
    rep["{{TOTAL_IMPRESIONES}}"] = _v(rtt.get("Impresiones","-"))
    rep["{{TOTAL_ALCANCE}}"]     = _v(rtt.get("Alcance","-"))
    for camp,prefix in [("Incentivos trujillo","INC_TRUJ"),("Incentivos Lima","INC_LIMA")]:
        r = list(_row(act["inc_comp"],"Campaña",camp).values())
        rep[f"{{{{{prefix}_LEADS_ACT}}}}"] = _v(r[2]) if len(r)>2 else "-"
        rep[f"{{{{{prefix}_CPL_ACT}}}}"]   = _v(r[3]) if len(r)>3 else "-"
        rep[f"{{{{{prefix}_INV_ACT}}}}"]   = _v(r[4]) if len(r)>4 else "-"
        rep[f"{{{{{prefix}_LEADS_ANT}}}}"] = _v(r[5]) if len(r)>5 else "-"
        rep[f"{{{{{prefix}_CPL_ANT}}}}"]   = _v(r[6]) if len(r)>6 else "-"
        rep[f"{{{{{prefix}_INV_ANT}}}}"]   = _v(r[7]) if len(r)>7 else "-"
    rit = list(_row(act["inc_comp"],"Plataforma","Total").values())
    rep["{{INC_TOTAL_LEADS_ACT}}"] = _v(rit[2]) if len(rit)>2 else "-"
    rep["{{INC_TOTAL_CPL_ACT}}"]   = _v(rit[3]) if len(rit)>3 else "-"
    rep["{{INC_TOTAL_INV_ACT}}"]   = _v(rit[4]) if len(rit)>4 else "-"
    rep["{{INC_TOTAL_LEADS_ANT}}"] = _v(rit[5]) if len(rit)>5 else "-"
    rep["{{INC_TOTAL_CPL_ANT}}"]   = _v(rit[6]) if len(rit)>6 else "-"
    rep["{{INC_TOTAL_INV_ANT}}"]   = _v(rit[7]) if len(rit)>7 else "-"
    for camp,prefix in [("Incentivos trujillo","INC_TRUJ"),("Incentivos Lima","INC_LIMA")]:
        r = _row(act["inc_met"],"Campaña",camp)
        rep[f"{{{{{prefix}_CTR}}}}"]      = _v(r.get("CTR ",r.get("CTR","-")))
        rep[f"{{{{{prefix}_CTR_UNI}}}}"]  = _v(r.get("CTR unico","-"))
        rep[f"{{{{{prefix}_PCT_MSJS}}}}"] = _v(r.get("% Mensajes","-"))
        rep[f"{{{{{prefix}_CPM}}}}"]      = _v(r.get("CPM","-"))
    rimt = _row(act["inc_met"],"Plataforma","Total") or _row(act["inc_met"],"Campaña","Total")
    rep["{{INC_TOTAL_CTR}}"]      = _v(rimt.get("CTR ",rimt.get("CTR","-")))
    rep["{{INC_TOTAL_CTR_UNI}}"]  = _v(rimt.get("CTR unico","-"))
    rep["{{INC_TOTAL_PCT_MSJS}}"] = _v(rimt.get("% Mensajes","-"))
    rep["{{INC_TOTAL_CPM}}"]      = _v(rimt.get("CPM","-"))
    for k,v in variaciones.items():
        rep[f"{{{{{k.upper()}}}}}"] = v
    rep["{{COMENTARIO_OVERVIEW}}"]     = texts.get("comentario_overview","")
    rep["{{COMENTARIO_METRICAS_SEC}}"] = texts.get("comentario_metricas_sec","")
    rep["{{COMENTARIO_INCENTIVOS}}"]   = texts.get("comentario_incentivos","")
    rep["{{COMENTARIO_INC_METRICAS}}"] = texts.get("comentario_inc_metricas","")
    rep["{{FB_COMENTARIO_CPL}}"]       = texts.get("fb_comentario_cpl","")
    rep["{{FB_COMENTARIO_LEADS}}"]     = texts.get("fb_comentario_leads","")
    rep["{{G_COMENTARIO_CPL}}"]        = texts.get("g_comentario_cpl","")
    rep["{{G_COMENTARIO_LEADS}}"]      = texts.get("g_comentario_leads","")
    rep["{{NEXT_STEPS_META}}"]         = texts.get("next_steps_meta","")
    rep["{{NEXT_STEPS_GOOGLE}}"]       = texts.get("next_steps_google","")
    rep["{{NEXT_STEPS_SEGUIMIENTO}}"]  = texts.get("next_steps_seguimiento","")
    return rep


def fill_pptx(pptx_bytes, replacements):
    from pptx import Presentation
    prs = Presentation(io.BytesIO(pptx_bytes))
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame: continue
            for para in shape.text_frame.paragraphs:
                full_text = "".join(run.text for run in para.runs)
                if not any(ph in full_text for ph in replacements): continue
                new_text = full_text
                for ph, val in replacements.items():
                    new_text = new_text.replace(ph, str(val))
                if para.runs:
                    para.runs[0].text = new_text
                    for run in para.runs[1:]: run.text = ""
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    result = output.read()
    del prs, output; gc.collect()
    return result


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT",5000)), debug=False)
