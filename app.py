"""
Sider Express — Report Web App
================================
Servidor Flask que recibe el XLSX + PPTX, genera textos con Claude,
rellena los placeholders y devuelve el PPTX actualizado.
"""

import os
import io
import json
import re
import tempfile
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename
import anthropic
import openpyxl
from pptx import Presentation
from pptx.util import Pt
import zipfile

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB máx

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")

# ─────────────────────────────────────────────
# RUTA PRINCIPAL
# ─────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    try:
        # 1. Recibir archivos
        if "xlsx" not in request.files or "pptx" not in request.files:
            return jsonify({"error": "Faltan archivos. Se necesitan el XLSX y el PPTX."}), 400

        xlsx_file = request.files["xlsx"]
        pptx_file = request.files["pptx"]

        # 2. Leer configuración del formulario
        config = {
            "mes":             request.form.get("mes", ""),
            "periodo_inicio":  request.form.get("periodo_inicio", ""),
            "periodo_fin":     request.form.get("periodo_fin", ""),
            "semana":          request.form.get("semana", ""),
            "fb_inv_var":      request.form.get("fb_inv_var", "-"),
            "fb_leads_var":    request.form.get("fb_leads_var", "-"),
            "fb_cpl_act":      request.form.get("fb_cpl_act", "-"),
            "fb_cpl_ant":      request.form.get("fb_cpl_ant", "-"),
            "fb_cpl_var":      request.form.get("fb_cpl_var", "-"),
            "g_inv_var":       request.form.get("g_inv_var", "-"),
            "g_leads_var":     request.form.get("g_leads_var", "-"),
            "g_cpl_var":       request.form.get("g_cpl_var", "-"),
        }

        # 3. Leer datos del XLSX en memoria
        xlsx_bytes = xlsx_file.read()
        data = extract_xlsx_data(xlsx_bytes, config)

        # 4. Generar textos con Claude
        texts = generate_texts_with_claude(data, config)

        # 5. Construir diccionario de reemplazos
        replacements = build_replacements(data, texts, config)

        # 6. Rellenar el PPTX con los reemplazos
        pptx_bytes = pptx_file.read()
        output_bytes = fill_pptx(pptx_bytes, replacements)

        # 7. Devolver el archivo
        mes = config.get("mes", "reporte").upper()
        pi  = config.get("periodo_inicio", "").replace("/", "")
        pf  = config.get("periodo_fin", "").replace("/", "")
        filename = f"Sider_Reporte_{mes}_{pi}_{pf}.pptx"

        return send_file(
            io.BytesIO(output_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# ─────────────────────────────────────────────
# EXTRACCIÓN DE DATOS DEL XLSX
# ─────────────────────────────────────────────
def extract_xlsx_data(xlsx_bytes, config):
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    sheets = wb.sheetnames

    def read_sheet(name, max_row=30):
        if name not in sheets:
            return []
        ws = wb[name]
        rows = []
        for row in ws.iter_rows(min_row=1, max_row=max_row, values_only=True):
            if any(v is not None for v in row):
                rows.append(list(row))
        return rows

    def rows_to_dicts(rows, header_row=0):
        if not rows or len(rows) <= header_row:
            return []
        headers = [str(h) if h is not None else f"col{i}"
                   for i, h in enumerate(rows[header_row])]
        result = []
        for row in rows[header_row + 1:]:
            if any(v is not None for v in row):
                padded = list(row) + [None] * (len(headers) - len(row))
                result.append(dict(zip(headers, padded)))
        return result

    # Leer pestañas clave
    pres_rows    = read_sheet("PRESUPUESTO", 15)
    total_rows   = read_sheet("TOTAL", 30)
    reg_comp_rows = read_sheet("CAMPAÑAS REGULARES COMPARATIVA ", 20)
    reg_met_rows  = read_sheet("CAMPAÑAS REGULARES METRICAS SEC", 20)
    inc_comp_rows = read_sheet("CAMPAÑAS INCENTIVOS COMPARATIVA", 15)
    inc_met_rows  = read_sheet("CAMPAÑAS INCENTIVOS METRICAS SE", 15)
    funnel_rows   = read_sheet("FUNNEL (REPORTE MENSUAL) ", 15)

    # Encontrar fila de cabecera real (la que tiene "Plataforma")
    def find_header(rows):
        for i, row in enumerate(rows):
            if row and any("plataforma" in str(v).lower() for v in row if v):
                return i
        return 0

    pres_dicts    = rows_to_dicts(pres_rows, find_header(pres_rows))
    reg_comp      = rows_to_dicts(reg_comp_rows, find_header(reg_comp_rows))
    reg_met       = rows_to_dicts(reg_met_rows, find_header(reg_met_rows))
    inc_comp      = rows_to_dicts(inc_comp_rows, find_header(inc_comp_rows))
    inc_met       = rows_to_dicts(inc_met_rows, find_header(inc_met_rows))

    return {
        "config":    config,
        "pres":      pres_dicts,
        "total_raw": total_rows,
        "reg_comp":  reg_comp,
        "reg_met":   reg_met,
        "inc_comp":  inc_comp,
        "inc_met":   inc_met,
        "funnel":    funnel_rows,
    }


# ─────────────────────────────────────────────
# GENERACIÓN DE TEXTOS CON CLAUDE
# ─────────────────────────────────────────────
def generate_texts_with_claude(data, config):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    reg_comp_clean = [r for r in data["reg_comp"]
                      if r.get("Plataforma") or r.get("Campaña")]
    inc_comp_clean = [r for r in data["inc_comp"]
                      if r.get("Plataforma") or r.get("Campaña")]

    prompt = f"""
Eres un Programmatic Analyst Senior. Analiza los datos de performance de Sider Express
y genera textos analíticos ejecutivos para el reporte semanal.

PERIODO: {config.get('periodo_inicio')} - {config.get('periodo_fin')}
MES: {config.get('mes')}

=== CAMPAÑAS REGULARES — COMPARATIVA ===
{json.dumps(reg_comp_clean, ensure_ascii=False, indent=2)}

=== CAMPAÑAS REGULARES — MÉTRICAS SECUNDARIAS ===
{json.dumps(data['reg_met'], ensure_ascii=False, indent=2)}

=== CAMPAÑAS INCENTIVOS — COMPARATIVA ===
{json.dumps(inc_comp_clean, ensure_ascii=False, indent=2)}

=== CAMPAÑAS INCENTIVOS — MÉTRICAS SECUNDARIAS ===
{json.dumps(data['inc_met'], ensure_ascii=False, indent=2)}

Genera textos analíticos directos, basados en datos. Máximo 4-5 oraciones por campo.
Sin viñetas. Destaca variaciones %, plataformas líderes, eficiencia de CPL y CTR.

Responde SOLO con JSON válido, sin backticks ni texto extra:

{{
  "comentario_overview": "...",
  "comentario_metricas_sec": "...",
  "comentario_incentivos": "...",
  "comentario_inc_metricas": "...",
  "fb_comentario_cpl": "CPL (Costo por lead): ...",
  "fb_comentario_leads": "Leads: ...",
  "g_comentario_cpl": "CPL (Costo por lead): ...",
  "g_comentario_leads": "Leads: ...",
  "next_steps_meta": "bullet 1\\nbullet 2\\nbullet 3\\nbullet 4",
  "next_steps_google": "bullet 1\\nbullet 2\\nbullet 3",
  "next_steps_seguimiento": "bullet 1\\nbullet 2\\nbullet 3"
}}
"""

    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = message.content[0].text.strip()
    # Limpiar backticks
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return json.loads(raw.strip())


# ─────────────────────────────────────────────
# CONSTRUCCIÓN DE REEMPLAZOS
# ─────────────────────────────────────────────
def _v(val):
    """Convierte valor a string seguro."""
    if val is None:
        return "-"
    if isinstance(val, float):
        if val == int(val):
            return str(int(val))
        return f"{val:.2f}"
    return str(val)


def _find_row(rows, key_col, key_val):
    """Busca una fila donde key_col contiene key_val (case-insensitive)."""
    for row in rows:
        cell = str(row.get(key_col, "") or "").strip().lower()
        if key_val.lower() in cell:
            return row
    return {}


def _get_raw(rows, label, col_idx):
    """Busca en una lista plana de listas por label en col 0 o 1."""
    for row in rows:
        if not row:
            continue
        for i in [0, 1]:
            if len(row) > i and str(row[i] or "").strip().lower() == label.lower():
                return _v(row[col_idx]) if len(row) > col_idx else "-"
    return "-"


def build_replacements(data, texts, config):
    rep = {}

    # ── Portada
    rep["{{MES}}"]            = config.get("mes", "")
    rep["{{PERIODO_INICIO}}"] = config.get("periodo_inicio", "")
    rep["{{PERIODO_FIN}}"]    = config.get("periodo_fin", "")
    rep["{{SEMANA}}"]         = config.get("semana", "")

    # ── PRESUPUESTO
    pres_items = [
        ("Meta Perfo",    "PRES_META_PERFO"),
        ("Meta Branding", "PRES_META_BRAND"),
        ("Google",        "PRES_GOOGLE"),
        ("Tik Tok",       "PRES_TIKTOK"),
        ("Total",         "PRES_TOTAL"),
    ]
    col_names = ["Real", "Meta", "Logrado"]
    col_keys  = ["REAL", "META", "LOG"]
    for label, prefix in pres_items:
        row = _find_row(data["pres"], "Plataforma", label)
        for col, key in zip(col_names, col_keys):
            rep[f"{{{{{prefix}_{key}}}}}"] = _v(row.get(col, "-"))

    # ── TOTAL (leads, ventas, facturación)
    total = data["total_raw"]
    rep["{{LEADS_FB_REAL}}"]     = _get_raw(total, "Facebook", 2)
    rep["{{LEADS_FB_META}}"]     = _get_raw(total, "Facebook", 3)
    rep["{{LEADS_FB_LOG}}"]      = _get_raw(total, "Facebook", 4)
    rep["{{LEADS_GOOGLE_REAL}}"] = _get_raw(total, "Google ", 2)
    rep["{{LEADS_TOTAL_REAL}}"]  = _get_raw(total, "Total", 2)
    rep["{{LEADS_CAL_PCT}}"]     = _get_raw(total, "Leads calificados", 2)
    rep["{{LEADS_CAL_REAL}}"]    = _get_raw(total, "Leads calificados", 3)
    rep["{{LEADS_CAL_META}}"]    = _get_raw(total, "Leads calificados", 4)
    rep["{{LEADS_CAL_LOG}}"]     = _get_raw(total, "Leads calificados", 5)
    rep["{{FACT_REAL}}"]         = _get_raw(total, "Facturación", 8)
    rep["{{FACT_META}}"]         = _get_raw(total, "Facturación", 9)
    rep["{{FACT_LOG}}"]          = _get_raw(total, "Facturación", 10)
    rep["{{VENTAS_REAL}}"]       = _get_raw(total, "VENTAS", 2)
    rep["{{VENTAS_META}}"]       = _get_raw(total, "VENTAS", 3)
    rep["{{VENTAS_LOG}}"]        = _get_raw(total, "VENTAS", 4)

    # ── CAMPAÑAS REGULARES COMPARATIVA
    camp_reg = [
        ("Lima",         "LIMA"),
        ("Trujillo B2C", "TRUJ_B2C"),
        ("Trujillo B2B", "TRUJ_B2B"),
    ]
    for camp, prefix in camp_reg:
        row = _find_row(data["reg_comp"], "Campaña", camp)
        vals = list(row.values())
        # columnas: Plataforma(0), Campaña(1), LeadsAct(2), CPLact(3), InvAct(4),
        #           LeadsAnt(5), CPLant(6), InvAnt(7)
        rep[f"{{{{{prefix}_LEADS_ACT}}}}"] = _v(vals[2]) if len(vals) > 2 else "-"
        rep[f"{{{{{prefix}_CPL_ACT}}}}"]   = _v(vals[3]) if len(vals) > 3 else "-"
        rep[f"{{{{{prefix}_INV_ACT}}}}"]   = _v(vals[4]) if len(vals) > 4 else "-"
        rep[f"{{{{{prefix}_LEADS_ANT}}}}"] = _v(vals[5]) if len(vals) > 5 else "-"
        rep[f"{{{{{prefix}_CPL_ANT}}}}"]   = _v(vals[6]) if len(vals) > 6 else "-"
        rep[f"{{{{{prefix}_INV_ANT}}}}"]   = _v(vals[7]) if len(vals) > 7 else "-"

    row_g = _find_row(data["reg_comp"], "Plataforma", "Google")
    vals_g = list(row_g.values())
    rep["{{GOOGLE_LEADS_ACT}}"] = _v(vals_g[2]) if len(vals_g) > 2 else "-"
    rep["{{GOOGLE_CPL_ACT}}"]   = _v(vals_g[3]) if len(vals_g) > 3 else "-"
    rep["{{GOOGLE_INV_ACT}}"]   = _v(vals_g[4]) if len(vals_g) > 4 else "-"
    rep["{{GOOGLE_LEADS_ANT}}"] = _v(vals_g[5]) if len(vals_g) > 5 else "-"
    rep["{{GOOGLE_CPL_ANT}}"]   = _v(vals_g[6]) if len(vals_g) > 6 else "-"
    rep["{{GOOGLE_INV_ANT}}"]   = _v(vals_g[7]) if len(vals_g) > 7 else "-"

    row_t = _find_row(data["reg_comp"], "Plataforma", "Total")
    vals_t = list(row_t.values())
    rep["{{TOTAL_LEADS_ACT}}"] = _v(vals_t[2]) if len(vals_t) > 2 else "-"
    rep["{{TOTAL_CPL_ACT}}"]   = _v(vals_t[3]) if len(vals_t) > 3 else "-"
    rep["{{TOTAL_INV_ACT}}"]   = _v(vals_t[4]) if len(vals_t) > 4 else "-"
    rep["{{TOTAL_LEADS_ANT}}"] = _v(vals_t[5]) if len(vals_t) > 5 else "-"
    rep["{{TOTAL_CPL_ANT}}"]   = _v(vals_t[6]) if len(vals_t) > 6 else "-"
    rep["{{TOTAL_INV_ANT}}"]   = _v(vals_t[7]) if len(vals_t) > 7 else "-"

    # ── MÉTRICAS SECUNDARIAS REGULARES
    met_reg = [("Lima","LIMA"),("Trujillo B2C","TRUJ_B2C"),("Trujillo B2B","TRUJ_B2B")]
    for camp, prefix in met_reg:
        row = _find_row(data["reg_met"], "Campaña", camp)
        rep[f"{{{{{prefix}_CTR}}}}"]      = _v(row.get("CTR ", row.get("CTR", "-")))
        rep[f"{{{{{prefix}_CTR_UNI}}}}"]  = _v(row.get("CTR unico", "-"))
        rep[f"{{{{{prefix}_PCT_MSJS}}}}"] = _v(row.get("% Mensajes", "-"))
        rep[f"{{{{{prefix}_CPM}}}}"]      = _v(row.get("CPM", "-"))
        rep[f"{{{{{prefix}_CLICS}}}}"]    = _v(row.get("Clic enlace", "-"))
        rep[f"{{{{{prefix}_ALCANCE}}}}"]  = _v(row.get("Alcance", "-"))

    row_tt = _find_row(data["reg_met"], "Campaña", "Total") or \
             _find_row(data["reg_met"], "Plataforma", "Total")
    rep["{{TOTAL_CTR}}"]       = _v(row_tt.get("CTR ", row_tt.get("CTR", "-")))
    rep["{{TOTAL_CTR_UNI}}"]   = _v(row_tt.get("CTR unico", "-"))
    rep["{{TOTAL_PCT_MSJS}}"]  = _v(row_tt.get("% Mensajes", "-"))
    rep["{{TOTAL_CPM}}"]       = _v(row_tt.get("CPM", "-"))
    rep["{{TOTAL_IMPRESIONES}}"]= _v(row_tt.get("Impresiones", "-"))
    rep["{{TOTAL_ALCANCE}}"]   = _v(row_tt.get("Alcance", "-"))

    # ── INCENTIVOS COMPARATIVA
    camp_inc = [
        ("Incentivos trujillo", "INC_TRUJ"),
        ("Incentivos Lima",     "INC_LIMA"),
    ]
    for camp, prefix in camp_inc:
        row = _find_row(data["inc_comp"], "Campaña", camp)
        vals = list(row.values())
        rep[f"{{{{{prefix}_LEADS_ACT}}}}"] = _v(vals[2]) if len(vals) > 2 else "-"
        rep[f"{{{{{prefix}_CPL_ACT}}}}"]   = _v(vals[3]) if len(vals) > 3 else "-"
        rep[f"{{{{{prefix}_INV_ACT}}}}"]   = _v(vals[4]) if len(vals) > 4 else "-"
        rep[f"{{{{{prefix}_LEADS_ANT}}}}"] = _v(vals[5]) if len(vals) > 5 else "-"
        rep[f"{{{{{prefix}_CPL_ANT}}}}"]   = _v(vals[6]) if len(vals) > 6 else "-"
        rep[f"{{{{{prefix}_INV_ANT}}}}"]   = _v(vals[7]) if len(vals) > 7 else "-"

    row_it = _find_row(data["inc_comp"], "Plataforma", "Total")
    vals_it = list(row_it.values())
    rep["{{INC_TOTAL_LEADS_ACT}}"] = _v(vals_it[2]) if len(vals_it) > 2 else "-"
    rep["{{INC_TOTAL_CPL_ACT}}"]   = _v(vals_it[3]) if len(vals_it) > 3 else "-"
    rep["{{INC_TOTAL_INV_ACT}}"]   = _v(vals_it[4]) if len(vals_it) > 4 else "-"
    rep["{{INC_TOTAL_LEADS_ANT}}"] = _v(vals_it[5]) if len(vals_it) > 5 else "-"
    rep["{{INC_TOTAL_CPL_ANT}}"]   = _v(vals_it[6]) if len(vals_it) > 6 else "-"
    rep["{{INC_TOTAL_INV_ANT}}"]   = _v(vals_it[7]) if len(vals_it) > 7 else "-"

    # ── INCENTIVOS MÉTRICAS SEC
    for camp, prefix in [("Incentivos trujillo","INC_TRUJ"),("Incentivos Lima","INC_LIMA")]:
        row = _find_row(data["inc_met"], "Campaña", camp)
        rep[f"{{{{{prefix}_CTR}}}}"]      = _v(row.get("CTR ", row.get("CTR", "-")))
        rep[f"{{{{{prefix}_CTR_UNI}}}}"]  = _v(row.get("CTR unico", "-"))
        rep[f"{{{{{prefix}_PCT_MSJS}}}}"] = _v(row.get("% Mensajes", "-"))
        rep[f"{{{{{prefix}_CPM}}}}"]      = _v(row.get("CPM", "-"))

    row_imt = _find_row(data["inc_met"], "Campaña", "Total") or \
              _find_row(data["inc_met"], "Plataforma", "Total")
    rep["{{INC_TOTAL_CTR}}"]      = _v(row_imt.get("CTR ", row_imt.get("CTR", "-")))
    rep["{{INC_TOTAL_CTR_UNI}}"]  = _v(row_imt.get("CTR unico", "-"))
    rep["{{INC_TOTAL_PCT_MSJS}}"] = _v(row_imt.get("% Mensajes", "-"))
    rep["{{INC_TOTAL_CPM}}"]      = _v(row_imt.get("CPM", "-"))

    # ── Funnel variaciones (vienen del formulario)
    rep["{{FB_INV_ACT}}"]   = config.get("total_inv_act",  rep.get("{{TOTAL_INV_ACT}}", "-"))
    rep["{{FB_LEADS_ACT}}"] = rep.get("{{TOTAL_LEADS_ACT}}", "-")
    rep["{{FB_INV_VAR}}"]   = config.get("fb_inv_var", "-")
    rep["{{FB_LEADS_VAR}}"] = config.get("fb_leads_var", "-")
    rep["{{FB_CPL_ACT}}"]   = config.get("fb_cpl_act", "-")
    rep["{{FB_CPL_ANT}}"]   = config.get("fb_cpl_ant", "-")
    rep["{{FB_CPL_VAR}}"]   = config.get("fb_cpl_var", "-")
    rep["{{G_LEADS_ACT}}"]  = rep.get("{{GOOGLE_LEADS_ACT}}", "-")
    rep["{{G_INV_ACT}}"]    = rep.get("{{GOOGLE_INV_ACT}}", "-")
    rep["{{G_CPL_ACT}}"]    = rep.get("{{GOOGLE_CPL_ACT}}", "-")
    rep["{{G_LEADS_ANT}}"]  = rep.get("{{GOOGLE_LEADS_ANT}}", "-")
    rep["{{G_INV_ANT}}"]    = rep.get("{{GOOGLE_INV_ANT}}", "-")
    rep["{{G_CPL_ANT}}"]    = rep.get("{{GOOGLE_CPL_ANT}}", "-")
    rep["{{G_INV_VAR}}"]    = config.get("g_inv_var", "-")
    rep["{{G_LEADS_VAR}}"]  = config.get("g_leads_var", "-")
    rep["{{G_CPL_VAR}}"]    = config.get("g_cpl_var", "-")

    # ── Textos de Claude
    rep["{{COMENTARIO_OVERVIEW}}"]     = texts.get("comentario_overview", "")
    rep["{{COMENTARIO_METRICAS_SEC}}"] = texts.get("comentario_metricas_sec", "")
    rep["{{COMENTARIO_INCENTIVOS}}"]   = texts.get("comentario_incentivos", "")
    rep["{{COMENTARIO_INC_METRICAS}}"] = texts.get("comentario_inc_metricas", "")
    rep["{{FB_COMENTARIO_CPL}}"]       = texts.get("fb_comentario_cpl", "")
    rep["{{FB_COMENTARIO_LEADS}}"]     = texts.get("fb_comentario_leads", "")
    rep["{{G_COMENTARIO_CPL}}"]        = texts.get("g_comentario_cpl", "")
    rep["{{G_COMENTARIO_LEADS}}"]      = texts.get("g_comentario_leads", "")
    rep["{{NEXT_STEPS_META}}"]         = texts.get("next_steps_meta", "")
    rep["{{NEXT_STEPS_GOOGLE}}"]       = texts.get("next_steps_google", "")
    rep["{{NEXT_STEPS_SEGUIMIENTO}}"]  = texts.get("next_steps_seguimiento", "")

    return rep


# ─────────────────────────────────────────────
# RELLENO DEL PPTX
# ─────────────────────────────────────────────
def fill_pptx(pptx_bytes, replacements):
    """
    Reemplaza todos los {{PLACEHOLDERS}} en el PPTX.
    Maneja el caso donde python-pptx fragmenta el texto en múltiples runs.
    """
    prs = Presentation(io.BytesIO(pptx_bytes))

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                # Primero reconstruimos el texto completo del párrafo
                full_text = "".join(run.text for run in para.runs)

                # Verificar si hay algún placeholder en este párrafo
                has_placeholder = any(ph in full_text for ph in replacements)
                if not has_placeholder:
                    continue

                # Aplicar todos los reemplazos al texto completo
                new_text = full_text
                for placeholder, value in replacements.items():
                    new_text = new_text.replace(placeholder, str(value))

                # Poner el texto completo en el primer run, vaciar el resto
                if para.runs:
                    para.runs[0].text = new_text
                    for run in para.runs[1:]:
                        run.text = ""

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output.read()


# ─────────────────────────────────────────────
# HEALTH CHECK
# ─────────────────────────────────────────────
@app.route("/health")
def health():
    return jsonify({"status": "ok", "time": datetime.utcnow().isoformat()})


if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
