import os, json, uuid, shutil
from datetime import datetime
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, abort
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

BASE_DIR   = Path(__file__).parent
DATA_DIR   = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
UPLOAD_DIR = BASE_DIR / "static" / "uploads"
for d in [DATA_DIR, OUTPUT_DIR, UPLOAD_DIR]: d.mkdir(parents=True, exist_ok=True)

ALLOWED = {"png","jpg","jpeg","gif","bmp","webp","heic","tiff"}
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
app.secret_key = "memoria-descriptiva-2024"

def allowed(f): return "." in f and f.rsplit(".",1)[1].lower() in ALLOWED
def load_form(fid):
    p = DATA_DIR / f"{fid}.json"
    if not p.exists(): return None
    return json.load(open(p, encoding="utf-8"))
def save_form(fid, data):
    data["updated_at"] = datetime.now().isoformat()
    json.dump(data, open(DATA_DIR/f"{fid}.json","w",encoding="utf-8"), ensure_ascii=False, indent=2)
def list_forms():
    forms=[]
    for p in sorted(DATA_DIR.glob("*.json"),key=os.path.getmtime,reverse=True):
        try:
            d=json.load(open(p,encoding="utf-8"))
            forms.append({"id":p.stem,"nombre":d.get("datos_generales",{}).get("nombre_tienda","Sin nombre"),
                "tipo_obra":d.get("datos_generales",{}).get("tipo_obra",""),
                "ciudad":d.get("datos_generales",{}).get("ciudad",""),
                "estado":d.get("estado","borrador"),"updated_at":d.get("updated_at","")})
        except: pass
    return forms

@app.route("/")
def index(): return render_template("index.html")
@app.route("/nuevo")
def nuevo():
    fid=str(uuid.uuid4())
    save_form(fid,{"id":fid,"estado":"borrador","created_at":datetime.now().isoformat()})
    return render_template("formulario.html", form_id=fid)
@app.route("/editar/<fid>")
def editar(fid):
    d=load_form(fid)
    if d is None: abort(404)
    return render_template("formulario.html", form_id=fid)
@app.route("/admin")
def admin(): return render_template("admin.html", forms=list_forms())

@app.route("/api/form/<fid>", methods=["GET"])
def api_get(fid):
    d=load_form(fid)
    if d is None: return jsonify({"error":"No encontrado"}),404
    return jsonify(d)
@app.route("/api/form/<fid>", methods=["POST"])
def api_save(fid):
    d=load_form(fid) or {"id":fid,"created_at":datetime.now().isoformat()}
    payload=request.get_json(force=True)
    d.update(payload)
    save_form(fid,d)
    return jsonify({"ok":True})
@app.route("/api/form/<fid>", methods=["DELETE"])
def api_delete(fid):
    p=DATA_DIR/f"{fid}.json"
    if p.exists(): p.unlink()
    up=UPLOAD_DIR/fid
    if up.exists(): shutil.rmtree(up)
    return jsonify({"ok":True})
@app.route("/api/upload/<fid>", methods=["POST"])
def api_upload(fid):
    if "file" not in request.files: return jsonify({"error":"No file"}),400
    file=request.files["file"]
    if not allowed(file.filename): return jsonify({"error":"Tipo no permitido"}),400
    dest=UPLOAD_DIR/fid; dest.mkdir(parents=True,exist_ok=True)
    ext=file.filename.rsplit(".",1)[1].lower()
    fname=f"{uuid.uuid4()}.{ext}"
    file.save(dest/fname)
    return jsonify({"url":f"/static/uploads/{fid}/{fname}","filename":fname})
@app.route("/api/upload/<fid>/<fname>", methods=["DELETE"])
def api_del_photo(fid,fname):
    fp=UPLOAD_DIR/fid/secure_filename(fname)
    if fp.exists(): fp.unlink()
    return jsonify({"ok":True})
@app.route("/api/generar/<fid>", methods=["POST"])
def api_generar(fid):
    d=load_form(fid)
    if d is None: return jsonify({"error":"No encontrado"}),404
    try:
        op=generar_word(d,fid)
        d["estado"]="completado"; d["word_path"]=str(op)
        save_form(fid,d)
        return jsonify({"ok":True,"download":f"/api/descargar/{fid}"})
    except Exception as e:
        import traceback
        return jsonify({"error":str(e),"trace":traceback.format_exc()}),500
@app.route("/api/descargar/<fid>")
def api_descargar(fid):
    d=load_form(fid)
    if d is None: abort(404)
    p=Path(d.get("word_path",""))
    if not p.exists(): abort(404)
    nombre=d.get("datos_generales",{}).get("nombre_tienda","Memoria") or "Memoria"
    return send_file(p,as_attachment=True,download_name=secure_filename(f"MEMORIA-{nombre}.docx"))

# ---- Word generator helpers ----
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def resolve_img_path(url):
    if not url: return None
    clean = url.lstrip("/").replace("/", os.sep)
    p1 = BASE_DIR / clean
    if p1.exists(): return p1
    p2 = BASE_DIR / "static" / clean
    if p2.exists(): return p2
    return None

def add_banner(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    shd = parse_xml(r'<w:shd {} w:fill="595959"/>'.format(nsdecls('w')))
    p._p.get_or_add_pPr().append(shd)
    r = p.add_run(f"  {text.upper()}  ")
    r.font.name = 'Arial'
    r.font.size = Pt(9.5)
    r.font.bold = True
    r.font.color.rgb = RGBColor(255, 255, 255)
    return p

def add_subtitle(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(text)
    r.font.name = 'Arial'
    r.font.size = Pt(9.0)
    r.font.bold = True
    return p

def add_body_p(doc, text, bold=False, space_after=2, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    if not text or not str(text).strip(): return None
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(str(text).strip())
    r.font.name = 'Arial'
    r.font.size = Pt(9.0)
    r.bold = bold
    return p

def add_field(doc, label, value):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.0
    r1 = p.add_run(f"{label}: ")
    r1.font.name = 'Arial'
    r1.font.size = Pt(9.0)
    r1.bold = True
    r2 = p.add_run(str(value) if value else "")
    r2.font.name = 'Arial'
    r2.font.size = Pt(9.0)
    return p

def add_photo_grid(doc, photos, fid, photo_counter_start=1):
    if not photos: return photo_counter_start
    photo_num = photo_counter_start
    for i in range(0, len(photos), 2):
        pair = photos[i:i+2]
        table = doc.add_table(rows=2, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            r'<w:tblBorders {} >'
            r'  <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'  <w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'  <w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
            r'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
            r'</w:tblBorders>'.format(nsdecls('w'))
        )
        tblPr.append(borders)
        
        col_width = Cm(8.85)
        for row in table.rows:
            row.cells[0].width = col_width
            row.cells[1].width = col_width
        
        for ci in range(2):
            ic = table.cell(0, ci)
            dc = table.cell(1, ci)
            set_cell_margins(ic, 80, 80, 100, 100)
            set_cell_margins(dc, 80, 80, 100, 100)
            
            if ci < len(pair):
                photo = pair[ci]
                url = photo.get("url", "")
                img_path = resolve_img_path(url)
                if img_path and img_path.exists():
                    try:
                        p = ic.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.paragraph_format.space_before = Pt(2)
                        p.paragraph_format.space_after = Pt(2)
                        p.add_run().add_picture(str(img_path), width=Inches(3.2))
                    except Exception as e:
                        print("Error inserting picture:", e)
                        ic.text = "[Error al insertar imagen]"
                else:
                    ic.text = "[Imagen no disponible]"
                
                desc = photo.get("descripcion", "").strip()
                p_desc = dc.paragraphs[0]
                p_desc.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p_desc.paragraph_format.space_before = Pt(2)
                p_desc.paragraph_format.space_after = Pt(2)
                
                if not desc.upper().startswith("FOTO"):
                    r_lbl = p_desc.add_run(f"FOTO {photo_num}. ")
                    r_lbl.font.name = "Arial"
                    r_lbl.font.size = Pt(9.0)
                    r_lbl.font.bold = True
                
                if desc:
                    r_txt = p_desc.add_run(desc)
                    r_txt.font.name = "Arial"
                    r_txt.font.size = Pt(9.0)
                photo_num += 1
            else:
                ic.text = ""
                dc.text = ""
        
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(4)
        sp.paragraph_format.line_spacing = 1.0

    return photo_num

def generar_word(data, fid):
    doc = Document()
    for s in doc.sections:
        s.top_margin = Cm(2.50)
        s.bottom_margin = Cm(2.50)
        s.left_margin = Cm(1.50)
        s.right_margin = Cm(2.34)
    
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Arial'
    style_normal.font.size = Pt(9.0)
    
    # 1. Top Logo Walmart
    logo_path = BASE_DIR / "static" / "img" / "walmart_logo.png"
    if logo_path.exists():
        lp = doc.add_paragraph()
        lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        lp.paragraph_format.space_before = Pt(0)
        lp.paragraph_format.space_after = Pt(4)
        lp.add_run().add_picture(str(logo_path), width=Inches(3.8))
    
    # 2. Store Title
    dg = data.get("datos_generales", {})
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tp.paragraph_format.space_before = Pt(4)
    tp.paragraph_format.space_after = Pt(10)
    store_name = dg.get("nombre_tienda", "").strip() or "MEMORIA DESCRIPTIVA"
    if not store_name.upper().startswith("BODEGA") and not store_name.upper().startswith("WALMART") and not store_name.upper().startswith("SAM"):
        title_text = f"BODEGA AURRERA “{store_name.upper()}”"
    else:
        title_text = store_name.upper()
    r_title = tp.add_run(title_text)
    r_title.font.name = "Arial"
    r_title.font.size = Pt(9.5)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0, 0, 255)

    # 3. Section: Datos Generales
    tipo_obra = dg.get("tipo_obra", "").strip() or "TRABAJOS DE REMODELACIÓN."
    p_to = doc.add_paragraph()
    p_to.paragraph_format.space_before = Pt(2)
    p_to.paragraph_format.space_after = Pt(2)
    r = p_to.add_run(f"TIPO DE OBRA:\t {tipo_obra.upper()}")
    r.font.name = "Arial"
    r.font.size = Pt(9.0)
    r.font.bold = True

    calle = dg.get("calle", "").strip()
    num = dg.get("numero", "").strip()
    col = dg.get("colonia", "").strip()
    cd = dg.get("ciudad", "").strip()
    edo = dg.get("estado_rep", "").strip()
    cp = dg.get("cp", "").strip()

    p_ub = doc.add_paragraph()
    p_ub.paragraph_format.space_before = Pt(2)
    p_ub.paragraph_format.space_after = Pt(2)
    r1 = p_ub.add_run(f"UBICACIÓN: \t   Calle:       {calle} \tNo.         {num}")
    r1.font.name = "Arial"
    r1.font.size = Pt(9.0)
    r1.font.bold = True

    p_col = doc.add_paragraph()
    p_col.paragraph_format.space_before = Pt(0)
    p_col.paragraph_format.space_after = Pt(2)
    r2 = p_col.add_run(f"   Colonia: {col}\t\tCiudad: {cd}")
    r2.font.name = "Arial"
    r2.font.size = Pt(9.0)
    r2.font.bold = True

    p_edo = doc.add_paragraph()
    p_edo.paragraph_format.space_before = Pt(0)
    p_edo.paragraph_format.space_after = Pt(4)
    r3 = p_edo.add_run(f"   Estado:   {edo}\t\t\tCP.        {cp}")
    r3.font.name = "Arial"
    r3.font.size = Pt(9.0)
    r3.font.bold = True

    for label, key in [
        ("SUPERFICIE DE CONSTRUCCIÓN", "sup_construccion"),
        ("SUPERFICIE A REMODELAR", "sup_remodelar"),
        ("SUPERFICIE TOTAL DEL PREDIO", "sup_total")
    ]:
        val = dg.get(key, "").strip()
        p_sup = doc.add_paragraph()
        p_sup.paragraph_format.space_before = Pt(0)
        p_sup.paragraph_format.space_after = Pt(2)
        r_lbl = p_sup.add_run(f"{label}: ----------------- ")
        r_lbl.font.name = "Arial"
        r_lbl.font.size = Pt(9.0)
        r_lbl.font.bold = True
        
        val_text = f"{val} M2." if val and not val.upper().endswith("M2") and not val.upper().endswith("M2.") else (val or "")
        r_val = p_sup.add_run(val_text)
        r_val.font.name = "Arial"
        r_val.font.size = Pt(9.0)

    # 4. Croquis de Localización
    p_cr = doc.add_paragraph()
    p_cr.paragraph_format.space_before = Pt(4)
    p_cr.paragraph_format.space_after = Pt(4)
    r_cr = p_cr.add_run("CROQUIS DE LOCALIZACIÓN:")
    r_cr.font.name = "Arial"
    r_cr.font.size = Pt(9.0)
    r_cr.font.bold = True

    cu = data.get("croquis", None)
    if cu and cu.get("url"):
        ip = resolve_img_path(cu["url"])
        if ip and ip.exists():
            try:
                t_cr = doc.add_table(rows=1, cols=1)
                t_cr.alignment = WD_TABLE_ALIGNMENT.CENTER
                tblPr = t_cr._tbl.tblPr
                borders = parse_xml(r'<w:tblBorders {} ><w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/></w:tblBorders>'.format(nsdecls('w')))
                tblPr.append(borders)
                cell_cr = t_cr.cell(0, 0)
                cell_cr.width = Cm(17.7)
                set_cell_margins(cell_cr, 80, 80, 80, 80)
                p_c = cell_cr.paragraphs[0]
                p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_c.add_run().add_picture(str(ip), width=Inches(5.5))
            except Exception as e:
                print("Error inserting croquis:", e)
                add_body_p(doc, "[Croquis no disponible]")
        else:
            add_body_p(doc, "[Croquis no disponible]")

    # 5. Descripción General
    add_banner(doc, "DESCRIPCIÓN GENERAL")
    dgen = data.get("descripcion_general", {})
    if dgen.get("texto"):
        add_body_p(doc, dgen["texto"], space_after=4)
    for item in dgen.get("lista_items", []):
        if item and item.strip():
            p_it = doc.add_paragraph(style="List Bullet")
            p_it.paragraph_format.space_before = Pt(0)
            p_it.paragraph_format.space_after = Pt(2)
            r_it = p_it.add_run(item.strip())
            r_it.font.name = "Arial"
            r_it.font.size = Pt(9.0)

    photo_num = 1

    # 6. Remodelación Exterior
    add_banner(doc, "DESCRIPCIÓN REMODELACIÓN EXTERIOR:")
    ext = data.get("remodelacion_exterior", {})
    for key, lbl in [
        ("cubierta", "Cubierta:"),
        ("estacionamiento", "Estacionamiento:"),
        ("anuncio_espectacular", "Anuncio Espectacular:"),
        ("fachadas", "Fachadas:"),
        ("anden", "Andén:"),
        ("area_servicio", "Área de Servicio:")
    ]:
        s = ext.get(key, {})
        txt = s.get("texto", "").strip()
        fotos = s.get("fotos", [])
        if txt or fotos:
            add_subtitle(doc, lbl)
            if txt:
                add_body_p(doc, txt, space_after=4)
            if fotos:
                photo_num = add_photo_grid(doc, fotos, fid, photo_counter_start=photo_num)

    # 7. Remodelación Interiores
    add_banner(doc, "DESCRIPCIÓN REMODELACIÓN INTERIORES:")
    intr = data.get("remodelacion_interiores", {})
    for key, lbl in [
        ("portico_acceso", "Pórtico de Acceso y Salida:"),
        ("oficinas_frontales", "Oficinas Frontales:"),
        ("sanitarios_clientes", "Sanitarios Clientes:"),
        ("piso_ventas", "Piso de Ventas:"),
        ("area_perecederos", "Área de Perecederos:"),
        ("comedor_asociados", "Comedor de Asociados:"),
        ("oficinas_posteriores", "Oficinas Posteriores:"),
        ("sanitarios_asociados", "Sanitarios de Asociados:"),
        ("acceso_personal", "Acceso de Personal y de Mercancía:"),
        ("facturacion_sistemas", "Facturación y Sistemas:"),
        ("trastienda", "Trastienda:")
    ]:
        s = intr.get(key, {})
        txt = s.get("texto", "").strip()
        fotos = s.get("fotos", [])
        if txt or fotos:
            add_subtitle(doc, lbl)
            if txt:
                add_body_p(doc, txt, space_after=4)
            if fotos:
                photo_num = add_photo_grid(doc, fotos, fid, photo_counter_start=photo_num)

    # 8. Giros de Negocio
    cm = data.get("giros_negocio", {}).get("consultorio_medico", "").strip()
    if cm:
        add_banner(doc, "GIROS DE NEGOCIO")
        add_subtitle(doc, "CONSULTORIO MÉDICO:")
        add_body_p(doc, cm, space_after=4)

    # 9. Descripción de los Trabajos
    tr2 = data.get("descripcion_trabajos", {})
    t_tr = tr2.get("texto", "").strip()
    t_area = tr2.get("area_trabajo", "").strip()
    t_afect = tr2.get("afectacion_estructural", "").strip()
    if t_tr or t_area or t_afect:
        add_banner(doc, "DESCRIPCIÓN DE LOS TRABAJOS:")
        if t_tr:
            add_body_p(doc, t_tr, space_after=4)
        if t_area:
            add_field(doc, "Área de Trabajo", t_area)
        if t_afect:
            add_field(doc, "Afectación en Elementos Estructurales", t_afect)

    # 10. Estructura
    et = data.get("estructura", {}).get("texto", "").strip()
    if et:
        add_banner(doc, "ESTRUCTURA")
        add_body_p(doc, et, space_after=4)

    # 11. Instalaciones
    inst = data.get("instalaciones", {})
    inst_items = [
        ("refrigeracion", "Sistema de refrigeración:"),
        ("aire", "Sistema de aire:"),
        ("electrica", "Instalación Eléctrica:"),
        ("hidro_sanitaria", "Instalación Hidro-Sanitaria:"),
        ("gas", "Instalación de Gas:"),
        ("filtrado", "Sistema de Filtrado:"),
        ("contra_incendio", "Instalación Contra Incendio:")
    ]
    has_inst = any(inst.get(k, "").strip() for k, _ in inst_items)
    if has_inst:
        add_banner(doc, "INSTALACIONES")
        for key, lbl in inst_items:
            t_val = inst.get(key, "").strip()
            if t_val:
                add_subtitle(doc, lbl)
                add_body_p(doc, t_val, space_after=4)

    # 12. Medidas de Seguridad
    seg = data.get("medidas_seguridad", {})
    s_gen = seg.get("texto", "").strip()
    s_cli = seg.get("medidas_clientes", "").strip()
    s_cons = seg.get("consideraciones", "").strip()
    if s_gen or s_cli or s_cons:
        add_banner(doc, "MEDIDAS DE SEGURIDAD")
        if s_gen:
            add_body_p(doc, s_gen, space_after=4)
        if s_cli:
            add_subtitle(doc, "Medidas para Clientes y Asociados:")
            add_body_p(doc, s_cli, space_after=4)
        if s_cons:
            add_subtitle(doc, "Consideraciones Adicionales:")
            add_body_p(doc, s_cons, space_after=4)

    # Footer
    ftr = doc.sections[0].footer
    fp2 = ftr.paragraphs[0] if ftr.paragraphs else ftr.add_paragraph()
    fp2.clear()
    fp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    fr2 = fp2.add_run(f"Memoria Descriptiva  |  Generado el {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    fr2.font.name = "Arial"
    fr2.font.size = Pt(8.0)
    fr2.font.color.rgb = RGBColor(128, 128, 128)

    op = OUTPUT_DIR / f"MEMORIA-{fid[:8]}.docx"
    doc.save(str(op))
    return op

if __name__=="__main__":
    port = int(os.environ.get("PORT", 5000))
    print("="*50); print(f"  Memoria Descriptiva - Puerto {port}")
    print("="*50)
    app.run(debug=False, host="0.0.0.0", port=port)
