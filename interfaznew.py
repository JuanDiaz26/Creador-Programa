# interfaz_programa.py — v53.0 (EXCEL ALINEACION MIXTA CORREGIDA + PDF PERFECTO)
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, Menu
import pandas as pd
import sqlite3, os, re, sys, traceback, json, difflib
from pathlib import Path
from datetime import date, datetime

# --- Excel ---
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.page import PageMargins

# --- Word ---
try:
    import docx
except ImportError:
    pass

# --- ReportLab ---
try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm, mm
    from reportlab.lib import colors
    from reportlab.platypus import Table, TableStyle, Paragraph
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT, TA_LEFT
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

# =============================================================================
#  SECCIÓN 1: CONFIGURACIÓN
# =============================================================================

def app_dir() -> Path:
    if getattr(sys, "frozen", False): return Path(sys.executable).parent
    return Path(__file__).parent

BASE_DIR = app_dir(); DATA_DIR = BASE_DIR / "data"; ASSETS_DIR = BASE_DIR / "assets"
DATA_DIR.mkdir(exist_ok=True); ASSETS_DIR.mkdir(exist_ok=True)
DB_PATH = DATA_DIR / "carreras.db"; NOMBRE_BD = str(DB_PATH)

programa_completo = [] 
indice_edicion = None 
DATOS_WORD_CACHED = []

# Variables GUI
entry_fecha = None; entry_nro_reunion = None; entry_nro_carrera = None; entry_horario = None; entry_premio = None
entry_distancia = None; entry_condicion = None; entry_premios_dinero = None; entry_apuesta = None
entry_incremento = None; entry_incremento_2 = None; combo_word = None; combo_dist = None
text_caballos = None; text_kilos = None; text_actuaciones = None; tabla_programa = None; lista_carreras = None
contador_carreras = None; btn_accion = None

# --- COLORES OFICIALES MANDILES ---
MANDILES = {
    "1": ("#d32f2f", "#ffffff"), "2": ("#ffffff", "#000000"), "3": ("#1565c0", "#ffffff"),
    "4": ("#fdd835", "#000000"), "5": ("#2e7d32", "#ffffff"), "6": ("#000000", "#fff200"),
    "7": ("#ef6c00", "#000000"), "8": ("#f48fb1", "#000000"), "9": ("#00bcd4", "#000000"),
    "10": ("#7b1fa2", "#ffffff"), "11": ("#9e9e9e", "#da2128"), "12": ("#71bf44", "#000000"),
    "13": ("#a05b09", "#fff200"), "14": ("#b71c1c", "#ffffff"), "15": ("#f3d19c", "#000000"),
    "default": ("#CCCCCC", "#000000")
}

RECORDS = {
    "700":  '700 metros - Record Dist.: 38" 2/5, Holly 28/02/2026',
    "800":  '800 metros - Record Dist.: 43" 1/5, Enloqueceme 14/12/2025',
    "1000": '1.000 metros - Record Dist.: 58" 2/5, Sarfo 23/07/2020',
    "1100": '1.100 metros - Record Dist.: 1\' 04" 3/5, Sold Out 29/08/2021',
    "1200": '1.200 metros - Record Dist.: 1\' 09" 4/5, Donald Music 28/11/2021',
    "1300": '1.300 metros - Record Dist.: 1\' 16" 2/5, Panatta 19/12/2021 - High Commander 08/10/2023 - Jolly Boy 30/11/2025',
    "1400": '1.400 metros - Record Dist.: 1\' 22" 3/5, Patani 29/10/2017 - Dipinto 28/11/2021',
    "1500": '1.500 metros - Record Dist.: 1\' 28" 4/5, Storm Chuck 24/08/2025',
    "1600": '1.600 metros - Record Dist.: 1\' 37" 2/5, Batman Crest 17/08/1975 - Teenek 11/06/2021 - Latan Craf 22/06/2025',
    "1800": '1.800 metros - Record Dist.: 1\' 51" 4/5, Sir Melody 26/11/2017',
    "2000": '2.000 metros - Record Dist.: 2\' 03" 2/5, Volynov 1978',
    "2200": '2.200 metros - Record Dist.: 2\'17"1/5, Frances Net 24/09/2016',
}

COLORS = {"bg": "#f0f2f5", "primary": "#248689", "accent": "#f16536", "card": "#ffffff", "ink": "#1f2937", "line": "#e5e7eb"}

# =============================================================================
#  SECCIÓN 2: LÓGICA
# =============================================================================

def formatear_cuerpos(valor):
    s = str(valor).strip()
    if any(x in s.upper() for x in ['CZA', 'PZO', 'HCO', 'S.A']): return s
    try:
        f = float(s); entero = int(f); dec = f - entero; frac = ""
        if abs(dec - 0.25) < 0.01: frac = "1/4"
        elif abs(dec - 0.50) < 0.01: frac = "1/2"
        elif abs(dec - 0.75) < 0.01: frac = "3/4"
        if entero > 0 and frac: res = f"{entero} {frac}"
        elif entero == 0 and frac: res = frac
        elif entero > 0: res = str(entero)
        else: res = str(s)
        return f"{res} cp"
    except: return s

def _inicializar_db_si_no_existe():
    conn = sqlite3.connect(NOMBRE_BD); c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS caballos (nombre TEXT PRIMARY KEY, padre_madre TEXT, pelo TEXT, ultima_edad TEXT, ultimo_peso TEXT, ultimo_jockey TEXT, caballeriza TEXT, cuidador TEXT, ultima_actuacion_externa TEXT, snapshot_programa_fecha DATE)''')
    c.execute('''CREATE TABLE IF NOT EXISTS actuaciones (id INTEGER PRIMARY KEY, fecha DATE, nombre_caballo TEXT, puesto_original INTEGER, puesto_final TEXT, jockey TEXT, cuerpos TEXT, ganador TEXT, segundo TEXT, margen TEXT, tiempo_ganador TEXT, pista TEXT, fue_distanciado BOOLEAN, observacion TEXT)''')
    conn.commit(); conn.close()

def conectar_y_cargar_datos():
    _inicializar_db_si_no_existe(); conn = sqlite3.connect(NOMBRE_BD)
    try: df_cab = pd.read_sql_query("SELECT * FROM caballos", conn); df_act = pd.read_sql_query("SELECT * FROM actuaciones", conn)
    except: df_cab = pd.DataFrame(); df_act = pd.DataFrame()
    finally: conn.close()
    if not df_cab.empty: df_cab = df_cab.rename(columns={'nombre':'Caballo', 'ultima_edad':'Edad', 'ultimo_peso':'Peso', 'ultimo_jockey':'Jockey-Descargo', 'padre_madre':'Padre - Madre', 'caballeriza':'Caballeriza', 'cuidador':'Cuidador', 'pelo':'Pelo'})
    if not df_act.empty: 
        df_act = df_act.rename(columns={'nombre_caballo':'Caballo', 'puesto_original':'Puesto Original', 'puesto_final':'Puesto Final', 'jockey':'Jockey', 'cuerpos':'Cuerpos al Ganador', 'ganador':'Ganador', 'segundo':'Segundo', 'margen':'Margen', 'tiempo_ganador':'Tiempo Ganador', 'pista':'Pista', 'fue_distanciado':'Fue Distanciado', 'fecha':'Fecha', 'observacion':'Observacion'})
        df_act['Fecha'] = pd.to_datetime(df_act['Fecha'], errors='coerce')
    return df_cab, df_act

def obtener_datos_caballo(nombre, db_cab, db_act):
    nombre = nombre.strip().upper()
    try: info = db_cab[db_cab['Caballo'] == nombre].iloc[0].to_dict()
    except: info = {'Caballo': nombre}
    
    acts = db_act[db_act['Caballo'] == nombre].sort_values(by='Fecha', ascending=True) 
    
    # --- 1. Calcular las siglas Locales (Tucumán) en orden cronológico ---
    ult_locales = []
    for _, row_act in acts.iterrows():
        ps = str(row_act['Puesto Final']).strip()
        tiempo = str(row_act.get('Tiempo Ganador', ''))
        
        if ps.isdigit():
            puesto = int(ps)
            if puesto >= 10:
                ult_locales.append('0')
            elif puesto == 1:
                # --- MAGIA DEL EXTRAOFICIAL (1e) ---
                es_extraoficial = False
                # Si no tiene el apóstrofe de los minutos (ej. 1'13")
                if "'" not in tiempo:
                    # Buscamos los segundos antes de la comilla doble (ej. 44")
                    m_sec = re.search(r"(\d{2})\"", tiempo)
                    if m_sec:
                        segundos = int(m_sec.group(1))
                        # Si los segundos están entre 36 y 50, es de 700 u 800 mts
                        if 36 <= segundos <= 50:
                            es_extraoficial = True
                
                ult_locales.append('1e' if es_extraoficial else '1')
                # -----------------------------------
            else:
                ult_locales.append(str(puesto))
        else:
            ult_locales.append('-')
            
    # --- 2. Traer las siglas Externas (Buenos Aires, etc) ---
    ext_str = str(info.get('ultima_actuacion_externa', '')).strip()
    ult_externas = [x.strip() for x in ext_str.split('-') if x.strip()] if ext_str and ext_str.lower() != 'nan' else []
    
    # --- 3. Combinar todo (Cronológicamente correcto) ---
    combined_full = ult_externas + ult_locales
    
    # --- 4. Aplicar lógica de prolijidad (Máximo 3 o 4) ---
    # Tomamos las últimas 4 por defecto
    total_ult = combined_full[-4:] if combined_full else []
    
    # Verificamos si en esas 4 hay siglas de afuera (letras que NO sean la 'e' de extraoficial)
    # Por ejemplo, si detecta 'P' o 'LP', tiene_afuera será True. Si solo hay '1e' o números, será False.
    tiene_afuera = any(c.isalpha() and c.lower() != 'e' for act in total_ult for c in act)
    
    # Si detecta que hay actuaciones de afuera en ese grupo, lo recorta a las 3 más nuevas
    if tiene_afuera:
        total_ult = combined_full[-3:] if combined_full else []
        
    cuatro = "-".join(total_ult) if total_ult else "Debuta"
    
    # Guardamos los datos en el diccionario final
    edad = info.get('Edad', '')
    info['E Kg'] = f"{edad} {info.get('Peso','')}".strip()
    info['4 Ult.'] = cuatro
    info['actuaciones'] = acts.tail(2) # Mostramos las 2 más recientes locales abajo
    info['texto_act_ext'] = str(info.get('texto_actuaciones_externas', '')).strip()
    
    return info

def cargar_word_entrada():
    f = filedialog.askopenfilename(filetypes=[("Archivos Word", "*.docx;*.doc")])
    if not f: return
    try: doc = docx.Document(f)
    except: messagebox.showerror("Error", "No se pudo leer el archivo."); return
    
    global DATOS_WORD_CACHED; DATOS_WORD_CACHED = []; curr = {}; capturing = False
    KEYWORDS = ("TURNO", "CLASICO", "CLÁSICO", "ESPECIAL", "HANDICAP", "GRAN PREMIO")
    for para in doc.paragraphs:
        txt = para.text.strip(); 
        if not txt: continue
        upper = txt.upper()
        if "LIQUIDARAN" in upper or "COMPUTAN" in upper or "INSCRIPCION" in upper: capturing = False; continue
        es_titulo = False
        if upper.startswith("PREMIO") and not upper.startswith("PREMIOS:"): es_titulo = True
        for k in KEYWORDS:
            if upper.startswith(k): es_titulo = True; break
        if es_titulo:
            if curr: DATOS_WORD_CACHED.append(curr)
            curr = {"nombre": txt, "distancia": "", "condicion_raw": "", "premios": ""}
            capturing = True
            m = re.search(r'(\d{1,2}[.]\d{3}|\d{3,4})\s*(?:m|mts|metros)', txt, re.I)
            if m: curr["distancia"] = m.group(1)
            continue
        if upper.startswith("PREMIOS:"): 
            if len(txt) > 120: continue 
            curr["premios"] = txt.split(':', 1)[1].strip(); capturing = False; continue
        if capturing and curr:
            if not curr["distancia"]:
                m = re.search(r'(\d{1,2}[.]\d{3}|\d{3,4})\s*(?:m|mts|metros)', txt, re.I)
                if m: curr["distancia"] = m.group(1)
                elif re.match(r'^\d{3,4}$', txt): curr["distancia"] = txt 
            if len(txt) > 10 and not re.match(r'^\d+$', txt): 
                if curr["condicion_raw"]: curr["condicion_raw"] += " " + txt
                else: curr["condicion_raw"] = txt
    if curr: DATOS_WORD_CACHED.append(curr)
    vals = [c.get("nombre", "Carrera") for c in DATOS_WORD_CACHED]; combo_word['values'] = vals
    if vals: combo_word.current(0); messagebox.showinfo("Cargado", f"{len(vals)} carreras detectadas."); aplicar_seleccion_word(None)
    else: messagebox.showwarning("Atención", "No se detectaron carreras.")

def aplicar_seleccion_word(e):
    idx = combo_word.current(); 
    if idx < 0: return
    d = DATOS_WORD_CACHED[idx]
    dist_orig = d.get("distancia",""); dist_key = dist_orig.replace('.', '').strip()
    entry_distancia.delete(0, tk.END)
    if dist_key in RECORDS: entry_distancia.insert(0, RECORDS[dist_key]); combo_dist.set(dist_key)
    else: entry_distancia.insert(0, dist_orig + " metros")
    entry_premios_dinero.delete(0, tk.END)
    try: dv = int(dist_key)
    except: dv = 0
    cat = "CAT. EXTRAOFICIAL" if dv <= 800 and dv > 0 else "CAT. INTERIOR"
    p_raw = d.get('premios','')
    if "COMPUTABLE" not in p_raw.upper(): entry_premios_dinero.insert(0, f"NO COMPUTABLE - {cat} - Premios: {p_raw}")
    else: entry_premios_dinero.insert(0, p_raw)
    entry_condicion.delete(0, tk.END); entry_condicion.insert(0, d.get("condicion_raw","").strip())

# =============================================================================
#  SECCIÓN 3: PERSISTENCIA
# =============================================================================

def accion_guardar_proyecto():
    if not programa_completo: return
    f = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Proyecto Programa", "*.json")])
    if not f: return
    estado = {"fecha": entry_fecha.get(), "reunion": entry_nro_reunion.get(), "carreras": programa_completo}
    try:
        with open(f, 'w', encoding='utf-8') as json_file: json.dump(estado, json_file, indent=4)
        messagebox.showinfo("Guardado", "Proyecto guardado.")
    except Exception as e: messagebox.showerror("Error", str(e))

def accion_cargar_proyecto():
    f = filedialog.askopenfilename(filetypes=[("Proyecto Programa", "*.json")])
    if not f: return
    try:
        with open(f, 'r', encoding='utf-8') as json_file: estado = json.load(json_file)
        entry_fecha.delete(0, tk.END); entry_fecha.insert(0, estado.get("fecha", ""))
        entry_nro_reunion.delete(0, tk.END); entry_nro_reunion.insert(0, estado.get("reunion", "22"))
        global programa_completo; programa_completo = estado.get("carreras", [])
        _refrescar_lista_carreras(); limpiar_formulario(); messagebox.showinfo("Cargado", "Proyecto cargado.")
    except Exception as e: messagebox.showerror("Error", str(e))

# =============================================================================
#  SECCIÓN 4: EXPORTAR PDF (v53)
# =============================================================================

def _clean_str(txt): return str(txt).replace('"', '').replace("Hs.", "").strip()
def _parse_money(txt):
    if not txt: return 0
    limpio = re.sub(r'[^\d]', '', str(txt)) # Solo digitos
    if not limpio: return 0
    return int(limpio)

def exportar_pdf(color_mode="digital"):
    if not HAS_REPORTLAB or not programa_completo: return
    tipo = "COLOR" if color_mode == "digital" else "BN"
    filepath = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=f"Programa_{tipo}.pdf")
    if not filepath: return

    try:
        # COLORES OFICIALES
        C_VERDE_OFICIAL = colors.HexColor("#248689")
        C_VERDE = C_VERDE_OFICIAL
        C_HEAD_BG = C_VERDE_OFICIAL if color_mode == "digital" else colors.white
        C_HEAD_TXT = colors.white if color_mode == "digital" else colors.black
        C_NARANJA = colors.HexColor("#ef6c00") 
        
        c = canvas.Canvas(filepath, pagesize=A4)
        W, H = A4; MX = 0.5 * cm; MY = 1.0 * cm
        styles = getSampleStyleSheet()
        
        # Estilos de Paragraph
        style_cell_center = ParagraphStyle('CellC', parent=styles['Normal'], fontName='Helvetica', fontSize=6.5, leading=7, alignment=TA_CENTER)
        style_cell_left = ParagraphStyle('CellL', parent=styles['Normal'], fontName='Helvetica', fontSize=6.5, leading=7, alignment=TA_LEFT)
        style_cell_right = ParagraphStyle('CellR', parent=styles['Normal'], fontName='Helvetica', fontSize=6.5, leading=7, alignment=TA_RIGHT)
        style_header = ParagraphStyle('HeadC', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=6.5, leading=7, alignment=TA_CENTER)
        style_cond = ParagraphStyle('Cond', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=9)
        style_legales = ParagraphStyle('Leg', parent=styles['Normal'], fontName='Helvetica', fontSize=7, leading=8, alignment=TA_JUSTIFY)
        
        LOGO_MAIN = ASSETS_DIR / "logo.png"; LOGO_WSP = ASSETS_DIR / "whatsapp.png"; LOGO_SOC = ASSETS_DIR / "redes.png"
        
        fecha_txt = entry_fecha.get().strip().upper()
        if not fecha_txt: fecha_txt = date.today().strftime("%d DE %B DE %Y").upper()
        nro_reunion = entry_nro_reunion.get().strip() or "22"

        def draw_institutional_header():
            y_top_box = H - 1.0*cm; h_top_box = 1.6*cm; w_box = W - 2*MX
            c.setStrokeColor(C_VERDE_OFICIAL); c.setLineWidth(2); c.rect(MX, y_top_box - h_top_box, w_box, h_top_box)
            if LOGO_MAIN.exists(): c.drawImage(str(LOGO_MAIN), MX + 0.3*cm, y_top_box - h_top_box + 0.1*cm, width=1.4*cm, height=1.4*cm, mask='auto', preserveAspectRatio=True)
            
            c.setFillColor(colors.black); c.setFont("Helvetica-BoldOblique", 16)
            c.drawCentredString(MX + w_box/2 + 1.0*cm, y_top_box - 0.7*cm, "HIPÓDROMO DE TUCUMÁN - PROGRAMA OFICIAL")
            
            c.setFillColor(C_NARANJA); c.setFont("Helvetica-Bold", 12)
            c.drawCentredString(MX + w_box/2 + 1.0*cm, y_top_box - 1.3*cm, f"REUNION Nº {nro_reunion} - {fecha_txt}")
            
            y_auth = y_top_box - h_top_box - 0.15*cm
            
            # Estilos de Comisión con Color Oficial y Negrita Real
            st_tit = ParagraphStyle('tit', fontName='Helvetica-Bold', fontSize=6.5, textColor=C_VERDE_OFICIAL, alignment=TA_CENTER)
            st_nom = ParagraphStyle('nom', fontName='Helvetica', fontSize=6.5, textColor=colors.black, alignment=TA_LEFT)
            
            # MAGIA 1: Reemplacé "" por Paragraph(" ", st_tit) para alinear perfectamente
            data_auth = [
                [Paragraph(" ", st_tit), Paragraph("<u>COMISIÓN DE CARRERAS</u>", st_tit), Paragraph("<u>VOCALES</u>", st_tit), Paragraph("<u>DELEGADO HIPODROMO</u>", st_tit)],
                [Paragraph("PRESIDENTE:", st_tit), Paragraph("Dr. Luis Alberto Gamboa", st_nom), Paragraph("Juan Ramon Rouges", st_nom), Paragraph("Estanislao Perez Garcia", st_nom)], 
                [Paragraph("VICE-PRESIDENTE:", st_tit), Paragraph("C.P.N Ernesto José Vidal Sanz", st_nom), Paragraph("Marcos Bruchmann", st_nom), ""], 
                [Paragraph("SECRETARIO:", st_tit), Paragraph("Ignacio Lopez Bustos", st_nom), Paragraph("Santiago Allende", st_nom), ""]
            ]
            t = Table(data_auth, colWidths=[3.2*cm, 6*cm, 4.5*cm, 6.3*cm])
            t.setStyle(TableStyle([
                ('BOX', (0,0), (-1,-1), 2, C_VERDE_OFICIAL), 
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), 
                ('LEFTPADDING', (0,0), (-1,-1), 1), ('RIGHTPADDING', (0,0), (-1,-1), 1),
                ('TOPPADDING', (0,0), (-1,-1), 1.5), ('BOTTOMPADDING', (0,0), (-1,-1), 1.5)
            ]))
            w_t, h_t = t.wrapOn(c, W, H); t.drawOn(c, MX, y_auth - h_t)
            
            txt_legal = "Admisión y permanencia: Las autoridades del Hipódromo de Tucumán ejercen la facultad de admisión y permanencia en las instalaciones del Hipódromo durante el desarrollo de la reunión hípica. Los profesionales y el público asistente se someten a las disposiciones del Reglamento General de Carreras y a las resoluciones de la Honorable Comisión de Carreras, cuyos fallos son inapelables. Los Boletos no cobrados solo se pagarán, los días de carreras de Tucumán y en el horario en que se desarrolle la reunión y tendrán validez, hasta 2 reuniones siguientes.-"
            y_leg = y_auth - h_t - 0.15*cm
            p = Paragraph(txt_legal, style_legales); w_leg, h_leg = p.wrap(w_box - 0.4*cm, 5*cm)
            c.setStrokeColor(C_VERDE_OFICIAL); c.setLineWidth(2)
            c.rect(MX, y_leg - h_leg - 0.2*cm, w_box, h_leg + 0.4*cm); p.drawOn(c, MX + 0.2*cm, y_leg - h_leg)
            
            y_box = y_leg - h_leg - 1.4*cm; h_warn = 1.1*cm
            c.setStrokeColor(C_VERDE_OFICIAL); c.setLineWidth(2)
            c.rect(MX, y_box, 9.2*cm, h_warn)
            c.setFillColor(colors.black); c.setFont("Helvetica-BoldOblique", 9)
            c.drawCentredString(MX + 4.6*cm, y_box + 0.65*cm, "El juego compulsivo es")
            c.drawCentredString(MX + 4.6*cm, y_box + 0.25*cm, "perjudicial para la salud.")
            
            c.rect(W - MX - 9.5*cm, y_box, 9.5*cm, h_warn)
            c.drawCentredString(W - MX - 4.75*cm, y_box + 0.65*cm, "Los retirados en las apuestas")
            c.drawCentredString(W - MX - 4.75*cm, y_box + 0.25*cm, "combinadas pasan al favorito.")
            
            return (y_box - 0.3*cm)

        def draw_race(carrera, x, y_start, width, idx_carrera):
            cab = carrera['cabecera']; h_head = 1.3*cm
            c.setFillColor(C_HEAD_BG); c.setStrokeColor(colors.black); c.setLineWidth(1)
            c.rect(x, y_start - h_head, width, h_head, fill=(color_mode=="digital"))
            
            c.setFillColor(C_HEAD_TXT); c.setFont("Helvetica-Bold", 17)
            c.drawString(x + 2*mm, y_start - 7.5*mm, f"{cab['nro_carrera']}º Carrera")
            clean_horario = _clean_str(cab['horario']).replace("Hs.", "")
            c.drawRightString(x + width - 2*mm, y_start - 7.5*mm, f"{clean_horario} Hs.")
            
            clean_premio = _clean_str(cab['premio']) 
            if clean_premio.upper().startswith("PREMIO"): clean_premio = clean_premio[6:].strip()
            c.setFont("Helvetica-Bold", 15)
            c.drawCentredString(x + width/2, y_start - 7.0*mm, f"PREMIO \"{clean_premio.upper()}\"")
            
            c.setFont("Helvetica-Bold", 8); full_dist = f"{cab['distancia']}"; dist_val = cab['distancia'].split()[0].replace('.','')
            if dist_val in RECORDS: full_dist = RECORDS[dist_val]
            c.drawCentredString(x + width/2, y_start - 11.5*mm, full_dist)
            
            # CONDICIÓN Saneada (sin |)
            clean_cond = cab['condicion'].replace("PREMIOS:", "").replace("|", " ").strip()
            p = Paragraph(clean_cond, style_cond)
            w_cond, h_cond = p.wrap(width, 3*cm) 
            y_curr = y_start - h_head - 2*mm
            p.drawOn(c, x, y_curr - h_cond)
            y_curr -= (h_cond + 4*mm) 
            
            y_top_caja_amarilla = y_curr + 2*mm 
            
            txt_premios = cab['premios_dinero'].replace("Premios:", "").strip()
            if "Premios:" in txt_premios: txt_premios = txt_premios.replace("Premios:", "")
            c.setFillColor(colors.black); c.setFont("Helvetica-Bold", 7.5)
            c.drawString(x, y_curr - 3*mm, txt_premios) 
            y_curr -= 3*mm 
            
            detalle_ap = cab['incremento_2']
            if detalle_ap: 
                c.setFont("Helvetica-Bold", 7); c.drawString(x, y_curr - 4*mm, detalle_ap.strip()); y_curr -= 4*mm
            
            bottom_left = y_curr
            txt_ap = cab['apuesta']; txt_inc = cab['incremento']
            box_w = 5.2*cm; box_x = x + width - box_w; box_h = 0.9*cm; center_box = box_x + (box_w/2)
            
            if txt_ap or txt_inc:
                if color_mode == "digital":
                    c.setFillColor(colors.lightyellow); c.setStrokeColor(colors.gold)
                    c.rect(box_x, y_top_caja_amarilla - box_h, box_w, box_h, fill=1, stroke=1)
                    c.setFillColor(colors.black)
                c.setFont("Helvetica-BoldOblique", 9)
                if txt_ap: c.drawCentredString(center_box, y_top_caja_amarilla - 4*mm, txt_ap)
                inc_val = _parse_money(cab['incremento'])
                if inc_val > 0: 
                    txt_inc_show = f"INCREMENTO: $ {inc_val:,.0f}".replace(",",".")
                    c.drawCentredString(center_box, y_top_caja_amarilla - 7.5*mm, txt_inc_show)
                    
            bottom_right = y_top_caja_amarilla - box_h
            y_curr = min(bottom_left, bottom_right) - 2*mm 
            h_info_block = (y_start - y_curr)
            
            # TABLA DE CABALLOS
            col_ws = [1.3*cm, 0.6*cm, 3.6*cm, 1.0*cm, 2.6*cm, 0.9*cm, 4.0*cm, 3.4*cm, 2.6*cm]
            headers_raw = ['4 Ult.', 'Nº', 'Caballo', 'Pelo', 'Jockey', 'E Kg', 'Padre-Madre', 'Caballeriza', 'Cuidador']
            headers_para = [Paragraph(h, style_header) for h in headers_raw]
            data = [headers_para]
            for row in carrera['tabla_caballos']:
                nro_raw = str(row[1]); key_mandil = "".join(filter(str.isdigit, nro_raw)); 
                if not key_mandil: key_mandil = "default"
                bg_hex, fg_hex = MANDILES.get(key_mandil, MANDILES['default'])
                nro_txt = f"<font color='{fg_hex}'><b>{nro_raw}</b></font>"
                if color_mode == "print": nro_txt = f"<b>{nro_raw}</b>"
                pm = Paragraph(str(row[6]), style_cell_left); caballeriza = Paragraph(str(row[7]), style_cell_left); cuidador = Paragraph(str(row[8]), style_cell_left)
                caballo = Paragraph(f"<b>{str(row[2])}</b>", style_cell_left); jockey = Paragraph(str(row[4]), style_cell_left)
                ult = Paragraph(str(row[0]), style_cell_right); nro = Paragraph(nro_txt, style_cell_center)
                pelo = Paragraph(str(row[3]), style_cell_center); ekg = Paragraph(str(row[5]), style_cell_center)
                data.append([ult, nro, caballo, pelo, jockey, ekg, pm, caballeriza, cuidador])
            t = Table(data, colWidths=col_ws, rowHeights=[0.55*cm] * len(data))
            
            ts = [('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('LEFTPADDING', (0,0), (-1,-1), 1), ('RIGHTPADDING', (0,0), (-1,-1), 1),
                ('LEFTPADDING', (2,0), (2,-1), 4), ('TOPPADDING', (0,0), (-1,-1), 0.5), ('BOTTOMPADDING', (0,0), (-1,-1), 0.5),
                ('ROWBACKGROUNDS', (1,0), (-1,-1), [colors.white])]
            for i, row in enumerate(carrera['tabla_caballos']):
                ridx = i + 1; nro_raw = str(row[1]); key_mandil = "".join(filter(str.isdigit, nro_raw))
                if not key_mandil: key_mandil = "default"
                bg_hex, _ = MANDILES.get(key_mandil, MANDILES['default']) 
                if color_mode == "print": bg_hex = "#ffffff"
                ts.append(('BACKGROUND', (1, ridx), (1, ridx), colors.HexColor(bg_hex)))
            t.setStyle(TableStyle(ts))
            w_t, h_t = t.wrapOn(c, width, H)
            y_curr -= h_t 
            t.drawOn(c, x, y_curr)
            
            # MAGIA 2: Respiro de solo 1mm con la tabla de arriba (antes 4mm)
            y_act = y_curr - 1*mm 
            
            # --- ÚLTIMAS ACTUACIONES CALIBRADAS (Matemática Pura) ---
            lines_act = carrera['actuaciones'].split('\n')
            count_lines = sum(1 for l in lines_act if l.strip())
            h_row_exacto = 5.5*mm # Salto exacto por renglón
            h_padding_total = 2*mm # Aire total (arriba+abajo) dentro de la caja gris
            
            h_acts = (count_lines * h_row_exacto) + h_padding_total 
            
            c.setFillColor(colors.whitesmoke); c.setStrokeColor(colors.lightgrey)
            c.rect(x, y_act - h_acts, width, h_acts, fill=1, stroke=1)
            
            tiene_derecha = any("||" in l for l in lines_act if l.strip())
            if tiene_derecha:
                c.setStrokeColor(C_VERDE_OFICIAL); c.setLineWidth(1.5)
                # Línea central contínua perfecta
                c.line(x + width/2, y_act - 1*mm, x + width/2, y_act - h_acts + 1*mm)
            
            # MAGIA 3 y 4: Sincronización de todas las filas. Todas miden y saltan 5.5mm
            # Empezamos el primer renglón a 1mm exacto del techo (la mitad del padding total)
            curr_y_txt = y_act - 1*mm 
            
            for l in lines_act:
                if not l.strip(): continue
                m = re.match(r'^(\d+[a-zA-Z]?)\s*[-\s]+(.*)', l)
                if m:
                    nro_raw, resto = m.groups(); key_mandil = "".join(filter(str.isdigit, nro_raw)) 
                    bg_hex, fg_hex = MANDILES.get(key_mandil, MANDILES['default'])
                    if color_mode == "print": bg_hex, fg_hex = "#ffffff", "#000000"
                    
                    c.setFillColor(colors.HexColor(bg_hex)); c.setStrokeColor(colors.black)
                    c.circle(x + 3.5*mm, curr_y_txt - 1.2*mm, 2.1*mm, fill=1, stroke=1) # Círculo un pelín más grande (2.1)
                    c.setFillColor(colors.HexColor(fg_hex)); c.setFont("Helvetica-Bold", 6.5)
                    c.drawCentredString(x + 3.5*mm, curr_y_txt - 2.5*mm, nro_raw)
                    
                    c.setFillColor(colors.black); c.setFont("Helvetica", 6.5) 
                    parts = resto.split("||"); izq = parts[0].strip(); der = parts[1].strip() if len(parts)>1 else ""
                    c.drawString(x + 8*mm, curr_y_txt - 2.5*mm, izq)
                    if der: c.drawString(x + width/2 + 3*mm, curr_y_txt - 2.5*mm, der)
                else: 
                    c.setFillColor(colors.black); c.setFont("Helvetica", 6.5)
                    c.drawString(x + 2*mm, curr_y_txt - 2.5*mm, l)
                
                c.setStrokeColor(colors.lightgrey); c.setLineWidth(0.5)
                c.line(x + 1*mm, curr_y_txt - 4.5*mm, x + width - 1*mm, curr_y_txt - 4.5*mm)
                
                curr_y_txt -= h_row_exacto # Salto exacto y uniforme
                
            return (h_info_block + h_t + h_acts + 0.8*cm)

        y_cursor = draw_institutional_header(); total_inc = 0; data_footer = []
        for i, car in enumerate(programa_completo):
            cab = car['cabecera']; monto1 = _parse_money(cab['incremento'])
            if monto1 > 0:
                total_inc += monto1
                nom_ap = cab['apuesta'].upper().replace("APUESTA", "").strip(); rango = 1
                if "CUATERNA" in nom_ap: rango=4
                elif "TRIPLO" in nom_ap: rango=3
                elif "QUINTUPLO" in nom_ap: rango=5
                elif "CADENA" in nom_ap: rango=6
                elif "DOBLE" in nom_ap: rango=2
                try: nro_start = int(cab['nro_carrera'])
                except: nro_start = 1
                end_nro = nro_start + rango - 1
                if rango == 1: c_str = f"{nro_start}º carrera"
                else: c_str = f"{nro_start}º y {end_nro}º carrera" if rango==2 else f"{nro_start}º a {end_nro}º carrera"
                data_footer.append([f"{nom_ap}: $ {monto1:,.0f}".replace(",",".") + f" ({c_str})"])

        def draw_footer_area(y_pos):
            fmt_tot = f"{total_inc:,.0f}".replace(",", ".")
            final_data = [[f"TOTAL INCREMENTOS Y POZOS: $ {fmt_tot}"]] + data_footer
            
            # MAGIA 5: Cuadro de resumen más grande (18.5cm) y con más padding (4)
            width_foot = 18.5*cm
            tf = Table(final_data, colWidths=[width_foot])
            ts_f = [
                ('BOX', (0,0), (-1,-1), 1, colors.darkgrey),
                ('BACKGROUND', (0,0), (0,0), colors.HexColor("#1f2937")), 
                ('TEXTCOLOR', (0,0), (0,0), colors.white),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('FONTNAME', (0,0), (0,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (0,0), 10), 
                ('BOTTOMPADDING', (0,0), (0,0), 4), ('TOPPADDING', (0,0), (0,0), 4), # Más aire
                ('FONTNAME', (0,1), (-1,-1), 'Helvetica-BoldOblique'),
                ('FONTSIZE', (0,1), (-1,-1), 8),
                ('BOTTOMPADDING', (0,1), (-1,-1), 2), ('TOPPADDING', (0,1), (-1,-1), 2), # Más aire
                ('ROWBACKGROUNDS', (1,0), (-1,-1), [colors.white, colors.whitesmoke])
            ]
            tf.setStyle(TableStyle(ts_f))
            w_f, h_f = tf.wrapOn(c, W, H)
            
            y_draw = MY + 1.2*cm 
            tf.drawOn(c, (W - width_foot)/2, y_draw) # Centrado perfecto
            
            y_img = 0.5*cm 
            if LOGO_WSP.exists(): c.drawImage(str(LOGO_WSP), MX, y_img, width=4.5*cm, height=1.3*cm, mask='auto', preserveAspectRatio=True)
            if LOGO_SOC.exists(): c.drawImage(str(LOGO_SOC), W - MX - 4.5*cm, y_img, width=4.5*cm, height=1.3*cm, mask='auto', preserveAspectRatio=True)
            return h_f + 1.5*cm

        if len(programa_completo) > 0: h_used = draw_race(programa_completo[0], MX, y_cursor, W - 2*MX, 1); y_cursor -= h_used
        draw_footer_area(MY + 1.0*cm); c.showPage(); y_cursor = H - MY
        for i, car in enumerate(programa_completo[1:], start=2):
            filas_cab = len(car['tabla_caballos']); lines_act = sum(1 for l in car['actuaciones'].split('\n') if l.strip())
            h_est = 3.0*cm + (filas_cab * 0.55*cm) + (lines_act * 0.55*cm) + 1.0*cm
            if y_cursor - h_est < MY: c.showPage(); y_cursor = H - MY
            h_used = draw_race(car, MX, y_cursor, W - 2*MX, i); y_cursor -= (h_used + 0.3*cm)
        c.save(); messagebox.showinfo("PDF Creado", f"Archivo generado: {filepath}")
    except Exception as e: traceback.print_exc(); messagebox.showerror("Error PDF", str(e))

# =============================================================================
#  SECCIÓN 5: EXCEL (FINAL v53 - ALINEACION CORREGIDA)
# =============================================================================

def exportar_programa_excel():
    if not programa_completo: messagebox.showwarning("Vacío", "No hay datos."); return
    fp = filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel Workbook","*.xlsx")])
    if not fp: return
    wb=Workbook(); ws=wb.active; ws.title="Programa"; ws.page_margins=PageMargins(left=0.25,right=0.25,top=0.75,bottom=0.75); r=1; thin=Side(style="thin"); med=Side(style="medium")
    for c in programa_completo:
        cab=c['cabecera']; ws.merge_cells(f'C{r}:I{r}'); ws[f'C{r}'].value=cab['premio'].upper(); ws[f'C{r}'].font=Font(name='Tahoma', size=15,bold=True); ws[f'C{r}'].alignment=Alignment(horizontal='center', vertical='center')
        ws.merge_cells(f'A{r}:B{r}'); ws.cell(row=r,column=1,value=f"{cab['nro_carrera']}º Carrera").fill=PatternFill("solid","000000"); ws.cell(row=r,column=1).font=Font(name='Arial Narrow', size=11, color="FFFFFF",bold=True); ws.cell(row=r,column=1).alignment = Alignment(horizontal='center', vertical='center') 
        ws.cell(row=r,column=10,value=cab['horario']).fill=PatternFill("solid","000000"); ws.cell(row=r,column=10).font=Font(name='Arial Narrow', size=11, color="FFFFFF",bold=True); ws.cell(row=r,column=10).alignment = Alignment(horizontal='center', vertical='center'); r+=1
        ws.merge_cells(f'A{r}:J{r}'); ws.cell(row=r,column=1,value=cab['distancia']).alignment=Alignment(horizontal='center'); ws.cell(row=r,column=1).font=Font(name='Utsaah', size=9, bold=True); r+=1
        condicion=c['cabecera']['condicion']; lineas=[x.strip() for x in condicion.split('|')] or [""]; 
        for lin in lineas: ws.merge_cells(f'A{r}:J{r}'); ws.cell(row=r,column=1,value=lin).alignment=Alignment(wrap_text=True); ws.cell(row=r,column=1).font=Font(name='Utsaah', size=7); r+=1
        ws.merge_cells(f'A{r}:H{r}'); ws.cell(row=r,column=1,value=cab['premios_dinero']); ws.cell(row=r,column=1).font=Font(name='Arial Narrow', size=8, bold=True)
        ws.merge_cells(f'I{r}:J{r}'); ws.cell(row=r,column=9,value=cab['apuesta']); ws.cell(row=r,column=9).font=Font(name='Arial Black', size=9, bold=True, italic=True); ws.cell(row=r,column=9).alignment=Alignment(horizontal='center',vertical='center'); r+=1
        ws.merge_cells(f'A{r}:H{r}'); ws.cell(row=r,column=1,value=cab['incremento_2']); ws.cell(row=r,column=1).font=Font(name='Arial Narrow', size=8, bold=True)
        inc_val = _parse_money(cab['incremento'])
        if inc_val > 0: ws.merge_cells(f'I{r}:J{r}'); ci=ws.cell(row=r,column=9,value=f"INCREMENTO: $ {inc_val:,.0f}".replace(",",".")); ci.font=Font(name='Arial Black',size=9,bold=True,italic=True); ci.alignment=Alignment(horizontal='center',vertical='center')
        else: r+=1 
        r+=1; fila_inicio_tabla=r; headers=['4 Ult.','Nº','Caballo','Pelo','Jockey','E Kg','Padre-Madre','','Caballeriza','Cuidador']; ws.merge_cells(f'G{r}:H{r}'); ws.cell(row=r,column=7).value='Padre - Madre'
        for col,h in enumerate(headers,1): 
            if col not in (7,8): ws.cell(row=r,column=col,value=h).font=Font(name='Calibri', size=8, bold=True)
        r+=1
        for row in c['tabla_caballos']:
            ws.merge_cells(f'G{r}:H{r}')
            for i in range(6): ws.cell(row=r,column=i+1,value=row[i])
            ws.cell(row=r,column=7,value=row[6]); ws.cell(row=r,column=9,value=row[7]); ws.cell(row=r,column=10,value=row[8]); r+=1
        fila_inicio_act = r 
        for l in c['actuaciones'].split('\n'):
            if l.strip():
                if "Debutante" in l: pass
                elif " - " not in l[-5:]: l += " - PN" 
                parts = l.split("||"); part1 = parts[0].strip(); rec = parts[1].strip() if len(parts) > 1 else ""; m = re.match(r'^(\d+)[-\s]+(.*)', part1)
                num_x = int(m.group(1)) if m else 0; ant = m.group(2).strip() if m else part1
                ws.cell(row=r,column=1,value=num_x); ws.merge_cells(f'B{r}:F{r}'); ws.cell(row=r,column=2,value=ant); ws.merge_cells(f'G{r}:J{r}'); ws.cell(row=r,column=7,value=rec); r+=1
        fila_fin = r - 1
        for row in ws.iter_rows(min_row=fila_inicio_tabla, max_row=fila_fin, min_col=1, max_col=10):
             for cell in row:
                 b=Border(left=med,right=med,top=thin,bottom=thin) 
                 if cell.row == fila_inicio_tabla: b.top=med
                 if cell.row == fila_fin: b.bottom=med
                 if cell.column == 1: b.left=med
                 if cell.column == 10: b.right=med
                 if cell.row == fila_inicio_act - 1: b.bottom=med
                 cell.border=b
                 # --- ALINEACION VERTICAL "MIDDLE" SIEMPRE + HORIZONTAL MIXTA ---
                 h_align = 'center' # Default horizontal
                 if cell.row >= fila_inicio_act: # Actuaciones
                     if cell.column in (2,7): h_align = 'left'
                     cell.font=Font(name='Calibri',size=7)
                     if cell.column == 1: cell.font=Font(name='Calibri',size=8,bold=True)
                 elif cell.row == fila_inicio_tabla: # Headers
                     cell.font=Font(name='Calibri',size=8,bold=True)
                 else: # Competidores
                     # 4Ult (1) Right; N(2), Pelo(4), EKg(6) Center; Resto Left
                     if cell.column == 1: h_align = 'right'
                     elif cell.column in (3, 5, 7, 9, 10): h_align = 'left' # Caballo, Jockey, Padre, Cab, Cui
                     is_bold = (cell.column == 2 or cell.column == 3); cell.font=Font(name='Calibri',size=8, bold=is_bold)
                 
                 cell.alignment = Alignment(horizontal=h_align, vertical='center') # SIEMPRE VERTICAL CENTER
        r+=1
    for k,w in dict(A=6.1,B=3.9,C=15.6,D=5.1,E=13.7,F=3.7,G=9,H=12.4,I=14.6,J=13.7).items(): ws.column_dimensions[k].width=w
    wb.save(fp); messagebox.showinfo("Listo","Excel Guardado")

# =============================================================================
#  SECCIÓN 6: UI CALLBACKS
# =============================================================================

def editar_jockey(event):
    item_id = tabla_programa.identify_row(event.y); col_id = tabla_programa.identify_column(event.x)
    if not item_id: return
    if col_id == '#5': # Columna Jockey
        vals = list(tabla_programa.item(item_id, 'values')); old_jockey = vals[4]
        new_jockey = simpledialog.askstring("Editar Jockey", f"Jockey actual: {old_jockey}\nNuevo Jockey:")
        if new_jockey is not None: vals[4] = new_jockey.strip(); tabla_programa.item(item_id, values=vals)

def generar_programa_en_tabla():
    if db_caballos.empty: messagebox.showwarning("BD vacía", "No hay datos."); return
    for i in tabla_programa.get_children(): tabla_programa.delete(i)
    text_actuaciones.delete("1.0", tk.END)
    
    nombres = [n.strip() for n in text_caballos.get("1.0", tk.END).strip().split('\n') if n.strip()]
    kilos_manuales = [k.strip() for k in text_kilos.get("1.0", tk.END).strip().split('\n') if k.strip()]
    numero = 1; ultimo = None
    
    # El "enumerate" es clave acá para que la "i" vaya contando 0, 1, 2... y coincida con los kilos
    for i, nombre_in in enumerate(nombres):
        es_a = False; nombre = nombre_in
        if re.search(r'\(\s*a\s*\)\s*$', nombre_in, flags=re.I): es_a = True; nombre = re.sub(r'\(\s*a\s*\)\s*$', '', nombre_in, flags=re.I).strip()
        
        # --- MEJORA: "¿Quisiste decir?" ---
        nombre_upper = nombre.upper()
        if not db_caballos.empty and nombre_upper not in db_caballos['Caballo'].values:
            lista_nombres_db = db_caballos['Caballo'].astype(str).tolist()
            sugerencias = difflib.get_close_matches(nombre_upper, lista_nombres_db, n=1, cutoff=0.7)
            if sugerencias:
                sugerencia = sugerencias[0]
                resp = messagebox.askyesno("Posible error de tipeo", f"No se encontró a '{nombre_upper}' en la base de datos.\n\n¿Quisiste decir '{sugerencia}'?")
                if resp:
                    nombre = sugerencia
        # -----------------------------------
        
        datos = obtener_datos_caballo(nombre, db_caballos, db_actuaciones)
        
        # --- MEJORA: KILOS MANUALES ---
        kilo_asignado = kilos_manuales[i] if i < len(kilos_manuales) else datos.get('Peso', '')
        datos['E Kg'] = f"{datos.get('Edad', '')} {kilo_asignado}".strip()
        # ------------------------------
        
        if es_a and ultimo is not None: nro = f"{ultimo}a"
        else: nro = str(numero); ultimo = numero
        if not es_a: numero += 1
        
        datos['Jockey-Descargo'] = ""; datos['Nº'] = nro; datos['Caballo'] = nombre.upper()
        tabla_programa.insert('', tk.END, values=[datos.get(c, '') for c in ['4 Ult.','Nº','Caballo','Pelo','Jockey-Descargo','E Kg','Padre - Madre','Caballeriza','Cuidador']])
        
        acts = datos.get('actuaciones')
        lineas_locales = []
        
        # --- 1. Armamos los textos Locales ---
        if acts is not None and not acts.empty:
            for _, a in acts.iterrows():
                f_fmt = a['Fecha'].strftime('%d/%m/%y') if pd.notna(a['Fecha']) else ''
                if str(a['Puesto Final']).strip().upper() == 'NC': 
                    lineas_locales.append(f"{f_fmt} - No Corrió.")
                    continue
                jk_full = str(a.get('Jockey',''))
                jk = f"{jk_full.split()[1][:1]}. {jk_full.split()[0]}" if len(jk_full.split())>1 else jk_full
                dist_txt = " (Dist.)" if str(a.get('Puesto Original')) != str(a.get('Puesto Final')) else ""
                pista = a.get('Pista', 'PN')
                if not pista: pista = 'PN'
                
                if str(a.get('Puesto Original')).strip() in ['1','1.0']:
                    margen = formatear_cuerpos(a.get('Margen',''))
                    lineas_locales.append(f"{f_fmt} - {jk} - 1º gan x {margen} a {str(a.get('Segundo','')).title()} - {a.get('Tiempo Ganador','')}{dist_txt} - {pista}")
                else:
                    cuerpos = formatear_cuerpos(a.get('Cuerpos al Ganador',''))
                    lineas_locales.append(f"{f_fmt} - {jk} - {a.get('Puesto Original')}º a {cuerpos} de {str(a.get('Ganador','')).title()} - {a.get('Tiempo Ganador','')}{dist_txt} - {pista}")

        # --- 2. Traemos los textos Externos ---
        texto_ext = datos.get('texto_act_ext', '')
        lineas_ext = [l.strip() for l in texto_ext.split('\n') if l.strip() and l.strip().lower() != 'nan']
        
        # OJO ACÁ: Ya no usamos reverse() en las de afuera, ni al final.
        # Simplemente sumamos la lista de afuera con la de Tucumán.
        # Como Tucumán ya viene ordenada cronológicamente (vieja a nueva) desde obtener_datos_caballo,
        # al sumar [viejas_afuera] + [locales_Tucuman], quedan en el orden perfecto.
        
        todas_las_lineas = lineas_ext + lineas_locales
        
        # Nos quedamos con las 2 más recientes (que ahora están al final de la lista)
        todas_las_lineas = todas_las_lineas[-2:] if todas_las_lineas else []
        
        if not todas_las_lineas:
            bloque = "Debutante"
        else:
            # Las unimos en el orden en el que están (vieja izquierda || nueva derecha)
            bloque = "   ||   ".join(todas_las_lineas)
            
        text_actuaciones.insert(tk.END, f"{nro}  {bloque}\n")

def obtener_datos_formulario():
    rows = [tabla_programa.item(i)['values'] for i in tabla_programa.get_children()]
    return {"cabecera": {"nro_carrera": entry_nro_carrera.get(), "premio": entry_premio.get(), "horario": entry_horario.get(), "distancia": entry_distancia.get(), "condicion": entry_condicion.get(), "premios_dinero": entry_premios_dinero.get(), "apuesta": entry_apuesta.get(), "incremento": entry_incremento.get(), "incremento_2": entry_incremento_2.get()}, "tabla_caballos": rows, "actuaciones": text_actuaciones.get("1.0", tk.END).strip()}

def limpiar_formulario():
    for e in [entry_nro_carrera, entry_premio, entry_horario, entry_distancia, entry_condicion, entry_premios_dinero, entry_apuesta, entry_incremento, entry_incremento_2]: e.delete(0, tk.END)
    text_caballos.delete("1.0", tk.END); text_kilos.delete("1.0", tk.END); text_actuaciones.delete("1.0", tk.END)
    for i in tabla_programa.get_children(): tabla_programa.delete(i)
    global indice_edicion; indice_edicion = None; btn_accion.config(text="Añadir Carrera")

def guardar_o_anadir_carrera():
    if not tabla_programa.get_children(): return
    data = obtener_datos_formulario(); global indice_edicion
    if indice_edicion is not None: programa_completo[indice_edicion] = data; messagebox.showinfo("OK", "Carrera Actualizada")
    else: programa_completo.append(data); messagebox.showinfo("OK", "Carrera Añadida")
    _refrescar_lista_carreras(); limpiar_formulario()

def cargar_carrera_para_editar():
    sel = lista_carreras.curselection(); 
    if not sel: return
    idx = int(sel[0]); limpiar_formulario(); global indice_edicion; indice_edicion = idx; btn_accion.config(text="Guardar Cambios")
    data = programa_completo[idx]; cab = data['cabecera']
    entry_nro_carrera.insert(0, cab['nro_carrera']); entry_premio.insert(0, cab['premio']); entry_horario.insert(0, cab['horario']); entry_distancia.insert(0, cab['distancia']); entry_condicion.insert(0, cab['condicion']); entry_premios_dinero.insert(0, cab['premios_dinero']); entry_apuesta.insert(0, cab['apuesta']); entry_incremento.insert(0, cab['incremento']); entry_incremento_2.insert(0, cab['incremento_2'])
    for row in data['tabla_caballos']: tabla_programa.insert('', tk.END, values=row); text_caballos.insert(tk.END, row[2] + "\n"); ekg_partes = str(row[5]).strip().split(maxsplit=1)
    kilo_val = ekg_partes[1] if len(ekg_partes) > 1 else ""
    text_kilos.insert(tk.END, kilo_val + "\n")
    text_actuaciones.insert(tk.END, data['actuaciones']); messagebox.showinfo("Editando", f"Editando carrera {cab['nro_carrera']}.")

def eliminar_carrera():
    sel = lista_carreras.curselection(); 
    if not sel: return
    programa_completo.pop(int(sel[0])); _refrescar_lista_carreras(); limpiar_formulario()

def _refrescar_lista_carreras():
    lista_carreras.delete(0, tk.END)
    for c in programa_completo: lista_carreras.insert(tk.END, f"{c['cabecera']['nro_carrera']}º - {c['cabecera']['premio']}")
    contador_carreras.set(f"Carreras: {len(programa_completo)}")

def accion_reset_db():
    if messagebox.askyesno("Reset", "Borrar DB?"):
        if os.path.exists(NOMBRE_BD): os.remove(NOMBRE_BD)
        _inicializar_db_si_no_existe(); global db_caballos, db_actuaciones; db_caballos, db_actuaciones = conectar_y_cargar_datos(); messagebox.showinfo("Info", "DB Reiniciada")

def accion_importar_programa(): messagebox.showinfo("OK", "Programa importado (Simulado)")
def accion_importar_resultados(): messagebox.showinfo("OK", "Resultados importados (Simulado)")

# =============================================================================
#  SECCIÓN 7: STARTUP
# =============================================================================

db_caballos, db_actuaciones = conectar_y_cargar_datos()
root = tk.Tk(); root.title("Gestión de Programas Hípicos v53.0"); root.configure(bg=COLORS["bg"])
try: root.iconbitmap(str(ASSETS_DIR/"programa.ico"))
except: pass

class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(self, bg=COLORS["bg"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_window = ttk.Frame(canvas, style="Card.TFrame")
        self.scrollable_window.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        self.window_id = canvas.create_window((0, 0), window=self.scrollable_window, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set); canvas.pack(side="left", fill="both", expand=True); scrollbar.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units")); canvas.bind("<Configure>", lambda e: canvas.itemconfig(self.window_id, width=e.width))

head = ttk.Frame(root, style="Card.TFrame", padding=15); head.pack(side=tk.TOP, fill=tk.X)
ttk.Label(head, text="GENERADOR DE PROGRAMAS PROFESIONAL", font=("Segoe UI", 16, "bold"), foreground=COLORS["primary"]).pack(side=tk.LEFT)
ttk.Button(head, text="Cargar desde Word", command=cargar_word_entrada).pack(side=tk.RIGHT)

foot = ttk.Frame(root, padding=15); foot.pack(side=tk.BOTTOM, fill=tk.X) 
contador_carreras = tk.StringVar(value="Carreras: 0"); ttk.Label(foot, textvariable=contador_carreras, font=("Segoe UI", 10)).pack(side=tk.LEFT)
ttk.Button(foot, text="PDF (Color)", command=lambda: exportar_pdf("digital")).pack(side=tk.RIGHT, padx=5)
ttk.Button(foot, text="PDF (B/N)", command=lambda: exportar_pdf("print")).pack(side=tk.RIGHT, padx=5)
ttk.Button(foot, text="Excel", command=exportar_programa_excel).pack(side=tk.RIGHT, padx=5)

main_scroll = ScrollableFrame(root); main_scroll.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
main_content = main_scroll.scrollable_window

form_container = ttk.Frame(main_content, padding=10); form_container.pack(fill=tk.X)
f1 = ttk.LabelFrame(form_container, text="Información de Carrera", padding=15); f1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)

ttk.Label(f1, text="Fecha Reunión (PDF):", style="Field.TLabel", foreground="red").grid(row=0, column=0, sticky="w", pady=5)
entry_fecha = ttk.Entry(f1); entry_fecha.grid(row=0, column=1, columnspan=3, sticky="we", pady=5)
entry_fecha.insert(0, date.today().strftime("%d DE %B DE %Y").upper())

# --- NUEVO CAMPO REUNION ---
ttk.Label(f1, text="Nº Reunión (PDF):", style="Field.TLabel", foreground="blue").grid(row=1, column=0, sticky="w", pady=5)
entry_nro_reunion = ttk.Entry(f1, width=10); entry_nro_reunion.grid(row=1, column=1, sticky="w", pady=5); entry_nro_reunion.insert(0, "22") 

ttk.Label(f1, text="Selección Word:", style="Field.TLabel").grid(row=2, column=0, sticky="w", pady=5)
combo_word = ttk.Combobox(f1, width=35, state="readonly"); combo_word.grid(row=2, column=1, columnspan=3, sticky="we", pady=5); combo_word.bind("<<ComboboxSelected>>", aplicar_seleccion_word)

ttk.Label(f1, text="Nº Carrera:", style="Field.TLabel").grid(row=3, column=0, sticky="w", pady=5); entry_nro_carrera = ttk.Entry(f1, width=10); entry_nro_carrera.grid(row=3, column=1, sticky="w", pady=5, padx=5)
ttk.Label(f1, text="Horario:", style="Field.TLabel").grid(row=3, column=2, sticky="w", pady=5); entry_horario = ttk.Entry(f1, width=15); entry_horario.grid(row=3, column=3, sticky="w", pady=5, padx=5)
ttk.Label(f1, text="Premio:", style="Field.TLabel").grid(row=4, column=0, sticky="w", pady=5); entry_premio = ttk.Entry(f1); entry_premio.grid(row=4, column=1, columnspan=3, sticky="we", pady=5)
ttk.Label(f1, text="Distancia:", style="Field.TLabel").grid(row=5, column=0, sticky="w", pady=5)
dist_var = tk.StringVar()
def _on_dist(*_): 
    k = dist_var.get().strip()
    if k in RECORDS: entry_distancia.delete(0, tk.END); entry_distancia.insert(0, RECORDS[k])
combo_dist = ttk.Combobox(f1, width=8, values=list(RECORDS.keys()), textvariable=dist_var); combo_dist.bind("<<ComboboxSelected>>", _on_dist); combo_dist.grid(row=5, column=1, sticky="w", pady=5, padx=5); entry_distancia = ttk.Entry(f1); entry_distancia.grid(row=5, column=2, columnspan=2, sticky="we", pady=5)
ttk.Label(f1, text="Condición (Usar | para salto):", style="Field.TLabel").grid(row=6, column=0, sticky="w", pady=5); entry_condicion = ttk.Entry(f1); entry_condicion.grid(row=6, column=1, columnspan=3, sticky="we", pady=5)
ttk.Label(f1, text="Premios ($):", style="Field.TLabel").grid(row=7, column=0, sticky="w", pady=5); entry_premios_dinero = ttk.Entry(f1); entry_premios_dinero.grid(row=7, column=1, columnspan=3, sticky="we", pady=5)

f2 = ttk.LabelFrame(form_container, text="Apuestas y Caballos", padding=15); f2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)
ttk.Label(f2, text="Apuesta (Título):", style="Field.TLabel").grid(row=0, column=0, sticky="w", pady=5); entry_apuesta = ttk.Entry(f2); entry_apuesta.grid(row=0, column=1, sticky="we", pady=5, padx=5)
ttk.Label(f2, text="Incremento ($):", style="Field.TLabel").grid(row=1, column=0, sticky="w", pady=5); entry_incremento = ttk.Entry(f2, width=15); entry_incremento.grid(row=1, column=1, sticky="w", pady=5, padx=5)
ttk.Label(f2, text="Detalle Apuestas:", style="Field.TLabel").grid(row=2, column=0, sticky="w", pady=5); entry_incremento_2 = ttk.Entry(f2, width=15); entry_incremento_2.grid(row=2, column=1, sticky="w", pady=5, padx=5)
ttk.Label(f2, text="Pegar Lista Caballos:").grid(row=3, column=0, sticky="nw", pady=5)
text_caballos = tk.Text(f2, height=6, width=22) # Lo achicamos a 22 de ancho
text_caballos.grid(row=3, column=1, rowspan=3, sticky="we", pady=5, padx=5)

ttk.Label(f2, text="Kilos:").grid(row=3, column=2, sticky="nw", pady=5)
text_kilos = tk.Text(f2, height=6, width=8) # Cajita nueva para los kilos
text_kilos.grid(row=3, column=3, rowspan=3, sticky="we", pady=5, padx=5)

btn_box = ttk.Frame(main_content, padding=10); btn_box.pack(fill=tk.X)
ttk.Button(btn_box, text="1. Procesar Tabla (Verificar)", command=generar_programa_en_tabla).pack(side=tk.LEFT, padx=10)
btn_accion = ttk.Button(btn_box, text="Añadir Carrera", command=guardar_o_anadir_carrera); btn_accion.pack(side=tk.LEFT, padx=10)
ttk.Button(btn_box, text="Limpiar Formulario", command=limpiar_formulario).pack(side=tk.LEFT, padx=10)

paned = ttk.PanedWindow(main_content, orient=tk.HORIZONTAL); paned.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
frame_left = ttk.Frame(paned); paned.add(frame_left, weight=4)
cols = ['4 Ult.','Nº','Caballo','Pelo','Jockey-Descargo','E Kg','Padre - Madre','Caballeriza','Cuidador']
tabla_programa = ttk.Treeview(frame_left, columns=cols, show='headings', height=10); 
for c in cols: tabla_programa.heading(c, text=c); tabla_programa.column(c, width=90)
tabla_programa.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
tabla_programa.bind("<Double-1>", editar_jockey)

frame_acts = ttk.Frame(frame_left); frame_acts.pack(fill=tk.BOTH, expand=True)
ttk.Label(frame_acts, text="Actuaciones Generadas (Editable):", style="Field.TLabel").pack(anchor="w")
scroll_acts = ttk.Scrollbar(frame_acts); scroll_acts.pack(side=tk.RIGHT, fill=tk.Y)
text_actuaciones = tk.Text(frame_acts, height=12, yscrollcommand=scroll_acts.set); text_actuaciones.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scroll_acts.config(command=text_actuaciones.yview)

frame_right = ttk.Frame(paned); paned.add(frame_right, weight=1)
ttk.Label(frame_right, text="Carreras en Programa:", style="Field.TLabel").pack(anchor="w", pady=5); lista_carreras = tk.Listbox(frame_right); lista_carreras.pack(fill=tk.BOTH, expand=True, padx=5)
ttk.Button(frame_right, text="✏️ Cargar para Editar", command=cargar_carrera_para_editar).pack(fill=tk.X, pady=5, padx=5); ttk.Button(frame_right, text="🗑️ Eliminar Seleccionada", command=eliminar_carrera).pack(fill=tk.X, pady=5, padx=5)

menubar = Menu(root); root.config(menu=menubar)
m_archivo = Menu(menubar, tearoff=0); menubar.add_cascade(label="Archivo", menu=m_archivo)
m_archivo.add_command(label="💾 Guardar Proyecto", command=accion_guardar_proyecto)
m_archivo.add_command(label="📂 Cargar Proyecto", command=accion_cargar_proyecto)
m_archivo.add_separator(); m_archivo.add_command(label="Salir", command=root.quit)

m_db = Menu(menubar, tearoff=0); menubar.add_cascade(label="Base de Datos", menu=m_db)
m_db.add_command(label="Importar Excel PROGRAMA", command=accion_importar_programa); m_db.add_command(label="Importar Excel RESULTADOS", command=accion_importar_resultados); m_db.add_separator(); m_db.add_command(label="⚠️ Resetear DB", command=accion_reset_db)

w, h = 1250, 850; root.geometry(f"{w}x{h}"); root.minsize(1100, 700); root.mainloop()