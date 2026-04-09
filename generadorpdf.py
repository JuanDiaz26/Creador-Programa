import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
import os
from pathlib import Path
import re
import traceback

# =============================================================================
# CONFIGURACIÓN DE RUTAS Y ASSETS
# =============================================================================
BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"
LOGO_MAIN = ASSETS_DIR / "logo.png"
LOGO_WSP = ASSETS_DIR / "whatsapp.png"
LOGO_SOC = ASSETS_DIR / "redes.png"

C_VERDE_OFICIAL = colors.HexColor("#248689")
C_NARANJA = colors.HexColor("#ef6c00")

MANDILES = {
    "1": ("#ff0000", "#ffffff"), "2": ("#ffffff", "#000000"),
    "3": ("#0000ff", "#ffffff"), "4": ("#ffff00", "#000000"),
    "5": ("#008000", "#ffffff"), "6": ("#000000", "#ffffff"),
    "7": ("#ffa500", "#000000"), "8": ("#ffc0cb", "#000000"),
    "9": ("#00ffff", "#000000"), "10": ("#800080", "#ffffff"),
    "11": ("#808080", "#ffffff"), "12": ("#a52a2a", "#ffffff"),
    "13": ("#add8e6", "#000000"), "14": ("#800000", "#ffffff"),
    "15": ("#00ff00", "#000000"), "default": ("#d3d3d3", "#000000")
}

# =============================================================================
# LÓGICA DE DIBUJO DEL PDF
# =============================================================================
class CreadorPDF:
    def __init__(self, ruta_excel, modo_color="digital"):
        self.ruta_excel = ruta_excel
        self.modo_color = modo_color
        self.carreras = []
        self.datos_portada = {"fecha": "FECHA NO ENCONTRADA", "reunion": "22"}
        
    def leer_excel(self):
        wb = openpyxl.load_workbook(self.ruta_excel, data_only=True)
        ws = wb.active
        
        carrera_actual = None
        leyendo_tabla = False
        leyendo_actuaciones = False
        
        for row in ws.iter_rows(values_only=True):
            fila = [str(celda).strip() if celda is not None else "" for celda in row]
            fila_texto = " ".join(fila).upper()
            
            # --- PORTADA ---
            m_fecha = re.search(r'(\d{1,2}\s+DE\s+[A-Z]+\s+DE\s+202\d)', fila_texto)
            if m_fecha: self.datos_portada["fecha"] = m_fecha.group(1)
            m_reunion = re.search(r'REUNION Nº\s*(\d+)', fila_texto)
            if m_reunion: self.datos_portada["reunion"] = m_reunion.group(1)
            
            # --- NUEVA CARRERA ---
            val0 = str(fila[0]).strip()
            if "º CARRERA" in val0.upper():
                if carrera_actual: self.carreras.append(carrera_actual)
                
                nro_m = re.search(r'(\d+)º', val0)
                carrera_actual = {
                    "cabecera": {
                        "nro_carrera": nro_m.group(1) if nro_m else "1",
                        "premio": fila[2] if len(fila) > 2 else "PREMIO",
                        "horario": fila[9] if len(fila) > 9 else "",
                        "distancia": "", "condicion": "", "premios_dinero": "", 
                        "apuesta": "", "incremento": "", "incremento_2": ""
                    },
                    "tabla_caballos": [],
                    "actuaciones": ""
                }
                leyendo_tabla = False; leyendo_actuaciones = False
                continue

            if carrera_actual:
                # --- CABECERA EXTENDIDA ---
                if not leyendo_tabla and not leyendo_actuaciones:
                    if "METROS" in val0.upper() or "MTS" in val0.upper():
                        carrera_actual["cabecera"]["distancia"] = val0
                    elif "COMPUTABLE" in val0.upper() or "AL 1º" in val0.upper() or "CAT." in val0.upper() or "PREMIOS:" in val0.upper():
                        carrera_actual["cabecera"]["premios_dinero"] = val0
                    elif "GANADOR" in val0.upper() or "EXACTA" in val0.upper() or "IMPERFECTA" in val0.upper() or "TRIFECTA" in val0.upper():
                        carrera_actual["cabecera"]["incremento_2"] = val0
                    elif val0 and val0.upper() != "4 ULT." and "FECHA" not in val0.upper():
                        carrera_actual["cabecera"]["condicion"] += val0 + " "
                    
                    # Apuestas (Caja Amarilla) - FIX: Agarrar col 8 o 9 y esquivar "CABALLERIZA"
                    val_apuesta = str(fila[8]).strip() if len(fila) > 8 else ""
                    if not val_apuesta and len(fila) > 9: val_apuesta = str(fila[9]).strip()
                    
                    if val_apuesta and "CABALLERIZA" not in val_apuesta.upper() and "CUIDADOR" not in val_apuesta.upper() and "JOCKEY" not in val_apuesta.upper():
                        if "INCREMENTO" in val_apuesta.upper():
                            carrera_actual["cabecera"]["incremento"] = val_apuesta
                        else:
                            carrera_actual["cabecera"]["apuesta"] = val_apuesta

                # --- DETECTAR TABLA ---
                if "CABALLO" in fila_texto and "JOCKEY" in fila_texto:
                    leyendo_tabla = True; continue
                
                # --- LEER CABALLOS ---
                if leyendo_tabla:
                    if not fila[2] and (fila[1] or fila[0]): # Fin de tabla
                        leyendo_tabla = False; leyendo_actuaciones = True
                    elif fila[2]: 
                        carrera_actual["tabla_caballos"].append([
                            fila[0], fila[1], fila[2], fila[3], fila[4], fila[5], 
                            fila[6], fila[8] if len(fila)>8 else "", fila[9] if len(fila)>9 else ""
                        ])
                
                # --- LEER ACTUACIONES ---
                if leyendo_actuaciones:
                    val1 = str(fila[1]).strip()
                    if val0 or val1:
                        if "DEBUTANTE" in val0.upper() or "DEBUTANTE" in val1.upper():
                            mandil = val0 if val0.isdigit() else (val1 if val1.isdigit() else " ")
                            carrera_actual["actuaciones"] += f"{mandil} - Debutante || \n"
                        else:
                            act_izq = val1
                            act_der = str(fila[6]).strip() if len(fila)>6 else ""
                            if act_der: carrera_actual["actuaciones"] += f"{val0} - {act_izq} || {act_der}\n"
                            else: carrera_actual["actuaciones"] += f"{val0} - {act_izq}\n"

        if carrera_actual: self.carreras.append(carrera_actual)

    def generar(self, filepath):
        if not self.carreras:
            raise Exception("No se detectaron carreras en el Excel.")
            
        c = canvas.Canvas(filepath, pagesize=A4)
        W, H = A4; MX = 0.5 * cm; MY = 0.5 * cm
        styles = getSampleStyleSheet()
        
        C_HEAD_BG = C_VERDE_OFICIAL if self.modo_color == "digital" else colors.white
        C_HEAD_TXT = colors.white if self.modo_color == "digital" else colors.black
        
        style_cell_center = ParagraphStyle('CellC', parent=styles['Normal'], fontName='Helvetica', fontSize=6.5, leading=7, alignment=TA_CENTER)
        style_cell_left = ParagraphStyle('CellL', parent=styles['Normal'], fontName='Helvetica', fontSize=6.5, leading=7, alignment=TA_LEFT)
        style_cell_right = ParagraphStyle('CellR', parent=styles['Normal'], fontName='Helvetica', fontSize=6.5, leading=7, alignment=TA_RIGHT)
        style_header = ParagraphStyle('HeadC', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=6.5, leading=7, alignment=TA_CENTER)
        style_cond = ParagraphStyle('Cond', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=9)
        style_legales = ParagraphStyle('Leg', parent=styles['Normal'], fontName='Helvetica', fontSize=7, leading=8, alignment=TA_JUSTIFY)
        
        def _clean_str(txt): return str(txt).replace('"', '').replace("Hs.", "").strip()
        def _parse_money(txt):
            limpio = re.sub(r'[^\d]', '', str(txt))
            return int(limpio) if limpio else 0

        def draw_institutional_header():
            y_curr = H - MY
            h_top_box = 1.5*cm; w_box = W - 2*MX
            c.setStrokeColor(C_VERDE_OFICIAL); c.setLineWidth(2)
            c.rect(MX, y_curr - h_top_box, w_box, h_top_box)
            if LOGO_MAIN.exists(): c.drawImage(str(LOGO_MAIN), MX + 0.3*cm, y_curr - h_top_box + 0.1*cm, width=1.3*cm, height=1.3*cm, mask='auto', preserveAspectRatio=True)
            
            c.setFillColor(colors.black); c.setFont("Helvetica-BoldOblique", 16)
            c.drawCentredString(MX + w_box/2 + 1.0*cm, y_curr - 0.65*cm, "HIPÓDROMO DE TUCUMÁN - PROGRAMA OFICIAL")
            c.setFillColor(C_NARANJA); c.setFont("Helvetica-Bold", 12)
            c.drawCentredString(MX + w_box/2 + 1.0*cm, y_curr - 1.25*cm, f"REUNION Nº {self.datos_portada['reunion']} - {self.datos_portada['fecha']}")
            y_curr -= (h_top_box + 0.15*cm)
            
            st_tit = ParagraphStyle('tit', fontName='Helvetica-Bold', fontSize=6.5, textColor=C_VERDE_OFICIAL, alignment=TA_CENTER)
            st_nom = ParagraphStyle('nom', fontName='Helvetica', fontSize=6.5, textColor=colors.black, alignment=TA_LEFT)
            data_auth = [
                [Paragraph(" ", st_tit), Paragraph("<u>COMISIÓN DE CARRERAS</u>", st_tit), Paragraph("<u>VOCALES</u>", st_tit), Paragraph("<u>DELEGADO HIPODROMO</u>", st_tit)],
                [Paragraph("PRESIDENTE:", st_tit), Paragraph("Dr. Luis Alberto Gamboa", st_nom), Paragraph("Juan Ramon Rouges", st_nom), Paragraph("Estanislao Perez Garcia", st_nom)], 
                [Paragraph("VICE-PRESIDENTE:", st_tit), Paragraph("C.P.N Ernesto José Vidal Sanz", st_nom), Paragraph("Marcos Bruchmann", st_nom), ""], 
                [Paragraph("SECRETARIO:", st_tit), Paragraph("Ignacio Lopez Bustos", st_nom), Paragraph("Santiago Allende", st_nom), ""]
            ]
            t = Table(data_auth, colWidths=[3.2*cm, 6*cm, 4.5*cm, 6.3*cm])
            t.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 2, C_VERDE_OFICIAL), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('TOPPADDING', (0,0), (-1,-1), 1.0), ('BOTTOMPADDING', (0,0), (-1,-1), 1.0)]))
            w_t, h_t = t.wrapOn(c, W, H); t.drawOn(c, MX, y_curr - h_t)
            y_curr -= (h_t + 0.2*cm)
            
            txt_legal = "Admisión y permanencia: Las autoridades del Hipódromo de Tucumán ejercen la facultad de admisión y permanencia en las instalaciones del Hipódromo durante el desarrollo de la reunión hípica. Los profesionales y el público asistente se someten a las disposiciones del Reglamento General de Carreras y a las resoluciones de la Honorable Comisión de Carreras, cuyos fallos son inapelables. Los Boletos no cobrados solo se pagarán, los días de carreras de Tucumán y en el horario en que se desarrolle la reunión y tendrán validez, hasta 2 reuniones siguientes.-"
            p = Paragraph(txt_legal, style_legales); w_leg, h_leg = p.wrap(w_box - 0.4*cm, 5*cm)
            c.setStrokeColor(C_VERDE_OFICIAL); c.setLineWidth(2)
            c.rect(MX, y_curr - h_leg - 0.2*cm, w_box, h_leg + 0.4*cm)
            p.drawOn(c, MX + 0.2*cm, y_curr - h_leg - 0.1*cm)
            y_curr -= (h_leg + 0.6*cm)
            
            h_warn = 1.0*cm
            c.rect(MX, y_curr - h_warn, 9.2*cm, h_warn)
            c.setFillColor(colors.black); c.setFont("Helvetica-BoldOblique", 8.5)
            c.drawCentredString(MX + 4.6*cm, y_curr - h_warn + 0.60*cm, "El juego compulsivo es")
            c.drawCentredString(MX + 4.6*cm, y_curr - h_warn + 0.25*cm, "perjudicial para la salud.")
            
            c.rect(W - MX - 9.5*cm, y_curr - h_warn, 9.5*cm, h_warn)
            c.drawCentredString(W - MX - 4.75*cm, y_curr - h_warn + 0.60*cm, "Los retirados en las apuestas")
            c.drawCentredString(W - MX - 4.75*cm, y_curr - h_warn + 0.25*cm, "combinadas pasan al favorito.")
            
            return (y_curr - h_warn - 0.3*cm)

        def draw_race(carrera, x, y_start, width, idx_carrera):
            cab = carrera['cabecera']; y_curr = y_start
            
            # --- FIX: Tamaños de cabecera ajustados para no chocar ---
            h_head = 1.2*cm
            c.setFillColor(C_HEAD_BG); c.setStrokeColor(colors.black); c.setLineWidth(1)
            c.rect(x, y_curr - h_head, width, h_head, fill=(self.modo_color=="digital"))
            
            c.setFillColor(C_HEAD_TXT); c.setFont("Helvetica-Bold", 13) # Reducido de 17 a 13
            c.drawString(x + 2*mm, y_curr - 7.5*mm, f"{cab['nro_carrera']}º Carrera")
            
            clean_horario = _clean_str(cab['horario']).replace("Hs.", "")
            c.setFont("Helvetica-Bold", 12) # Reducido de 15 a 12
            c.drawRightString(x + width - 2*mm, y_curr - 7.5*mm, f"{clean_horario} Hs.")
            
            clean_premio = _clean_str(cab['premio']) 
            if clean_premio.upper().startswith("PREMIO"): clean_premio = clean_premio[6:].strip()
            c.setFont("Helvetica-Bold", 13) # Reducido de 15 a 13
            c.drawCentredString(x + width/2, y_curr - 7.5*mm, f"PREMIO \"{clean_premio.upper()}\"")
            
            c.setFont("Helvetica-Bold", 8)
            c.drawCentredString(x + width/2, y_curr - 11.5*mm, cab['distancia'])
            y_curr -= (h_head + 2*mm)
            
            clean_cond = cab['condicion'].replace("|", " ").strip()
            if clean_cond:
                p = Paragraph(clean_cond, style_cond)
                w_cond, h_cond = p.wrap(width, 3*cm) 
                p.drawOn(c, x, y_curr - h_cond)
                y_curr -= (h_cond + 4*mm) 
            
            # --- FIX: Premios, Detalle de Apuesta y Caja Amarilla bien alineados ---
            txt_premios = cab['premios_dinero'].strip()
            detalle_ap = cab.get('incremento_2', '').strip()
            txt_ap = cab['apuesta']; txt_inc = cab['incremento']
            
            c.setFillColor(colors.black); c.setFont("Helvetica-Bold", 7.5)
            c.drawString(x, y_curr - 3*mm, txt_premios) 
            if detalle_ap: 
                c.setFont("Helvetica-Bold", 7.5); c.drawString(x, y_curr - 7*mm, detalle_ap)
            
            if txt_ap or txt_inc:
                box_w = 5.2*cm; box_h = 0.9*cm; box_x = x + width - box_w; center_box = box_x + (box_w/2)
                if self.modo_color == "digital":
                    c.setFillColor(colors.lightyellow); c.setStrokeColor(colors.gold)
                    c.rect(box_x, y_curr - box_h + 1*mm, box_w, box_h, fill=1, stroke=1)
                    c.setFillColor(colors.black)
                
                # Fix: Rescatar nombre apuesta si falta
                if not txt_ap and txt_inc: txt_ap = txt_inc.split("$")[0].strip() if "$" in txt_inc else "APUESTA"
                
                c.setFont("Helvetica-BoldOblique", 9)
                c.drawCentredString(center_box, y_curr - 3.5*mm, txt_ap)
                inc_val = _parse_money(txt_inc)
                if inc_val > 0: 
                    c.drawCentredString(center_box, y_curr - 7.0*mm, f"INCREMENTO: $ {inc_val:,.0f}".replace(",","."))
            
            y_curr -= (1.0*cm)
            
            col_ws = [1.3*cm, 0.6*cm, 3.6*cm, 1.0*cm, 2.6*cm, 0.9*cm, 4.0*cm, 3.4*cm, 2.6*cm]
            headers_raw = ['4 Ult.', 'Nº', 'Caballo', 'Pelo', 'Jockey', 'E Kg', 'Padre-Madre', 'Caballeriza', 'Cuidador']
            data = [[Paragraph(h, style_header) for h in headers_raw]]
            for row in carrera['tabla_caballos']:
                nro_raw = str(row[1]); key_mandil = "".join(filter(str.isdigit, nro_raw))
                if not key_mandil: key_mandil = "default"
                bg_hex, fg_hex = MANDILES.get(key_mandil, MANDILES['default'])
                nro_txt = f"<font color='{fg_hex}'><b>{nro_raw}</b></font>" if self.modo_color == "digital" else f"<b>{nro_raw}</b>"
                
                data.append([
                    Paragraph(str(row[0]), style_cell_right), Paragraph(nro_txt, style_cell_center),
                    Paragraph(f"<b>{str(row[2])}</b>", style_cell_left), Paragraph(str(row[3]), style_cell_center),
                    Paragraph(str(row[4]), style_cell_left), Paragraph(str(row[5]), style_cell_center),
                    Paragraph(str(row[6]), style_cell_left), Paragraph(str(row[7]), style_cell_left), Paragraph(str(row[8]), style_cell_left)
                ])
            
            t = Table(data, colWidths=col_ws, rowHeights=[0.48*cm] * len(data)) # Reducido a 0.48 para ahorrar espacio
            ts = [('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                  ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                  ('LEFTPADDING', (0,0), (-1,-1), 1), ('RIGHTPADDING', (0,0), (-1,-1), 1),
                  ('LEFTPADDING', (2,0), (2,-1), 4), ('TOPPADDING', (0,0), (-1,-1), 0.2), ('BOTTOMPADDING', (0,0), (-1,-1), 0.2)]
                  
            for i, row in enumerate(carrera['tabla_caballos']):
                bg_hex, _ = MANDILES.get("".join(filter(str.isdigit, str(row[1]))) or "default", MANDILES['default']) 
                if self.modo_color == "print": bg_hex = "#ffffff"
                ts.append(('BACKGROUND', (1, i+1), (1, i+1), colors.HexColor(bg_hex)))
            
            t.setStyle(TableStyle(ts))
            w_t, h_t = t.wrapOn(c, width, H); t.drawOn(c, x, y_curr - h_t)
            y_curr -= (h_t + 1*mm)
            
            # --- FIX: Centrado vertical matemático de Actuaciones ---
            lines_act = carrera['actuaciones'].split('\n')
            count_lines = sum(1 for l in lines_act if l.strip())
            h_row_exacto = 5.0*mm 
            h_acts = count_lines * h_row_exacto
            
            c.setFillColor(colors.whitesmoke); c.setStrokeColor(colors.lightgrey)
            c.rect(x, y_curr - h_acts, width, h_acts, fill=1, stroke=1)
            
            if any("||" in l for l in lines_act if l.strip()):
                c.setStrokeColor(C_VERDE_OFICIAL); c.setLineWidth(1.5)
                c.line(x + width/2, y_curr, x + width/2, y_curr - h_acts)
            
            curr_y_txt = y_curr
            for l in lines_act:
                if not l.strip(): continue
                curr_y_txt -= h_row_exacto 
                
                c.setStrokeColor(colors.lightgrey); c.setLineWidth(0.5)
                c.line(x + 1*mm, curr_y_txt, x + width - 1*mm, curr_y_txt)
                
                m = re.match(r'^(\d+[a-zA-Z]?)\s*[-\s]+(.*)', l)
                if m:
                    nro_raw, resto = m.groups(); key_mandil = "".join(filter(str.isdigit, nro_raw)) 
                    bg_hex, fg_hex = MANDILES.get(key_mandil or "default", MANDILES['default'])
                    if self.modo_color == "print": bg_hex, fg_hex = "#ffffff", "#000000"
                    
                    c.setFillColor(colors.HexColor(bg_hex)); c.setStrokeColor(colors.black)
                    c.circle(x + 3.5*mm, curr_y_txt + 2.5*mm, 2.1*mm, fill=1, stroke=1)
                    c.setFillColor(colors.HexColor(fg_hex)); c.setFont("Helvetica-Bold", 6.5)
                    c.drawCentredString(x + 3.5*mm, curr_y_txt + 1.4*mm, nro_raw)
                    
                    c.setFillColor(colors.black); c.setFont("Helvetica", 6.5) 
                    parts = resto.split("||"); izq = parts[0].strip(); der = parts[1].strip() if len(parts)>1 else ""
                    c.drawString(x + 8*mm, curr_y_txt + 1.4*mm, izq)
                    if der: c.drawString(x + width/2 + 3*mm, curr_y_txt + 1.4*mm, der)
                else: 
                    c.setFillColor(colors.black); c.setFont("Helvetica", 6.5)
                    c.drawString(x + 2*mm, curr_y_txt + 1.4*mm, l)
                
            return (y_start - (y_curr - h_acts))

        # --- FIX: Cuadro de Incrementos Fiel al Excel ---
        total_inc = 0; apuestas_list = []
        for car in self.carreras:
            cab = car['cabecera']
            txt_ap = str(cab.get('apuesta', '')).strip()
            inc_val = _parse_money(cab.get('incremento', ''))
            
            if inc_val > 0:
                total_inc += inc_val
                nom_ap = txt_ap.split(":")[0].strip() if txt_ap else "APUESTA"
                if nom_ap == "APUESTA" or not nom_ap:
                    m_name = re.match(r'([A-Za-z]+)', str(cab.get('incremento', '')))
                    nom_ap = m_name.group(1).upper() if m_name else "POZO"
                    
                monto_str = f"$ {inc_val:,.0f}.-".replace(",", ".")
                
                rango = 1
                if "CUATERNA" in nom_ap: rango = 4
                elif "TRIPLO" in nom_ap: rango = 3
                elif "QUINTUPLO" in nom_ap: rango = 5
                elif "CADENA" in nom_ap: rango = 6
                elif "DOBLE" in nom_ap: rango = 2
                
                try: nro_start = int(cab['nro_carrera'])
                except: nro_start = 1
                end_nro = nro_start + rango - 1
                
                races = [f"{n}º" for n in range(nro_start, end_nro + 1)]
                c_str = ("; ".join(races[:-1]) + " y " + races[-1] + " carrera.-") if len(races) > 1 else (races[0] + " carrera.-")
                apuestas_list.append(f"{nom_ap} {cab.get('apuesta', '').split(' ')[-1] if '$' in cab.get('apuesta', '') else ''}: {monto_str} {c_str}")

        def draw_footer_area():
            y_curr = MY + 6.0*cm
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 10)
            c.drawCentredString(W/2, y_curr, f"TOTAL INCREMENTOS Y POZOS: $ {total_inc:,.0f}.-".replace(",", "."))
            y_curr -= 0.6*cm
            
            c.setFont("Helvetica-Bold", 9)
            for ap in apuestas_list:
                c.drawCentredString(W/2, y_curr, ap)
                y_curr -= 0.45*cm
            
            y_img = MY
            if LOGO_WSP.exists(): c.drawImage(str(LOGO_WSP), MX, y_img, width=4.5*cm, height=1.3*cm, mask='auto', preserveAspectRatio=True)
            if LOGO_SOC.exists(): c.drawImage(str(LOGO_SOC), W - MX - 4.5*cm, y_img, width=4.5*cm, height=1.3*cm, mask='auto', preserveAspectRatio=True)

        y_cursor = draw_institutional_header()
        if len(self.carreras) > 0: 
            h_used = draw_race(self.carreras[0], MX, y_cursor, W - 2*MX, 1)
        
        draw_footer_area()
        c.showPage()
        
        # --- FIX: Motor de 2 carreras por hoja adaptativo ---
        y_cursor = H - MY
        for i, car in enumerate(self.carreras[1:], start=2):
            # Cálculo exacto de espacio para no rebanar tablas
            filas_cab = len(car['tabla_caballos'])
            lines_act = sum(1 for l in car['actuaciones'].split('\n') if l.strip())
            h_est = 2.5*cm + (filas_cab * 0.48*cm) + (lines_act * 0.5*cm) + 1.0*cm
            
            if y_cursor - h_est < MY: 
                c.showPage()
                y_cursor = H - MY
                
            h_used = draw_race(car, MX, y_cursor, W - 2*MX, i)
            y_cursor -= (h_used + 0.3*cm) 
            
        c.save()

# =============================================================================
# INTERFAZ GRÁFICA
# =============================================================================
class AppGeneradorPDF:
    def __init__(self, root):
        self.root = root
        self.root.title("Exportador PDF Oficial - Hipódromo")
        self.root.geometry("450x300")
        self.root.configure(bg="#f4f4f5")
        self.ruta_excel = None
        
        style = ttk.Style()
        style.configure("TButton", font=("Segoe UI", 11, "bold"), padding=10)
        style.configure("TLabel", font=("Segoe UI", 10), background="#f4f4f5")
        
        ttk.Label(root, text="Generador de PDF", font=("Segoe UI", 16, "bold"), foreground="#248689").pack(pady=20)
        self.lbl_archivo = ttk.Label(root, text="Ningún archivo seleccionado", foreground="gray")
        self.lbl_archivo.pack(pady=5)
        ttk.Button(root, text="📂 Cargar Excel Terminado", command=self.cargar_archivo).pack(pady=10, fill=tk.X, padx=50)
        
        frame_modo = ttk.Frame(root)
        frame_modo.pack(pady=10)
        ttk.Label(frame_modo, text="Modo de Impresión:").pack(side=tk.LEFT, padx=5)
        self.var_modo = tk.StringVar(value="digital")
        ttk.Radiobutton(frame_modo, text="Color (Digital)", variable=self.var_modo, value="digital").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(frame_modo, text="B/N (Imprenta)", variable=self.var_modo, value="print").pack(side=tk.LEFT, padx=5)
        
        self.btn_generar = ttk.Button(root, text="📄 Generar PDF Oficial", command=self.procesar_pdf, state=tk.DISABLED)
        self.btn_generar.pack(pady=20, fill=tk.X, padx=50)

    def cargar_archivo(self):
        ruta = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if ruta:
            self.ruta_excel = ruta
            self.lbl_archivo.config(text=f"Archivo: {os.path.basename(ruta)}", foreground="black")
            self.btn_generar.config(state=tk.NORMAL)

    def procesar_pdf(self):
        if not self.ruta_excel: return
        ruta_salida = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=f"Programa_Oficial_{self.var_modo.get().upper()}.pdf", filetypes=[("Archivo PDF", "*.pdf")])
        if ruta_salida:
            try:
                creador = CreadorPDF(self.ruta_excel, self.var_modo.get())
                creador.leer_excel()
                creador.generar(ruta_salida)
                messagebox.showinfo("¡Éxito!", f"PDF generado impecable en:\n{ruta_salida}")
            except Exception as e:
                traceback.print_exc()
                messagebox.showerror("Error de Generación", f"Hubo un problema:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppGeneradorPDF(root)
    root.mainloop()