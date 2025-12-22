# migracion.py (CLI para generar data/carreras.db)
import pandas as pd
import re
import os
import sqlite3
from pathlib import Path
import sys

# =========================
#   RUTAS PORTABLES
# =========================
def app_dir() -> Path:
    # Si está empaquetado (PyInstaller) -> carpeta del .exe
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    # En desarrollo -> carpeta del .py
    return Path(__file__).parent

BASE_DIR = app_dir()
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)  # asegura .\data\

# DB dentro de .\data\ al lado del exe/py
NOMBRE_BD = str((DATA_DIR / 'carreras.db').resolve())

# Carpeta base para buscar excels cuando pases sólo el nombre
BUSCADORES = [
    BASE_DIR,                         # mismo folder que el exe/py
    BASE_DIR / "programas",           # subcarpeta programas (opcional)
    BASE_DIR / "resultados",          # subcarpeta resultados (opcional)
    Path.cwd(),                       # por si lo ejecutás desde otra carpeta
]

def _resolver_archivo(nombre: str) -> str | None:
    """Devuelve ruta existente para 'nombre' buscando en BUSCADORES."""
    p = Path(nombre)
    if p.exists():
        return str(p.resolve())
    for root in BUSCADORES:
        cand = (root / nombre).resolve()
        if cand.exists():
            return str(cand)
    return None

# Debug de filas rechazadas (del PROGRAMA)
DEBUG = True
DEBUG_RECHAZOS_FILE = str((BASE_DIR / 'rechazos_programa.csv').resolve())

# =========================================================
#                      UTILIDADES
# =========================================================
def find_col_index_by_keyword(header_row, keyword):
    kw = str(keyword).lower()
    for i, cell in enumerate(header_row):
        if kw in str(cell).lower():
            return i
    return None

def _extraer_tiempo_limpio(df_carrera_actual):
    """
    Extrae el tiempo ganador de un bloque de resultados.
    Soporta:
      - 1'13" / 1' 13"
      - 1'13" 1/5
      - 44"
      - 44" 4/5
    """
    try:
        texto = ' '.join(' '.join(row.astype(str)) for _, row in df_carrera_actual.iterrows())
        m_ti = re.search(r'tiempo\s*[:\-]?\s*(.*)', texto, re.I)
        segmento = m_ti.group(1) if m_ti else texto

        m1 = re.search(r"(\d+)\s*'\s*(\d{1,2})\s*\"\s*(\d\s*/\s*\d)?", segmento)
        if m1:
            mm = int(m1.group(1)); ss = int(m1.group(2))
            frac = m1.group(3)
            frac = re.sub(r"\s+", "", frac) if frac else ""
            frac = (" " + frac) if frac else ""
            return f"{mm}'{ss:02d}\"{frac}"

        m2 = re.search(r"(\d{1,2})\s*\"\s*(\d\s*/\s*\d)?", segmento)
        if m2:
            ss = int(m2.group(1))
            frac = m2.group(2)
            frac = re.sub(r"\s+", "", frac) if frac else ""
            frac = (" " + frac) if frac else ""
            return f"{ss}\"{frac}"

        return "N/D"
    except Exception:
        return "N/D"

def _extraer_estado_pista(df_carrera_actual):
    txt = ' '.join(df_carrera_actual.apply(lambda r: ' '.join(r.astype(str)), axis=1)).upper()
    if 'BARROSA' in txt: return 'PB'
    if 'PESADA'  in txt: return 'PP'
    if 'HUMEDA' in txt or 'HÚMEDA' in txt: return 'PH'
    if 'FANGOSA' in txt: return 'PF'
    if 'NORMAL'  in txt: return 'PN'
    return 'PN'

def _parse_no_corrieron(df_carrera_actual):
    texto = ' '.join(df_carrera_actual.apply(lambda r: ' '.join(r.astype(str)), axis=1))
    if 'CORRIERON TODOS' in texto.upper():
        return []
    salida = []
    m = re.search(r'NO\s+CORRIO(N|ERON)\s*[:=]\s*(.+?)(?:\.|$)', texto, re.I | re.S)
    if not m:
        return []
    partes = re.split(r'\s*,\s*|\s+y\s+', m.group(2))
    for p in partes:
        nm = re.search(r'\((\d+)\)\s*([A-Za-z0-9 .\'\-ÑñÁÉÍÓÚÜáéíóú]+?)(?:\s*\(([^)]+)\))?$', p.strip())
        if nm:
            salida.append({'dorsal': nm.group(1),
                           'nombre': nm.group(2).strip().upper(),
                           'motivo': (nm.group(3) or '').strip()})
    return salida

def _log_rechazo(archivo, hoja, razon, row_series):
    if not DEBUG: return
    fila = ' | '.join([str(x) for x in row_series.values])
    mode = 'a' if os.path.exists(DEBUG_RECHAZOS_FILE) else 'w'
    with open(DEBUG_RECHAZOS_FILE, mode, encoding='utf-8') as f:
        if mode == 'w':
            f.write('archivo,hoja,razon,fila\n')
        f.write(f'"{archivo}","{hoja}","{razon}","{fila.replace(chr(34), chr(39))}"\n')

# =========================================================
#                    CARGA DE RESULTADOS
# =========================================================
def cargar_historial_actuaciones(filepath, sheet_name):
    ruta = _resolver_archivo(filepath)
    if ruta is None:
        print(f"  -> Aviso: no se encontró resultados '{filepath}' en {', '.join(map(str, BUSCADORES))}")
        return None

    try:
        df_full = pd.read_excel(ruta, sheet_name=sheet_name, header=None, dtype=str).fillna('')
    except Exception as e:
        print(f"  -> Error al leer resultados {ruta}: {e}")
        return None

    all_performances = []
    df_str = df_full.astype(str)

    pattern = re.compile(r'\d+[ºª]\s*CARRERA', re.IGNORECASE)
    race_start_indices = [idx for idx, row in df_str.iterrows() if any(pattern.search(str(cell)) for cell in row)]
    if not race_start_indices:
        return None

    estado_pista_general = _extraer_estado_pista(df_str)

    for i, start_index in enumerate(race_start_indices):
        fecha_str = sheet_name
        end_index = race_start_indices[i + 1] if i + 1 < len(race_start_indices) else len(df_full)
        df_carrera = df_str.iloc[start_index:end_index]

        dist_note = next((str(c) for _, r in df_carrera.iterrows() for c in r if re.search(r'distanciad', str(c), re.I)), None)
        inc_note  = next((str(c) for _, r in df_carrera.iterrows() for c in r if re.search(r'tierra|rodó|suelta', str(c), re.I)), None)

        nuevo_puesto_global, ultimo_flag = None, False
        if dist_note and 'ultim' in dist_note.lower():
            ultimo_flag = True
        elif dist_note:
            m_num = re.search(r'al\s*(\d+)', dist_note, re.I)
            if m_num: nuevo_puesto_global = int(m_num.group(1))

        results_start_index = -1
        for local_idx, row in df_carrera.iterrows():
            c1 = str(row.iloc[1]).strip().upper()
            c2 = str(row.iloc[2]).strip()
            if (c1.isdigit() or c1 == 'U') and len(c2) > 0:
                results_start_index = local_idx
                break

        tiempo = _extraer_tiempo_limpio(df_carrera)
        no_corrieron = _parse_no_corrieron(df_carrera)

        if results_start_index == -1:
            for nc in no_corrieron:
                all_performances.append({
                    'Fecha': fecha_str, 'CABALLO': nc['nombre'],
                    'Puesto Original': None, 'Puesto Final': 'NC', 'Jockey': '',
                    'Cuerpos al Ganador': '', 'Ganador': '', 'Segundo': '', 'Margen': '',
                    'Tiempo Ganador': tiempo, 'Pista': estado_pista_general,
                    'Fue Distanciado': False, 'Observacion': nc['motivo'] or 'No corrió'
                })
            continue

        llegadas, pos = [], 1
        for idx in range(results_start_index, end_index):
            row = df_full.iloc[idx]
            if any(re.search(r'divid', str(c), re.I) for c in row): break
            nombre = str(row.iloc[2]).strip()
            if nombre == '':
                if all(c == '' for c in row.astype(str)): break
                else: continue

            llegadas.append({
                'CABALLO': nombre.upper().replace('(*)', '').strip(),
                'Puesto Original': pos,
                'Jockey': str(row.iloc[4]).strip(),
                'Cuerpos al Ganador': str(row.iloc[7]).strip(),
                'TieneAsterisco': '(*)' in ' '.join(row.astype(str)).upper()
            })
            pos += 1

        if not llegadas:
            for nc in no_corrieron:
                all_performances.append({
                    'Fecha': fecha_str, 'CABALLO': nc['nombre'],
                    'Puesto Original': None, 'Puesto Final': 'NC', 'Jockey': '',
                    'Cuerpos al Ganador': '', 'Ganador': '', 'Segundo': '', 'Margen': '',
                    'Tiempo Ganador': tiempo, 'Pista': estado_pista_general,
                    'Fue Distanciado': False, 'Observacion': nc['motivo'] or 'No corrió'
                })
            continue

        if ultimo_flag: nuevo_puesto_global = len(llegadas)

        ganador = llegadas[0]['CABALLO']
        segundo = llegadas[1]['CABALLO'] if len(llegadas) > 1 else ''
        margen  = llegadas[1]['Cuerpos al Ganador'] if len(llegadas) > 1 else ''

        for perf in llegadas:
            puesto_final, fue_dist, obs = perf['Puesto Original'], False, ''
            if perf['TieneAsterisco']:
                if dist_note and nuevo_puesto_global:
                    puesto_final, fue_dist = int(nuevo_puesto_global), True
                elif inc_note:
                    puesto_final, obs = '*', inc_note.strip().lstrip('(*)').strip()

            all_performances.append({
                'Fecha': fecha_str, 'CABALLO': perf['CABALLO'],
                'Puesto Original': perf['Puesto Original'], 'Puesto Final': puesto_final,
                'Jockey': perf['Jockey'], 'Cuerpos al Ganador': perf['Cuerpos al Ganador'],
                'Ganador': ganador, 'Segundo': segundo, 'Margen': margen,
                'Tiempo Ganador': tiempo, 'Pista': estado_pista_general,
                'Fue Distanciado': fue_dist, 'Observacion': obs
            })

        presentes = {p['CABALLO'] for p in llegadas}
        for nc in no_corrieron:
            if nc['nombre'] not in presentes:
                all_performances.append({
                    'Fecha': fecha_str, 'CABALLO': nc['nombre'],
                    'Puesto Original': None, 'Puesto Final': 'NC', 'Jockey': '',
                    'Cuerpos al Ganador': '', 'Ganador': ganador, 'Segundo': segundo, 'Margen': margen,
                    'Tiempo Ganador': tiempo, 'Pista': estado_pista_general,
                    'Fue Distanciado': False, 'Observacion': nc['motivo'] or 'No corrió'
                })

    if not all_performances: return None
    df = pd.DataFrame(all_performances)
    df['Fecha'] = pd.to_datetime(df['Fecha'].str.replace('-', '/'), format='%d/%m/%y', errors='coerce')
    return df

# =========================================================
#                      CARGA PROGRAMA
# =========================================================
def _fila_parece_caballo_suave(row, col_map, archivo, hoja):
    idx_nombre = col_map.get('CABALLO')
    if idx_nombre is None:
        _log_rechazo(archivo, hoja, 'sin índice CABALLO', row)
        return False

    nombre = str(row.iloc[idx_nombre]).strip().upper()
    if not nombre:
        _log_rechazo(archivo, hoja, 'vacío', row); return False
    if re.match(r'^\d{1,2}/\d{1,2}/\d{2,4}', nombre):
        _log_rechazo(archivo, hoja, 'parece fecha', row); return False
    if '"' in nombre:
        _log_rechazo(archivo, hoja, 'título con comillas', row); return False
    if not re.search(r'[A-ZÁÉÍÓÚÜÑ]', nombre):
        _log_rechazo(archivo, hoja, 'sin letras', row); return False

    idx_ekg = col_map.get('E Kg')
    ekg_txt = str(row.iloc[idx_ekg]).strip() if idx_ekg is not None else ''
    ok_ekg = bool(re.match(r'^\D*(\d{1,2})\s+(\d{2,3})\D*$', ekg_txt))

    def _celda_ok(key):
        i = col_map.get(key)
        if i is None: return False
        txt = str(row.iloc[i]).strip()
        return bool(re.search(r'[A-ZÁÉÍÓÚÜÑ]', txt))

    jockey_ok = _celda_ok('Jockey-Descargo')
    cab_ok    = _celda_ok('Caballeriza')
    pm_ok     = _celda_ok('Padre - Madre')

    if ok_ekg or jockey_ok or cab_ok or pm_ok:
        return True

    _log_rechazo(archivo, hoja, 'sin E Kg y sin columnas de apoyo', row)
    return False

def cargar_base_de_datos_caballos(filepath, sheet_name):
    ruta = _resolver_archivo(filepath)
    if ruta is None:
        print(f"  -> Aviso: no se encontró programa '{filepath}' en {', '.join(map(str, BUSCADORES))}")
        return None

    try:
        df_full = pd.read_excel(ruta, sheet_name=sheet_name, header=None, dtype=str).fillna('')
    except Exception as e:
        print(f"  -> Error al leer programa {ruta}: {e}")
        return None

    all_horses = []
    df_str = df_full.astype(str)

    header_indices = [i for i, row in df_str.iterrows() if any('caballo' in str(c).lower() for c in row)]
    if not header_indices:
        return None

    for i, header_index in enumerate(header_indices):
        header_row = df_str.iloc[header_index]
        col_map = {
            'CABALLO':           find_col_index_by_keyword(header_row, 'caballo'),
            'Pelo':              find_col_index_by_keyword(header_row, 'pelo'),
            'Jockey-Descargo':   find_col_index_by_keyword(header_row, 'jockey'),
            'E Kg':              find_col_index_by_keyword(header_row, 'kg'),
            'Padre - Madre':     find_col_index_by_keyword(header_row, 'padre'),
            'Caballeriza':       find_col_index_by_keyword(header_row, 'caballeriza'),
            'Cuidador':          find_col_index_by_keyword(header_row, 'cuidador'),
            '4 Ult.':            find_col_index_by_keyword(header_row, 'ult.')
        }
        if col_map.get('CABALLO') is None:
            continue

        end_of_block = header_indices[i + 1] if i + 1 < len(header_indices) else len(df_full)

        current_index = header_index + 1
        while current_index < end_of_block:
            row = df_full.iloc[current_index]

            if all(str(x) == '' for x in row.values):
                current_index += 1
                continue

            if _fila_parece_caballo_suave(row, col_map, filepath, sheet_name):
                try:
                    horse = {key: (row.iloc[idx] if idx is not None else '')
                             for key, idx in col_map.items()}
                    horse['CABALLO'] = str(horse['CABALLO']).strip().upper()
                    if horse['CABALLO'] and horse['CABALLO'] != 'NAN':
                        all_horses.append(horse)
                except IndexError:
                    pass

            current_index += 1

    if not all_horses:
        return None
    return pd.DataFrame(all_horses)

# =========================================================
#                       ESQUEMA DE BD
# =========================================================
def crear_base_de_datos():
    try:
        if os.path.exists(NOMBRE_BD):
            os.remove(NOMBRE_BD)
    except Exception:
        pass

    if DEBUG:
        try:
            if os.path.exists(DEBUG_RECHAZOS_FILE):
                os.remove(DEBUG_RECHAZOS_FILE)
        except Exception:
            pass

    print(f"Creando la nueva base de datos '{NOMBRE_BD}'...")
    conn = sqlite3.connect(NOMBRE_BD)
    c = conn.cursor()

    c.execute('''
        CREATE TABLE caballos (
            nombre TEXT PRIMARY KEY,
            padre_madre TEXT,
            pelo TEXT,
            ultima_edad TEXT,
            ultimo_peso TEXT,
            ultimo_jockey TEXT,
            caballeriza TEXT,
            cuidador TEXT,
            ultima_actuacion_externa TEXT,
            snapshot_programa_fecha DATE
        )
    ''')

    c.execute('''
        CREATE TABLE actuaciones (
            id INTEGER PRIMARY KEY,
            fecha DATE,
            nombre_caballo TEXT,
            puesto_original INTEGER,
            puesto_final TEXT,
            jockey TEXT,
            cuerpos TEXT,
            ganador TEXT,
            segundo TEXT,
            margen TEXT,
            tiempo_ganador TEXT,
            pista TEXT,
            fue_distanciado BOOLEAN,
            observacion TEXT
        )
    ''')

    conn.commit()
    conn.close()
    print(f"Base de datos '{NOMBRE_BD}' creada.")

# =========================================================
#                   EJECUCIÓN PRINCIPAL
# =========================================================
if __name__ == "__main__":
    # <<< EDITÁ ESTO CON TUS ARCHIVOS >>>
    LISTA_PROGRAMAS = {
        '16-02-25': '16 DE FEBRERO DE 2025.xlsx', '23-02-25': '23 DE FEBRERO DE 2025.xlsx',
        '16-03-25': '16 DE MARZO DE 2025.xlsx',   '30-03-25': '30 DE MARZO DE 2025.xlsx',
        '13-04-25': '13 DE ABRIL DE 2025.xlsx',   '26-04-25': '26 DE ABRIL DE 2025.xlsx',
        '11-05-25': '11 DE MAYO DE 2025.xlsx',    '25-05-25': '25 DE MAYO DE 2025.xlsx',
        '08-06-25': '08 DE JUNIO DE 2025.xlsx',   '22-06-25': '22 DE JUNIO DE 2025.xlsx',
        '13-07-25': '13 DE JULIO DE 2025.xlsx',   '27-07-25': '27 DE JULIO DE 2025.xlsx',
        '10-08-25': '10 DE AGOSTO DE 2025.xlsx',  '24-08-25': '24 DE AGOSTO DE 2025.xlsx',
        '07-09-25': '07 DE SEPTIEMBRE DE 2025.xlsx','24-09-25': '24 DE SEPTIEMBRE DE 2025.xlsx',
        '05-10-25': '05 DE OCTUBRE DE 2025.xlsx', '18-10-25': '18 DE OCTUBRE DE 2025.xlsx', '09-11-25': '09 DE NOVIEMBRE DE 2025.xlsx',
        '30-11-25': '30 DE NOVIEMBRE DE 2025.xlsx',
    }

    LISTA_RESULTADOS = {
        '16-02-25': 'Resultados 16-02-25.xlsx', '23-02-25': 'Resultados 23-02-25.xlsx',
        '16-03-25': 'Resultados 16-03-25.xlsx', '30-03-25': 'Resultados 30-03-25.xlsx',
        '13-04-25': 'Resultados 13-04-25.xlsx', '26-04-25': 'Resultados 26-04-25.xlsx',
        '11-05-25': 'Resultados 11-05-25.xlsx', '25-05-25': 'Resultados 25-05-25.xlsx',
        '08-06-25': 'Resultados 08-06-25.xlsx', '22-06-25': 'Resultados 22-06-25.xlsx',
        '13-07-25': 'Resultados 13-07-25.xlsx', '27-07-25': 'Resultados 27-07-25.xlsx',
        '10-08-25': 'Resultados 10-08-25.xlsx', '24-08-25': 'Resultados 24-08-25.xlsx',
        '07-09-25': 'Resultados 07-09-25.xlsx', '24-09-25': 'Resultados 24-09-25.xlsx',
        '05-10-25': 'Resultados 05-10-25.xlsx', '18-10-25': 'Resultados 18-10-25.xlsx',
        '09-11-25': 'Resultados 09-11-25.xlsx', '30-11-25': 'Resultados 30-11-25.xlsx',
    }
    # >>> FIN EDITABLE <<<

    crear_base_de_datos()
    conn = sqlite3.connect(NOMBRE_BD)

    # --------- CABALLOS ---------
    print("\n--- Iniciando migración de CABALLOS ---")
    for fecha_hoja, archivo in LISTA_PROGRAMAS.items():
        ruta = _resolver_archivo(archivo)
        if ruta is None:
            print(f"Programa no encontrado: {archivo}")
            continue

        print(f"Programa: {Path(ruta).name}...")
        df_cab = cargar_base_de_datos_caballos(archivo, fecha_hoja)
        if df_cab is None:
            continue

        for _, row in df_cab.iterrows():
            nombre = str(row.get('CABALLO', '')).strip().upper()
            e_kg_str = str(row.get('E Kg', '')).strip()
            edad, peso = '', ''
            m = re.match(r'^\D*(\d{1,2})\s+(\d{2,3})\D*$', e_kg_str)
            if m:
                edad, peso = m.group(1), m.group(2)

            actu_ext = row.get('4 Ult.', '') or ''
            if 'debuta' in actu_ext.lower():
                actu_ext = ''

            conn.execute('INSERT OR IGNORE INTO caballos (nombre) VALUES (?)', (nombre,))
            conn.execute('''
                UPDATE caballos
                SET padre_madre=?, pelo=?, ultima_edad=?, ultimo_peso=?,
                    ultimo_jockey=?, caballeriza=?, cuidador=?, ultima_actuacion_externa=?,
                    snapshot_programa_fecha=?
                WHERE nombre=?
            ''', (row.get('Padre - Madre', ''), row.get('Pelo', ''), edad, peso,
                  row.get('Jockey-Descargo', ''), row.get('Caballeriza', ''),
                  row.get('Cuidador', ''), actu_ext, fecha_hoja, nombre))
    conn.commit()
    print("--- CABALLOS OK ---")

    # --------- ACTUACIONES ---------
    print("\n--- Iniciando migración de ACTUACIONES ---")
    conn.execute('DELETE FROM actuaciones')
    for fecha_hoja, archivo in LISTA_RESULTADOS.items():
        ruta = _resolver_archivo(archivo)
        if ruta is None:
            print(f"Resultados no encontrados: {archivo}")
            continue

        print(f"Resultados: {Path(ruta).name}...")
        df_act = cargar_historial_actuaciones(archivo, fecha_hoja)
        if df_act is None:
            continue
        colmap = {
            'Fecha': 'fecha', 'CABALLO': 'nombre_caballo',
            'Puesto Original': 'puesto_original', 'Puesto Final': 'puesto_final',
            'Jockey': 'jockey', 'Cuerpos al Ganador': 'cuerpos',
            'Ganador': 'ganador', 'Segundo': 'segundo', 'Margen': 'margen',
            'Tiempo Ganador': 'tiempo_ganador', 'Pista': 'pista',
            'Fue Distanciado': 'fue_distanciado', 'Observacion': 'observacion'
        }
        df_act.rename(columns=colmap).to_sql('actuaciones', conn, if_exists='append', index=False)
    conn.commit()
    print("--- ACTUACIONES OK ---")

    # --------- RESUMEN ---------
    cur = conn.cursor()
    tot_cab = cur.execute('SELECT COUNT(*) FROM caballos').fetchone()[0]
    tot_act = cur.execute('SELECT COUNT(*) FROM actuaciones').fetchone()[0]
    print("\n--- MIGRACIÓN COMPLETA ---")
    print(f"DB: {NOMBRE_BD}")
    print(f"Total de caballos únicos: {tot_cab}")
    print(f"Total de actuaciones: {tot_act}")
    conn.close()
