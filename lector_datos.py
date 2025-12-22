import pandas as pd
import re
import os
import difflib
import sqlite3
from typing import Optional

# --- CONFIGURACIÓN ---
NOMBRE_BD = 'carreras.db'
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)

# ---------------------- Utilidades ----------------------

def _to_bool(x) -> bool:
    sx = str(x).strip().lower()
    if sx.isdigit():
        return bool(int(sx))
    return sx in ('true', 't', 'si', 'sí', '1')

def _puesto_es_numero(x) -> Optional[int]:
    """
    Devuelve int si el puesto es numérico (1,2,3...), sino None.
    Soporta strings tipo '1', '2.0', etc.
    """
    try:
        s = str(x).strip()
        if s == '' or not re.match(r'^\d+(\.0+)?$', s):
            return None
        return int(float(s))
    except:
        return None

def convertir_a_fraccion(valor):
    """
    Convierte '1.5' -> '1 1/2', '0.25' -> '1/4', etc. Deja 'cp' y no numéricos tal cual.
    """
    try:
        valor_str = str(valor).strip().replace('cp', '').strip()
        if '.' in valor_str:
            partes = valor_str.split('.')
            entero = partes[0]
            decimal = float('0.' + partes[1])
            if decimal == 0.5:
                fraccion = '1/2'
            elif decimal == 0.25:
                fraccion = '1/4'
            elif decimal == 0.75:
                fraccion = '3/4'
            else:
                return valor
            return fraccion if entero == '0' else f"{entero} {fraccion}"
        return valor
    except:
        return valor

def marcador_4ult(series_puestos_ordenadas_desc):
    """
    Recibe una Serie de Puesto Final ordenada por Fecha DESC (nuevo->viejo).
    - Toma las últimas 4 (de más nueva a más vieja), luego invierte para que la última quede a la DERECHA.
    - Puestos >=10 se muestran como '0'.
    - No numéricos (NC, U, *) se conservan.
    """
    vals = []
    for x in series_puestos_ordenadas_desc.head(4):
        sx = str(x).strip()
        n = _puesto_es_numero(sx)
        if n is None:
            vals.append(sx if sx else '-')
        else:
            vals.append('0' if n >= 10 else str(n))
    return "-".join(reversed(vals)) if vals else "Debuta"

def _es_debutante_hoy(row_caballo, actuaciones_orden_desc: pd.DataFrame) -> bool:
    """
    Regla: si en Programa figura sin '4 Ult.' real (guardamos ultima_actuacion_externa = '')
    y aún no tiene ninguna actuación numérica (todas NC o sin registros) => sigue Debutante.
    En el primer día con puesto numérico deja de ser Debutante.
    """
    sin_tabulada_programa = str(row_caballo.get('ultima_actuacion_externa', '')).strip() == ''

    if actuaciones_orden_desc.empty:
        return True if sin_tabulada_programa else False

    tiene_actuacion_numerica = actuaciones_orden_desc['Puesto Final'].apply(_puesto_es_numero).notna().any()
    if sin_tabulada_programa and not tiene_actuacion_numerica:
        return True
    return False

# ---- helpers de presentación ----

def _title_caballo(nombre: str) -> str:
    """Camel Case básico para nombres de caballos/carreras."""
    s = str(nombre or '').strip()
    if not s:
        return s
    s = s.title()
    # preservar D' / O' / L' etc (capitalizo la letra siguiente)
    s = re.sub(r"\b([DOL])'([A-Za-z])", lambda m: m.group(1).upper() + "'" + m.group(2).upper(), s)
    return s

def _jockey_corto_desde_full(full: str) -> str:
    """
    Formatea 'Vizcarra Jose A.' -> 'J. Vizcarra'
    y 'Vai Angel' -> 'A. Vai'. Detecta orden más probable.
    """
    full = str(full or '').strip()
    if not full:
        return ''
    p = [t for t in full.split() if t]
    if len(p) >= 2:
        # Heurística de orden: si el segundo token no termina con '.', lo tomo como NOMBRE y el último como APELLIDO.
        # (En tus excels suele venir 'Vizcarra Jose A.' o 'Vai Angel')
        nombre = p[1] if not p[1].endswith('.') else p[0]
        apellido = p[0] if not p[1].endswith('.') else (p[-1] if len(p) >= 2 else p[0])
        return f"{nombre[0].upper()}. {_title_caballo(apellido)}"
    return _title_caballo(full)

def _jockey_corto(j_from_actuacion: str, j_from_programa: str) -> str:
    """
    Si en resultados viene como iniciales tipo 'V. A.' usamos el del PROGRAMA para reconstruir.
    Si viene completo, formateamos directo.
    """
    ja = str(j_from_actuacion or '').strip()
    jf = str(j_from_programa or '').strip()

    if re.fullmatch(r"[A-Za-z]\.\s*[A-Za-z]\.?", ja):
        # sólo iniciales -> reconstruyo con el del programa
        return _jockey_corto_desde_full(jf) if jf else ja

    # ya viene completo -> formateo
    return _jockey_corto_desde_full(ja if ja else jf)

# ---------------------- Carga de datos ----------------------

def conectar_y_cargar_datos():
    if not os.path.exists(NOMBRE_BD):
        print(f"Error: No se encuentra la base de datos '{NOMBRE_BD}'. Ejecuta primero 'migracion.py'.")
        return None, None

    print(f"Conectando a la base de datos '{NOMBRE_BD}'...")
    conn = sqlite3.connect(NOMBRE_BD)
    df_caballos = pd.read_sql_query("SELECT * FROM caballos", conn)
    df_actuaciones = pd.read_sql_query("SELECT * FROM actuaciones", conn)
    conn.close()
    print("¡Datos cargados en memoria listos para consultar!")

    # Normalizo nombres de columnas
    df_caballos = df_caballos.rename(columns={
        'nombre': 'Caballo',
        'ultima_edad': 'Edad',
        'ultimo_peso': 'Peso',
        'ultimo_jockey': 'Jockey-Descargo',
        'padre_madre': 'Padre - Madre',
        'caballeriza': 'Caballeriza',
        'cuidador': 'Cuidador',
        'pelo': 'Pelo'
    })
    df_actuaciones = df_actuaciones.rename(columns={
        'nombre_caballo': 'Caballo',
        'puesto_original': 'Puesto Original',
        'puesto_final': 'Puesto Final',
        'jockey': 'Jockey',
        'cuerpos': 'Cuerpos al Ganador',
        'ganador': 'Ganador',
        'segundo': 'Segundo',
        'margen': 'Margen',
        'tiempo_ganador': 'Tiempo Ganador',
        'pista': 'Pista',
        'fue_distanciado': 'Fue Distanciado',
        'fecha': 'Fecha',
        'observacion': 'Observacion'
    })

    # Tipos / normalizaciones básicas
    df_actuaciones['Fecha'] = pd.to_datetime(df_actuaciones['Fecha'], errors='coerce')
    if 'Fue Distanciado' in df_actuaciones.columns:
        df_actuaciones['Fue Distanciado'] = df_actuaciones['Fue Distanciado'].apply(_to_bool)

    for col in ['Caballo', 'Ganador', 'Segundo']:
        if col in df_actuaciones.columns:
            df_actuaciones[col] = df_actuaciones[col].astype(str).str.upper().str.strip()

    if 'Caballo' in df_caballos.columns:
        df_caballos['Caballo'] = df_caballos['Caballo'].astype(str).str.upper().str.strip()

    return df_caballos, df_actuaciones

# ---------------------- Reporte de un caballo ----------------------

def generar_texto_actuaciones(nombre_caballo, db_caballos, db_actuaciones):
    nombre_caballo_upper = nombre_caballo.strip().upper()
    try:
        info_caballo_raw = db_caballos[db_caballos['Caballo'] == nombre_caballo_upper].iloc[0]
    except IndexError:
        # Sugerencia si no lo encuentra
        lista_nombres_caballos = db_caballos['Caballo'].tolist()
        sugerencias = difflib.get_close_matches(nombre_caballo_upper, lista_nombres_caballos, n=1, cutoff=0.7)
        if sugerencias:
            print(f"\nNo se encontró a '{nombre_caballo}'. ¿Quisiste decir '{sugerencias[0]}'?")
        else:
            print(f"\nNo se encontró a '{nombre_caballo}'.")
        return

    # Actuaciones del caballo ordenadas por fecha DESC
    actuaciones_caballo = db_actuaciones[db_actuaciones['Caballo'] == nombre_caballo_upper] \
        .sort_values(by='Fecha', ascending=False)

    # --- "4 Ult." ---
    ultimas_4 = marcador_4ult(actuaciones_caballo['Puesto Final'])
    es_debutante = _es_debutante_hoy(info_caballo_raw, actuaciones_caballo)
    a_ult_texto = "Debuta" if es_debutante else ultimas_4

    # Armar visual del bloque PROGRAMA
    info_caballo_display = info_caballo_raw.copy()
    try:
        edad = int(float(info_caballo_display.get('Edad', '')))
        peso = int(float(info_caballo_display.get('Peso', '')))
        info_caballo_display['E Kg'] = f"{edad} {peso}"
    except (ValueError, TypeError):
        info_caballo_display['E Kg'] = f"{info_caballo_display.get('Edad', '')} {info_caballo_display.get('Peso', '')}"

    info_caballo_display['4 Ult.'] = a_ult_texto
    info_caballo_display['Nº'] = info_caballo_raw.get('Nº', '')
    columnas_ordenadas = ['4 Ult.', 'Nº', 'Caballo', 'Pelo', 'Jockey-Descargo', 'E Kg', 'Padre - Madre', 'Caballeriza', 'Cuidador']

    print("\n--- PROGRAMA ---")
    print(info_caballo_display[columnas_ordenadas].to_string())

    # --- ÚLTIMAS ACTUACIONES ---
    print("\n--- ÚLTIMAS ACTUACIONES ---")
    if actuaciones_caballo.empty:
        print("Sin actuaciones registradas." + (" (Debutante)" if es_debutante else ""))
        return

    jockey_programa_full = str(info_caballo_raw.get('Jockey-Descargo', '')).strip()

    # Mostramos las 2 más recientes
    for _, a in actuaciones_caballo.head(2).iterrows():
        fecha_txt = a['Fecha'].strftime('%d/%m/%y') if pd.notna(a['Fecha']) else ''

        # Si NC
        if str(a['Puesto Final']).strip().upper() == 'NC':
            obs = str(a.get('Observacion', '')).strip()
            obs_txt = f" ({obs})" if obs else ""
            print(f"{fecha_txt} - No Corrió{obs_txt}.")
            continue

        # Jockey formateado correctamente
        jockey_corto = _jockey_corto(str(a.get('Jockey', '')).strip(), jockey_programa_full)

        puesto_orig = _puesto_es_numero(a.get('Puesto Original'))
        puesto_original_str = f"{puesto_orig}º" if puesto_orig else str(a.get('Puesto Original', '')).strip()

        fue_dist = bool(a.get('Fue Distanciado', False))
        dist_txt = " - Distanciado" if fue_dist else ""

        if puesto_orig == 1:
            # Ganador
            margen = convertir_a_fraccion(a.get('Margen', ''))
            segundo_upper = str(a.get('Segundo', '')).strip().upper()
            segundo_txt = _title_caballo(segundo_upper)
            
            # Busco si el segundo fue distanciado (en TODA la tabla)
            seg_info = db_actuaciones[
                (db_actuaciones['Caballo'] == segundo_upper) &
                (db_actuaciones['Fecha'] == a['Fecha'])
            ]
            if not seg_info.empty and bool(seg_info.iloc[0].get('Fue Distanciado', False)):
                segundo_txt += " (Dist.)"


            print(f"{fecha_txt} - {jockey_corto} - 1º gan x {margen} cp a {segundo_txt} - {a.get('Tiempo Ganador','')} - {a.get('Pista','')}{dist_txt}")
        else:
            # No ganador
            cuerpos = convertir_a_fraccion(a.get('Cuerpos al Ganador', ''))
            ganador_txt = _title_caballo(str(a.get('Ganador','')).strip())
            print(f"{fecha_txt} - {jockey_corto} - {puesto_original_str} a {cuerpos} cp de {ganador_txt} - {a.get('Tiempo Ganador','')} - {a.get('Pista','')}{dist_txt}")

# ---------------------- CLI ----------------------

def _listar_caballos(db_caballos):
    print("\n--- LISTA DE CABALLOS CARGADOS ---")
    nombres_ordenados = sorted(db_caballos['Caballo'].unique())
    columnas = 4
    for i in range(0, len(nombres_ordenados), columnas):
        print(" | ".join(f"{name:<25}" for name in nombres_ordenados[i:i+columnas]))
    print("----------------------------------")

def _listar_debutantes(db_caballos, db_actuaciones):
    """
    Lista rápida de los que hoy siguen debutantes (según reglas).
    """
    print("\n--- POSIBLES DEBUTANTES ---")
    count = 0
    # Preindexamos actuaciones por caballo para que sea rápido
    acts_por_caballo = {
        c: df.sort_values('Fecha', ascending=False)
        for c, df in db_actuaciones.groupby('Caballo')
    }
    for _, row in db_caballos.iterrows():   # <-- acá antes decía df_caballos
        caballo = row['Caballo']
        acts = acts_por_caballo.get(caballo, pd.DataFrame(columns=db_actuaciones.columns))
        if _es_debutante_hoy(row, acts):
            print(f"- {caballo}")
            count += 1
    if count == 0:
        print("No hay debutantes detectados con las reglas actuales.")


# --- Ejecución Principal ---
if __name__ == "__main__":
    db_caballos, db_actuaciones = conectar_y_cargar_datos()
    if db_caballos is not None and db_actuaciones is not None:
        print("\n¡Bienvenido al programa de carreras v3.1!")
        print("Comandos: '!lista' para ver todos, '!debutantes' para ver posibles debutantes, 'salir' para terminar.")
        while True:
            nombre = input("\n> Introduce el nombre del caballo: ").strip()
            if nombre.lower() == 'salir':
                break
            if nombre.lower() == '!lista':
                _listar_caballos(db_caballos)
                continue
            if nombre.lower() == '!debutantes':
                _listar_debutantes(db_caballos, db_actuaciones)
                continue
            if not nombre:
                continue
            generar_texto_actuaciones(nombre, db_caballos, db_actuaciones)
