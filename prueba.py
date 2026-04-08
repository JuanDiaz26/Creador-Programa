import sqlite3
conn = sqlite3.connect('data/carreras.db')
cursor = conn.cursor()
resultado = cursor.execute("SELECT ultima_actuacion_externa, texto_actuaciones_externas FROM caballos WHERE nombre='DIVINO TESORO'").fetchone()
conn.close()

if resultado:
    print(f"Siglas en DB (4 Ult.): {resultado[0]}")
    print(f"Texto Largo en DB:\n{resultado[1]}")
else:
    print("Caballo no encontrado en DB.")