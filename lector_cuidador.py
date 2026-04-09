import sqlite3

# Asegurate de poner el nombre exacto de tu archivo de base de datos acá
# (El mismo que usás en tu variable NOMBRE_BD)
NOMBRE_BD = 'carreras.db' 

def buscar_cuidador_dudoso():
    try:
        conn = sqlite3.connect(NOMBRE_BD)
        c = conn.cursor()
        
        # El % sirve como comodín. Busca cualquier texto que contenga "ejero"
        query = "SELECT DISTINCT cuidador FROM caballos WHERE cuidador LIKE '%rey%'"
        c.execute(query)
        resultados = c.fetchall()
        
        if resultados:
            print("¡Bingo! Encontré estos cuidadores en tu base de datos:")
            for fila in resultados:
                # fila[0] porque fetchall devuelve una lista de tuplas
                print(f"-> {fila[0]}") 
        else:
            print("No encontré a nadie que se parezca a Ovejero u Obejero.")
            
    except Exception as e:
        print(f"Error al leer la base de datos: {e}")
    finally:
        if 'conn' in locals():
            conn.close()

if __name__ == "__main__":
    buscar_cuidador_dudoso()