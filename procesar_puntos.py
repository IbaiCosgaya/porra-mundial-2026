import os
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime

DOCUMENTO_ID = "1G4cyvvzsPw7p4yDgXiuPorsP9CDnqPuSXC-wR_H4IIY"

def conectar_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    if "GOOGLE_CREDENTIALS" in os.environ:
        creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    else:
        creds = ServiceAccountCredentials.from_json_keyfile_name("credenciales.json", scope)
        
    cliente = gspread.authorize(creds)
    return cliente.open_by_key(DOCUMENTO_ID).get_worksheet(0)

def calcular_resultado_1X2(goles_l, goles_v):
    if goles_l > goles_v: return '1'
    elif goles_l < goles_v: return '2'
    else: return 'X'

def normalizar_texto(texto):
    if not texto: return ""
    t = str(texto).strip().upper()
    for k, v in {"Á":"A","É":"E","Í":"I","Ó":"O","Ú":"U"}.items():
        t = t.replace(k, v)
    return t

def celda_a_marcador(valor):
    if not valor: return None
    v_str = str(valor).strip()
    
    if '-' in v_str and not any(c.isalpha() for c in v_str):
        return v_str

    formatos_fecha = ["%d-%b", "%d/%m", "%d-%m", "%b-%d"]
    for formato in formatos_fecha:
        try:
            dt = datetime.strptime(v_str, formato)
            return f"{dt.day}-{dt.month}"
        except ValueError:
            continue
            
    return v_str

def celda_a_regla_partido(valor):
    if not valor: return ""
    v_str = str(valor).strip()
    formatos_fecha = ["%d-%b", "%d/%m", "%d-%m"]
    for formato in formatos_fecha:
        try:
            dt = datetime.strptime(v_str, formato)
            return f"{dt.day}/{dt.month}"
        except ValueError:
            continue
    return v_str

try:
    print("🔄 Conectando con Google Sheets...")
    hoja_excel = conectar_google_sheets()
    todas_las_filas = hoja_excel.get_all_values()
    
    # La fila 0 contiene los títulos o nombres de las columnas
    fila_nombres = todas_las_filas[0]

    participantes = {}
    # Mapeamos los participantes basándonos en las columnas a partir de la F (índice 5)
    for idx, nombre in enumerate(fila_nombres):
        nombre_limpio = nombre.strip()
        if nombre_limpio and nombre_limpio not in ("PORRA MUNDIAL 2026", "Resultado real", "FECHA", "HORA", "PARTIDO", "Exacto/1X2", ""):
            # Guardamos la columna exacta de cada participante (ej: Asier está en la columna F -> índice 5)
            participantes[nombre_limpio] = {"col_apuesta": idx, "puntos_totales": 0}

    print("✅ Participantes mapeados con sus columnas:", list(participantes.keys()))

    # Recorremos todas las filas de datos (empezando desde el índice 1, que es la fila 2 de Excel)
    for idx, fila in enumerate(todas_las_filas[1:], start=2):
        if len(fila) < 6: continue

        # Extraemos las columnas base (C, D, E) correspondientes a los índices 2, 3 y 4
        pregunta_partido = fila[2].strip()
        regla_raw = fila[3].strip()
        col_resultado_real = fila[4].strip()

        # Si no hay resultado real definitivo puesto en la columna E, saltamos la fila
        if not col_resultado_real:
            continue

        regla_puntos = celda_a_regla_partido(regla_raw)
        real_norm = normalizar_texto(col_resultado_real)

        # 🎯 DETECTAR TIPO DE FILA
        # Si la regla contiene "5/2" o la celda real tiene un guion de marcador (ej: "2-0")
        if "5/2" in regla_raw or "5/2" in regla_puntos or ("-" in col_resultado_real and not any(c.isalpha() for c in col_resultado_real)):
            tipo = "PARTIDO"
        elif "/" in regla_raw and ("S" in regla_raw.upper() or "N" in regla_raw.upper()):
            tipo = "SI_NO"
        else:
            tipo = "PREGUNTA_ABIERTA"

        # --- ⚽ PROCESAR PARTIDOS ⚽ ---
        if tipo == "PARTIDO":
            marcador_real = celda_a_marcador(col_resultado_real)
            if not marcador_real or '-' not in marcador_real:
                continue
            try:
                re_l, re_v = map(int, marcador_real.split("-"))
                signo_real = calcular_resultado_1X2(re_l, re_v)
            except ValueError:
                continue

            for nombre, cols in participantes.items():
                # Cada participante tiene su apuesta completa (ej: "3-1") en SU PROPIA COLUMNA
                val_apuesta = fila[cols["col_apuesta"]].strip()
                if not val_apuesta:
                    continue

                marcador_apuesta = celda_a_marcador(val_apuesta)
                if not marcador_apuesta or '-' not in marcador_apuesta:
                    continue

                try:
                    ap_l, ap_v = map(int, marcador_apuesta.split("-"))
                    signo_apuesta = calcular_resultado_1X2(ap_l, ap_v)

                    if ap_l == re_l and ap_v == re_v:
                        participantes[nombre]["puntos_totales"] += 5  # Pleno al marcador exacto
                    elif signo_apuesta == signo_real:
                        participantes[nombre]["puntos_totales"] += 2  # Acierto de 1X2
                except ValueError:
                    pass

        # --- 🔲 PROCESAR SÍ / NO 🔲 ---
        elif tipo == "SI_NO":
            letra_real = "S" if real_norm in ("SI", "S") else "N"
            puntos_por_acierto = 0
            # Parseamos reglas del tipo S10 / N3
            for p in regla_raw.upper().replace(" ", "").split("/"):
                if p.startswith(letra_real):
                    nums = ''.join(filter(str.isdigit, p))
                    puntos_por_acierto = int(nums) if nums else 0

            for nombre, cols in participantes.items():
                apuesta_user = normalizar_texto(fila[cols["col_apuesta"]])
                letra_apuesta = "S" if apuesta_user in ("SI", "S") else "N"
                if letra_apuesta == letra_real:
                    participantes[nombre]["puntos_totales"] += puntos_por_acierto

        # --- 📝 PROCESAR PREGUNTAS ABIERTAS 📝 ---
        elif tipo == "PREGUNTA_ABIERTA":
            try:
                nums = ''.join(filter(str.isdigit, regla_raw))
                puntos_regla = int(nums) if nums else 10
            except:
                puntos_regla = 10

            for nombre, cols in participantes.items():
                apuesta_user = normalizar_texto(fila[cols["col_apuesta"]])
                if apuesta_user == real_norm:
                    participantes[nombre]["puntos_totales"] += puntos_regla

    # Generación de Ranking Final Ordenado
    ranking = [{"Nombre": k, "Puntos": v["puntos_totales"]} for k, v in participantes.items()]
    df_ranking = pd.DataFrame(ranking).sort_values(by="Puntos", ascending=False).reset_index(drop=True)
    df_ranking.index += 1

    print("\n🏆 CLASIFICACIÓN CORREGIDA 🏆")
    print(df_ranking)

    # Guardar cambios
    datos_json = df_ranking.to_json(orient="records", force_ascii=False, indent=4)
    with open("clasificacion.json", "w", encoding="utf-8") as f:
        f.write(datos_json)
    print("\n💾 'clasificacion.json' actualizado con éxito!")

except Exception as e:
    import traceback
    traceback.print_exc()
    print(f"❌ Error: {e}")
