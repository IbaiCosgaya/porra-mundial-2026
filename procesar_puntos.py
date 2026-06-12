import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime

DOCUMENTO_ID = "1G4cyvvzsPw7p4yDgXiuPorsP9CDnqPuSXC-wR_H4IIY"

import os
import json

# ... (resto de funciones de cálculo se quedan igual)

def conectar_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    
    # Intentar leer desde las variables secretas de GitHub Actions
    if "GOOGLE_CREDENTIALS" in os.environ:
        creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    else:
        # Por si lo sigues ejecutando en local en tu PC
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
    """
    Convierte el texto de la celda al marcador 'local-visitante'.
    Si Excel lo transformó en texto de fecha (ej: '3-ene', '03/01'), 
    lo parsea y extrae el día (local) y el mes (visitante).
    """
    if not valor: return None
    v_str = str(valor).strip()
    
    # Caso 1: Ya es un formato limpio '3-1' o '2-0'
    if '-' in v_str and not any(c.isalpha() for c in v_str):
        return v_str

    # Caso 2: Intentar detectar si Google Sheets lo pasó como texto de fecha
    formatos_fecha = ["%d-%b", "%d/%m", "%d-%m", "%b-%d"]
    for formato in formatos_fecha:
        try:
            dt = datetime.strptime(v_str, formato)
            # Reemplazamos el año por el actual si hace falta, pero nos interesan día y mes
            return f"{dt.day}-{dt.month}"
        except ValueError:
            continue
            
    return v_str

def celda_a_regla_partido(valor):
    """
    Convierte la columna D al string de regla correcto.
    Si viene como fecha (ej: '5-feb' o '5/2' convertido), extrae día/mes.
    """
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
    
    fila_nombres = todas_las_filas[0]

    participantes = {}
    for idx, nombre in enumerate(fila_nombres):
        nombre_limpio = nombre.strip()
        if nombre_limpio and nombre_limpio not in ("PORRA MUNDIAL 2026", "Resultado real"):
            participantes[nombre_limpio] = {"col_apuesta": idx, "puntos_totales": 0}

    print("✅ Participantes mapeados:", list(participantes.keys()))

    for idx, fila in enumerate(todas_las_filas[2:], start=3):
        if len(fila) < 6: continue

        pregunta_partido = fila[2].strip()
        regla_raw = fila[3]
        col_resultado_real = fila[4]

        # Si no hay resultado real en la columna E, ignoramos la fila por completo
        if not col_resultado_real or str(col_resultado_real).strip() == "":
            continue

        # Convertir regla protegiéndola de las fechas (ej: '5/2')
        regla_puntos = celda_a_regla_partido(regla_raw)

        # Detectar tipo de fila de forma estricta
        if "/" in regla_puntos and not regla_puntos.replace("/","").replace(" ","").isdigit():
            tipo = "SI_NO"
        elif "5/2" in regla_puntos or "/" in regla_puntos or "-" in pregunta_partido:
            tipo = "PARTIDO"
        else:
            tipo = "PREGUNTA_ABIERTA"

        real_norm = normalizar_texto(str(col_resultado_real))

        # --- PROCESAR PARTIDOS ---
        if tipo == "PARTIDO":
            marcador_real = celda_a_marcador(col_resultado_real)
            if not marcador_real or '-' not in marcador_real:
                continue
            try:
                re_l, re_v = map(int, marcador_real.split("-"))
                signo_real = calcular_resultado_1X2(re_l, re_v)
            except ValueError:
                print(f"  ⚠️ No se puede parsear resultado real '{col_resultado_real}' en fila {idx}")
                continue

            for nombre, cols in participantes.items():
                val_local = fila[cols["col_apuesta"]].strip()
                val_visit = fila[cols["col_apuesta"] + 1].strip()

                # Si las celdas están vacías, no apostó
                if val_local == "" or val_visit == "":
                    continue

                # Forzar conversión anti-fechas por si acaso para las celdas individuales
                ap_l_clean = celda_a_marcador(val_local) if '-' in val_local else val_local
                ap_v_clean = celda_a_marcador(val_visit) if '-' in val_visit else val_visit

                try:
                    # Al estar divididos en dos celdas, los leemos directamente como enteros
                    ap_l = int(ap_l_clean)
                    ap_v = int(ap_v_clean)
                    signo_apuesta = calcular_resultado_1X2(ap_l, ap_v)

                    if ap_l == re_l and ap_v == re_v:
                        participantes[nombre]["puntos_totales"] += 5
                    elif signo_apuesta == signo_real:
                        participantes[nombre]["puntos_totales"] += 2
                except ValueError:
                    # Por si metieron un guion dentro de una celda individual por error
                    try:
                        marcador_combinado = celda_a_marcador(val_local)
                        ap_l, ap_v = map(int, marcador_combinado.split("-"))
                        signo_apuesta = calcular_resultado_1X2(ap_l, ap_v)
                        if ap_l == re_l and ap_v == re_v:
                            participantes[nombre]["puntos_totales"] += 5
                        elif signo_apuesta == signo_real:
                            participantes[nombre]["puntos_totales"] += 2
                    except:
                        pass

        # --- PROCESAR SÍ / NO ---
        elif tipo == "SI_NO":
            letra_real = "S" if real_norm == "SI" else "N"
            puntos_por_acierto = 0
            for p in regla_puntos.upper().replace(" ", "").split("/"):
                if p.startswith(letra_real):
                    nums = ''.join(filter(str.isdigit, p))
                    puntos_por_acierto = int(nums) if nums else 0

            for nombre, cols in participantes.items():
                apuesta_user = normalizar_texto(fila[cols["col_apuesta"]])
                letra_apuesta = "S" if apuesta_user == "SI" else "N"
                if letra_apuesta == letra_real:
                    participantes[nombre]["puntos_totales"] += puntos_por_acierto

        # --- PROCESAR PREGUNTAS ABIERTAS ---
        elif tipo == "PREGUNTA_ABIERTA":
            try:
                nums = ''.join(filter(str.isdigit, regla_puntos))
                puntos_regla = int(nums) if nums else 10
            except:
                puntos_regla = 10

            for nombre, cols in participantes.items():
                apuesta_user = normalizar_texto(fila[cols["col_apuesta"]])
                if apuesta_user == real_norm:
                    participantes[nombre]["puntos_totales"] += puntos_regla

    # Generación de Ranking
    ranking = [{"Nombre": k, "Puntos": v["puntos_totales"]} for k, v in participantes.items()]
    df_ranking = pd.DataFrame(ranking).sort_values(by="Puntos", ascending=False).reset_index(drop=True)
    df_ranking.index += 1

    print("\n🏆 CLASIFICACIÓN CORREGIDA CON ARREGLO DE FECHAS 🏆")
    print(df_ranking)

    datos_json = df_ranking.to_json(orient="records", force_ascii=False, indent=4)
    with open("clasificacion.json", "w", encoding="utf-8") as f:
        f.write(datos_json)
    print("\n💾 'clasificacion.json' actualizado con éxito!")

except Exception as e:
    import traceback
    traceback.print_exc()
    print(f"❌ Error: {e}")