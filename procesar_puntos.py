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
    return cliente.open_by_key(DOCUMENTO_ID).worksheet("Preguntas y resultados")

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

    formatos_fecha = ["%d-%b", "%d/%m", "%d-%m", "%b-%d", "%Y-%m-%d"]
    for formato in formatos_fecha:
        try:
            dt = datetime.strptime(v_str, formato)
            return f"{dt.day}-{dt.month}"
        except ValueError:
            continue
            
    return v_str

try:
    print("🔄 Conectando con la pestaña 'Preguntas y resultados'...")
    hoja_excel = conectar_google_sheets()
    todas_las_filas = hoja_excel.get_all_values()
    
    fila_nombres = []
    indice_fila_nombres = 0
    
    for idx_f, fila in enumerate(todas_las_filas):
        fila_str = [str(celda).strip() for celda in fila]
        if "Asier" in fila_str or "Beñat" in fila_str or "Rufo" in fila_str:
            fila_nombres = fila
            indice_fila_nombres = idx_f
            break
            
    if not fila_nombres:
        fila_nombres = todas_las_filas[0]
        indice_fila_nombres = 0

    participantes = {}
    palabras_excluidas = ("PORRA MUNDIAL 2026", "Resultado real", "FECHA", "HORA", "PARTIDO", "Exacto/1X2", "Puntos", "Apuesta", "Pregunta/Partido", "Exacto/1X2 (90min)", "")
    
    for idx, nombre in enumerate(fila_nombres):
        nombre_limpio = nombre.strip()
        if nombre_limpio and nombre_limpio not in palabras_excluidas:
            participantes[nombre_limpio] = {"col_apuesta": idx, "puntos_totales": 0}

    print("✅ Participantes reales localizados:", list(participantes.keys()))

    for idx, fila in enumerate(todas_las_filas[indice_fila_nombres + 1:], start=indice_fila_nombres + 2):
        if len(fila) < 6: continue

        pregunta_partido = fila[2].strip()
        regla_raw = fila[3].strip()
        col_resultado_real = fila[4].strip()

        if not col_resultado_real or col_resultado_real in palabras_excluidas:
            continue

        real_norm = normalizar_texto(col_resultado_real)

        # 🎯 DETECTAR TIPO DE FILA Y EXTRAER REGLA DINÁMICA
        # Buscaremos patrones tipo "5/2", "6/3", "7/3" en la columna de la regla
        es_partido = False
        puntos_pleno = 5  # Por defecto
        puntos_1X2 = 2    # Por defecto

        if "/" in regla_raw and not any(c in regla_raw.upper() for c in ("S", "N")):
            partes_regla = regla_raw.split("/")
            if len(partes_regla) == 2:
                try:
                    puntos_pleno = int(''.join(filter(str.isdigit, partes_regla[0])))
                    puntos_1X2 = int(''.join(filter(str.isdigit, partes_regla[1])))
                    es_partido = True
                except ValueError:
                    pass

        # Forzar tipo partido si la celda tiene formato marcador puro tipo "2-1"
        if "-" in col_resultado_real and not any(c.isalpha() for c in col_resultado_real):
            es_partido = True

        if es_partido:
            tipo = "PARTIDO"
        elif "/" in regla_raw and ("S" in regla_raw.upper() or "N" in regla_raw.upper()):
            tipo = "SI_NO"
        else:
            tipo = "PREGUNTA_ABIERTA"

        # --- ⚽ PROCESAR PARTIDOS (CON REGLA DINÁMICA 5/2, 6/3...) ---
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
                val_apuesta = fila[cols["col_apuesta"]].strip()
                if not val_apuesta or val_apuesta in palabras_excluidas:
                    continue

                marcador_apuesta = celda_a_marcador(val_apuesta)
                if not marcador_apuesta or '-' not in marcador_apuesta:
                    continue

                try:
                    ap_l, ap_v = map(int, marcador_apuesta.split("-"))
                    signo_apuesta = calcular_resultado_1X2(ap_l, ap_v)

                    # Aplicamos dinámicamente los puntos que correspondan a esta fila
                    if ap_l == re_l and ap_v == re_v:
                        participantes[nombre]["puntos_totales"] += puntos_pleno
                    elif signo_apuesta == signo_real:
                        participantes[nombre]["puntos_totales"] += puntos_1X2
                except ValueError:
                    pass

        # --- 🔲 PROCESAR SÍ / NO ---
        elif tipo == "SI_NO":
            letra_real = "S" if real_norm in ("SI", "S") else "N"
            puntos_por_acierto = 0
            for p in regla_raw.upper().replace(" ", "").split("/"):
                if p.startswith(letra_real):
                    nums = ''.join(filter(str.isdigit, p))
                    puntos_por_acierto = int(nums) if nums else 0

            for nombre, cols in participantes.items():
                apuesta_user = normalizar_texto(fila[cols["col_apuesta"]])
                letra_apuesta = "S" if apuesta_user in ("SI", "S") else "N"
                if letra_apuesta == letra_real:
                    participantes[nombre]["puntos_totales"] += puntos_por_acierto

        # --- 📝 PROCESAR PREGUNTAS ABIERTAS ---
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

    # Ranking Final Ordenado
    ranking = [{"Nombre": k, "Puntos": v["puntos_totales"]} for k, v in participantes.items()]
    df_ranking = pd.DataFrame(ranking).sort_values(by="Puntos", ascending=False).reset_index(drop=True)
    df_ranking.index += 1

    print("\n🏆 CLASIFICACIÓN REAL CALCULADA (PUNTOS DINÁMICOS) 🏆")
    print(df_ranking)

    datos_json = df_ranking.to_json(orient="records", force_ascii=False, indent=4)
    with open("clasificacion.json", "w", encoding="utf-8") as f:
        f.write(datos_json)
    print("\n💾 'clasificacion.json' sincronizado correctamente!")

except Exception as e:
    import traceback
    traceback.print_exc()
    print(f"❌ Error: {e}")
