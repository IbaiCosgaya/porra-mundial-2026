"""
Calculadora de puntos - Porra Mundial 2026
--------------------------------------------
Corrige los siguientes bugs del script original:

1. PUNTOS DE PARTIDO FIJOS -> ahora se leen fila a fila desde la columna D,
   porque cada fase tiene su propio valor (2/5, 3/6, 3/7, 4/8, 5/10, 6/12) y
   además hay partidos de fases distintas mezclados bajo la misma cabecera.
2. FECHA INVERTIDA (día/mes) -> se corrige el orden para que el resultado de
   partido salga como "goles_local-goles_visitante" (mes=local, día=visitante),
   y para la columna de puntos como "exacto/signo" (primer número=exacto,
   segundo=signo), que es como Google las está mostrando.
3. CONCATENACIÓN DE DÍGITOS ("10.0" -> "100") -> se reemplaza por un parseo
   con float() antes de convertir a entero.
4. DETECCIÓN DE TIPO POR PALABRAS CLAVE (frágil) -> se sustituye por
   detección explícita según el prefijo del código de pregunta (A, B, C, D)
   y el formato de la propia regla de puntos, sin listas de nombres de
   jugadores que haya que mantener a mano.
5. FINALISTAS -> se añade la regla especial de puntuación parcial
   (ambos aciertos = 15, uno = 6, ninguno = 0).
"""

import os
import json
import unicodedata
import gspread
from oauth2client.service_account import ServiceAccountCredentials

DOCUMENTO_ID = "1G4cyvvzsPw7p4yDgXiuPorsP9CDnqPuSXC-wR_H4IIY"

PALABRAS_EXCLUIDAS = (
    "PORRA MUNDIAL 2026", "RESULTADO REAL", "FECHA", "HORA", "PARTIDO",
    "EXACTO/1X2", "PUNTOS", "APUESTA", "PREGUNTA/PARTIDO",
    "EXACTO/1X2 (90MIN)", ""
)


def conectar_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    if "GOOGLE_CREDENTIALS" in os.environ:
        creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    else:
        creds = ServiceAccountCredentials.from_json_keyfile_name("credenciales.json", scope)
    cliente = gspread.authorize(creds)
    return cliente.open_by_key(DOCUMENTO_ID).worksheet("Preguntas y resultados")


def normalizar_texto(texto):
    """Quita tildes, espacios sobrantes y pasa a minúsculas para comparar sin errores."""
    if texto is None:
        return ""
    t = unicodedata.normalize("NFKD", str(texto).strip())
    t = "".join(c for c in t if not unicodedata.combining(c))
    return " ".join(t.split()).lower()


def celda_a_marcador(valor):
    """
    Convierte una celda de RESULTADO/APUESTA de partido a (goles_local, goles_visitante).
    Soporta tanto '2-1' literal como fechas mal exportadas tipo '2026-02-01'
    (que Google guarda como día=1, mes=2, y hay que leer como goles_local=mes, goles_visitante=día).
    """
    if not valor:
        return None
    v = str(valor).strip()
    if "-" not in v and "/" not in v:
        return None
    sep = "-" if "-" in v else "/"
    partes = v.split(sep)

    if len(partes) == 2:
        # Ya viene como "N-N" tal cual (no es una fecha) -> se usa directamente
        try:
            return int(partes[0]), int(partes[1])
        except ValueError:
            return None

    if len(partes) == 3:
        # Formato fecha AAAA-MM-DD (o variantes) -> goles_local = mes, goles_visitante = día
        try:
            a, b, c = partes
            if len(a) == 4:          # AAAA-MM-DD
                anio, mes, dia = a, b, c
            elif len(c) == 4:        # DD-MM-AAAA
                dia, mes, anio = a, b, c
            else:
                return None
            return int(mes), int(dia)
        except ValueError:
            return None

    return None


def celda_a_puntos_partido(valor):
    """
    Convierte la celda de la columna D (puntos del partido) a (pts_exacto, pts_signo).
    Google la muestra como 'exacto/signo' (ej. '5/2', '6/3', '7/3', '8/4', '10/5', '12/6'),
    o si arrastra formato fecha completo, como 'AAAA-MM-DD' donde día=exacto, mes=signo.
    """
    if valor is None or valor == "":
        return None
    v = str(valor).strip()
    sep = "-" if "-" in v else ("/" if "/" in v else None)
    if sep is None:
        return None
    partes = v.split(sep)

    if len(partes) == 2:
        try:
            return int(partes[0]), int(partes[1])   # (exacto, signo)
        except ValueError:
            return None

    if len(partes) == 3:
        try:
            a, b, c = partes
            if len(a) == 4:
                _, mes, dia = a, b, c
            elif len(c) == 4:
                dia, mes, _ = a, b, c
            else:
                return None
            return int(dia), int(mes)               # (exacto, signo)
        except ValueError:
            return None

    return None


def calcular_signo(goles_l, goles_v):
    if goles_l > goles_v:
        return "1"
    if goles_l < goles_v:
        return "2"
    return "X"


def es_formato_si_no(regla_texto):
    """Detecta reglas tipo 'S20 / N3' sin depender de palabras clave del enunciado."""
    if not regla_texto or "/" not in regla_texto:
        return False, None, None
    tokens = [t.strip().upper() for t in regla_texto.split("/")]
    if len(tokens) != 2:
        return False, None, None
    letras = [t[:1] for t in tokens]
    if set(letras) == {"S", "N"}:
        return True, tokens[0], tokens[1]
    return False, None, None


def puntos_de_token(token, letra_buscada):
    """Extrae el número de un token tipo 'S20' o 'N3' si coincide con la letra buscada."""
    if token[:1] != letra_buscada:
        return 0
    numeros = "".join(c for c in token if c.isdigit() or c == ".")
    return float(numeros) if numeros else 0


def calcular_puntos_finalistas(real_texto, apuesta_texto):
    """+15 si acierta los dos finalistas, +6 si acierta solo uno, 0 si ninguno."""
    def parse_equipos(txt):
        if not txt:
            return set()
        t = str(txt).replace("/", "-")
        return {normalizar_texto(x) for x in t.split("-") if normalizar_texto(x)}

    reales = parse_equipos(real_texto)
    apostados = parse_equipos(apuesta_texto)
    aciertos = reales & apostados
    if len(aciertos) == 2:
        return 15
    if len(aciertos) == 1:
        return 6
    return 0


try:
    print("🔄 Conectando con la pestaña 'Preguntas y resultados'...")
    hoja_excel = conectar_google_sheets()
    todas_las_filas = hoja_excel.get_all_values()

    fila_nombres, indice_fila_nombres = [], 0
    for idx_f, fila in enumerate(todas_las_filas):
        fila_str = [str(c).strip() for c in fila]
        if "Asier" in fila_str or "Beñat" in fila_str or "Rufo" in fila_str:
            fila_nombres, indice_fila_nombres = fila, idx_f
            break
    if not fila_nombres:
        fila_nombres, indice_fila_nombres = todas_las_filas[0], 0

    participantes = {}
    for idx, nombre in enumerate(fila_nombres):
        nombre_limpio = nombre.strip()
        if nombre_limpio and nombre_limpio.upper() not in PALABRAS_EXCLUIDAS:
            participantes[nombre_limpio] = {"col_apuesta": idx, "puntos_totales": 0}

    print("✅ Participantes localizados:", list(participantes.keys()))

    for idx, fila in enumerate(todas_las_filas[indice_fila_nombres + 1:], start=indice_fila_nombres + 2):
        if len(fila) < 5:
            continue

        codigo = fila[1].strip() if len(fila) > 1 else ""
        pregunta_partido = fila[2].strip() if len(fila) > 2 else ""
        regla_raw = fila[3].strip() if len(fila) > 3 else ""
        col_resultado_real = fila[4].strip() if len(fila) > 4 else ""

        if not col_resultado_real or col_resultado_real.upper() in PALABRAS_EXCLUIDAS:
            continue  # sin resultado real todavía -> nadie puntúa esta fila

        # --- Detección de tipo de fila ---
        marcador_real = celda_a_marcador(col_resultado_real)
        puntos_partido = celda_a_puntos_partido(regla_raw)
        es_partido = marcador_real is not None and puntos_partido is not None
        es_finalistas = codigo == "D8"
        es_numerica = codigo == "D7"
        es_si_no, token_s, token_n = es_formato_si_no(regla_raw)

        # --- ⚽ PARTIDOS ---
        if es_partido:
            re_l, re_v = marcador_real
            pts_exacto, pts_signo = puntos_partido
            signo_real = calcular_signo(re_l, re_v)

            for nombre, cols in participantes.items():
                val_apuesta = fila[cols["col_apuesta"]].strip() if cols["col_apuesta"] < len(fila) else ""
                marcador_apuesta = celda_a_marcador(val_apuesta)
                if not marcador_apuesta:
                    continue
                ap_l, ap_v = marcador_apuesta
                if ap_l == re_l and ap_v == re_v:
                    participantes[nombre]["puntos_totales"] += pts_exacto
                elif calcular_signo(ap_l, ap_v) == signo_real:
                    participantes[nombre]["puntos_totales"] += pts_signo

        # --- 🏆 FINALISTAS (regla especial) ---
        elif es_finalistas:
            for nombre, cols in participantes.items():
                val_apuesta = fila[cols["col_apuesta"]] if cols["col_apuesta"] < len(fila) else ""
                participantes[nombre]["puntos_totales"] += calcular_puntos_finalistas(col_resultado_real, val_apuesta)

        # --- 🔢 PREGUNTA NUMÉRICA (D7: nº de goles del máximo goleador) ---
        elif es_numerica:
            try:
                pts_regla = float(regla_raw)
            except ValueError:
                pts_regla = 0
            try:
                real_num = float(col_resultado_real)
            except ValueError:
                real_num = None
            for nombre, cols in participantes.items():
                val_apuesta = fila[cols["col_apuesta"]] if cols["col_apuesta"] < len(fila) else ""
                try:
                    ap_num = float(val_apuesta)
                except ValueError:
                    ap_num = None
                if real_num is not None and ap_num is not None and ap_num == real_num:
                    participantes[nombre]["puntos_totales"] += pts_regla

        # --- 🔲 SÍ / NO ---
        elif es_si_no:
            real_norm = normalizar_texto(col_resultado_real)
            letra_real = "S" if real_norm in ("si", "s") else ("N" if real_norm in ("no", "n") else None)
            if letra_real is None:
                continue
            puntos_por_acierto = max(puntos_de_token(token_s, letra_real), puntos_de_token(token_n, letra_real))

            for nombre, cols in participantes.items():
                val_apuesta = fila[cols["col_apuesta"]] if cols["col_apuesta"] < len(fila) else ""
                apuesta_norm = normalizar_texto(val_apuesta)
                letra_apuesta = "S" if apuesta_norm in ("si", "s") else ("N" if apuesta_norm in ("no", "n") else None)
                if letra_apuesta == letra_real:
                    participantes[nombre]["puntos_totales"] += puntos_por_acierto

        # --- 📝 SELECCIÓN / COMPARACIÓN (Campeón, Máximo goleador, OUT más tarde, Más goles X vs Y...) ---
        else:
            try:
                pts_fijos = float(regla_raw)
            except ValueError:
                pts_fijos = 0
            real_norm = normalizar_texto(col_resultado_real)
            for nombre, cols in participantes.items():
                val_apuesta = fila[cols["col_apuesta"]] if cols["col_apuesta"] < len(fila) else ""
                apuesta_norm = normalizar_texto(val_apuesta)
                if apuesta_norm and apuesta_norm == real_norm:
                    participantes[nombre]["puntos_totales"] += pts_fijos

    # --- Ranking final ---
    ranking = sorted(
        ({"Nombre": k, "Puntos": v["puntos_totales"]} for k, v in participantes.items()),
        key=lambda x: -x["Puntos"]
    )

    print("\n🏆 CLASIFICACIÓN 🏆")
    for i, r in enumerate(ranking, 1):
        print(f"{i}. {r['Nombre']}: {r['Puntos']} pts")

    with open("clasificacion.json", "w", encoding="utf-8") as f:
        json.dump(ranking, f, ensure_ascii=False, indent=4)
    print("\n💾 'clasificacion.json' guardado correctamente!")

except Exception as e:
    import traceback
    traceback.print_exc()
    print(f"❌ Error: {e}")
