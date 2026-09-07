# -*- coding: utf-8 -*-
"""
Streamlit — Descarga de "Medidores de Generación" (COES)

Ejecutar:
    pip install streamlit requests pandas openpyxl
    streamlit run app_medidores_generacion.py
"""

import io
import re
import time
import unicodedata
from calendar import monthrange
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

# ─────────────────────────────────────────────────────────────
#  CONSTANTES DEL ENDPOINT / FORMULARIO (extraídas del HTML real)
# ─────────────────────────────────────────────────────────────
TIPOS_GENERACION = {
    "EÓLICA": "4",
    "HIDROELÉCTRICA": "1",
    "SOLAR": "3",
    "TERMOELÉCTRICA": "2",
}

CENTRAL_OPCIONES = {
    "TODOS": "0",
    "COES": "1",
    "GENERACION RER": "3",
}

# Mapeo nombre -> ID interno COES, extraído del <select id="cbEmpresas"> real del formulario
EMPRESAS = {
    "ACCIONA ENERGIA PERU S.A.C": "14260",
    "ADINELSA ADN": "69",
    "AGRO INDUSTRIAL PARAMONGA S.A.": "15214",
    "AGROAURORA S.A.C.": "11772",
    "AGROINDUSTRIAS SAN JACINTO S.A.A.": "12439",
    "AGROLMOS SOCIEDAD ANONIMA - AGROLMOS S.A.": "11777",
    "AGUAS Y ENERGIA PERU": "10481",
    "ANDEAN POWER S.A.C.": "12056",
    "ASOCIACIÓN SANTA LUCIA DE CHACAS": "12362",
    "ATRIA ENERGIA S.A.C.": "13196",
    "BIOENERGIA DEL CHIRA S.A.": "12896",
    "CASA GRANDE S.A.A.": "10628",
    "CELARIS ENERGY S.A.": "15707",
    "CELEPSA": "10420",
    "CELEPSA RENOVABLES S.R.L.": "12708",
    "CENTRALES SANTA ROSA S.A.C.": "13165",
    "CHINANGO S.A.C.": "10901",
    "COENERGY S.A.C.": "15571",
    "COLCA SOLAR S.A.C.": "12584",
    "CORPORACION MINERA DEL PERU S.A.": "11095",
    "EGASA": "17",
    "EGEJUNIN": "11153",
    "EGEMSA": "58",
    "EGESUR": "19",
    "ELECTRICA YANAPAMPA SAC": "10684",
    "ELECTRO ORIENTE": "30",
    "ELECTRO SUR ESTE": "27",
    "ELECTRO UCAYALI": "40",
    "ELECTRO ZAÑA S.A.C.": "11228",
    "ELECTRONOROESTE S.A.": "23",
    "ELECTRONORTE S.A.": "24",
    "ELECTROPERU": "2",
    "EMPRESA DE GENERACION ELECTRICA CANCHAYLLO SAC": "11389",
    "EMPRESA DE GENERACION ELECTRICA SANTA ANA S.A.C.": "14173",
    "EMPRESA DE GENERACION HUALLAGA": "11412",
    "EMPRESA DE GENERACION HUANZA": "206",
    "EMPRESA ELECTRICA AGUA AZUL": "11544",
    "EMPRESA ELECTRICA RIO DOBLE": "11058",
    "ENEL GENERACION PIURA S.A.": "12097",
    "ENERGÍA EÓLICA S.A.": "10552",
    "ENERGIA RENOVABLE DEL SUR S.A.": "11429",
    "ENERGIA RENOVABLE LA JOYA S.A.": "12884",
    "ENGIE ENERGIA PERU S.A.A.": "15624",
    "EOLICA CARAVELI S.A.C.": "15392",
    "FENIX POWER PERÚ": "10725",
    "GENERACIÓN ANDINA S.A.C.": "11527",
    "GENERADORA DE ENERGÍA DEL PERÚ": "10647",
    "GM OPERACIONES S.A.C.": "14973",
    "GR CORTARRAMA SOCIEDAD ANONIMA CERRADA": "11981",
    "GR PAINO SOCIEDAD ANONIMA CERRADA": "11840",
    "GR TARUCA SOCIEDAD ANONIMA CERRADA": "11841",
    "GR VALE S.A.C.": "12974",
    "HIDROCAÑETE S.A.": "10974",
    "HIDROELECTRICA HUANCHOR S.A.C.": "11258",
    "HIDROELECTRICA RAPAZ S.A.C.": "11644",
    "HUAURA POWER GROUP S.A.": "11444",
    "HYDRO GLOBAL PERÚ S.A.C.": "11940",
    "HYDRO PATAPO S.A.C.": "12364",
    "ILLAPU ENERGY": "11149",
    "INFRAESTRUCTURAS Y ENERGIAS DEL PERU S.A.C.": "11528",
    "INLAND ENERGY SAC": "12634",
    "INTI JOYA S.A.C.": "15849",
    "INVERSIONES SHAQSHA S.A.C.": "12624",
    "JOYA SOLAR S.A.C.": "13726",
    "KALLPA GENERACION S.A.": "12479",
    "KONDU SAC": "14807",
    "LA VIRGEN": "11185",
    "LUZ DEL SUR": "13",
    "MAJA ENERGIA S.A.C.": "10916",
    "MAJES ARCUS S.A.C.": "13965",
    "MINERA CERRO VERDE": "67",
    "MINERA CORONA": "108",
    "MOQUEGUA FV S.A.C.": "11217",
    "ORAZUL ENERGY PERÚ": "12480",
    "ORYGEN PERU S.A.A.": "15259",
    "PANAMERICANA SOLAR SAC.": "11102",
    "PARQUE EOLICO MARCONA S.A.C.": "13120",
    "PARQUE EOLICO TRES HERMANAS S.A.C.": "11218",
    "PETRAMAS": "11063",
    "PETROPERU": "149",
    "PLANTA DE RESERVA FRIA DE GENERACION DE ETEN S.A.": "11323",
    "REPARTICIÓN ARCUS S.A.C.": "13966",
    "SAMAY I S.A.C": "15009",
    "SAN GABAN": "61",
    "SDE PIURA": "10913",
    "SDF ENERGIA SAC": "10587",
    "SHOUGESA": "8",
    "SINERSA": "138",
    "STATKRAFT S.A": "12758",
    "TACNA SOLAR SAC.": "11103",
    "TERMOCHILCA S.A.C.": "15080",
    "TERMOSELVA": "10",
    "TRANSMISION ANDINA DE GENERACION S.A.C.": "15683",
    "UNACEM PERU S.A.": "14342",
    "VARI ENERGIA S.A.C.": "13984",
    "YURA": "167",
    "AGRO INDUSTRIAL PARAMONGA": "10422",
    "CEMENTO ANDINO": "180",
    "CERRO DEL AGUILA S.A.": "11146",
    "COGENERACION OQUENDO SAC": "15014",
    "EDEGEL": "4",
    "EEPSA": "5",
    "EGENOR": "9",
    "ELECTRICA SANTA ROSA": "76",
    "EMPRESA CONCESIONARIA ENERGIA LIMPIA SAC": "11563",
    "EMPRESA DE GENERACIÓN ELÉCTRICA CHEVES S.A.": "10636",
    "EMPRESA DE GENERACION ELECTRICA RIO BAÑOS S.A.C.": "11129",
    "EMPRESA DE GENERACION ELECTRICA SANTA ANA": "11509",
    "ENEL GENERACION PERU S.A.A.": "12096",
    "ENEL GREEN POWER PERU S.A.": "11395",
    "ENEL GREEN POWER PERU S.A.C": "13783",
    "ENERSUR": "18",
    "ENGIE": "48",
    "HIDROELECTRICA SANTA CRUZ": "10582",
    "HIDROMARAÑON": "11064",
    "KALLPA GENERACION": "47",
    "MAJES ARCUS": "11100",
    "MAPLE ETANOL": "10755",
    "ORAZUL ENERGY": "12190",
    "PARQUE EOLICO MARCONA S.R.L.": "11053",
    "PERUANA DE INVERSIONES EN ENERGIAS RENOVABLES S.A.": "10984",
    "REPARTICIÓN ARCUS": "11101",
    "SAMAY I S.A.": "11486",
    "SN POWER": "6",
    "STATKRAFT": "11567",
    "TERMOCHILCA": "10767",
    "UNION ANDINA DE CEMENTO": "11894",
}

PARAMETROS_EXPORTAR = {
    "Potencia Activa (MW)": "1",
    "Potencia Reactiva (MW)": "5",
    "Servicios Auxiliares": "3",
    "Potencia Reactiva Capacitiva (MVAR)": "2",
    "Potencia Reactiva Inductiva (MVAR)": "4",
}

FORMATOS = {
    "Excel Horizontal": "1",
    "Excel Vertical": "2",
    "CSV": "3",
}

FORMATO_EXTENSION = {"1": "xlsx", "2": "xlsx", "3": "csv"}

URL_PAGINA   = "https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion"
URL_EXPORTAR = "https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion/Exportar"

def validar_antes_de_exportar(fecha_inicial: date, fecha_final: date, parametros: str, tipo: str) -> str | None:
    """Replica las validaciones que hace medidores.js (exportarFormato) ANTES
    de armar la URL de exportación, para fallar rápido con un mensaje claro
    en vez de gastar una petición que el servidor rechazaría igual.
    Devuelve un mensaje de error, o None si está todo bien."""
    if (fecha_final - fecha_inicial).days > 31:
        return "El lapso de tiempo no puede ser mayor a 1 mes."
    lista_parametros = [p for p in parametros.split(",") if p]
    if tipo == "3" and len(lista_parametros) != 1:
        return "Para la exportación a CSV solo debe seleccionar un parámetro."
    if not lista_parametros:
        return "Seleccione un parámetro a exportar."
    return None

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": URL_PAGINA,
    # El flujo real (medidores.js) exporta con window.open(url) -- una
    # navegación normal del navegador, no una llamada $.ajax(). Por eso NO
    # se manda "X-Requested-With: XMLHttpRequest" aquí (antes sí se mandaba,
    # heredado de cuando el código pensaba que "Exportar" era un endpoint
    # AJAX como los demás, p.ej. "empresas" o "ValidarParametro", que sí lo
    # necesitan y si acaso se llaman aparte).
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}




def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    s.get(URL_PAGINA, timeout=30)  # cookies de sesión
    return s


def _nombre_desde_content_disposition(resp: requests.Response, default: str) -> str:
    cd = resp.headers.get("Content-Disposition", "")
    m = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)"?', cd)
    if m:
        return m.group(1)
    return default


def descargar_medidores_generacion(
    fecha_inicial: str,
    fecha_final: str,
    empresas: str,
    tipos_generacion: str,
    central: str,
    parametros: str,
    tipo: str,
    tipos_empresa: str = "",
):
    """
    Replica el flujo REAL actual del sitio (medidores.js -> exportarFormato):
    una ÚNICA petición GET a ".../medidoresgeneracion/Exportar" con todos los
    parámetros en la URL. Así es como lo hace el botón "Exportar Datos" del
    portal: arma `controlador + "Exportar?" + $.param(modelo)` y hace
    `window.open(enlace)` -- el navegador simplemente navega a esa URL y
    descarga el archivo que el servidor devuelve directo, sin pasos previos.

    Esto reemplaza un flujo anterior de 3 pasos (validarexportacion + POST
    exportar + GET descargar) que ya no coincide con el sitio real: ese flujo
    devolvía 404 en el paso "exportar" para absolutamente todos los meses
    probados, lo que confirmó que esos endpoints separados ya no existen (se
    revisó el medidores.js real servido por COES para corregirlo).

    Devuelve (contenido_bytes, nombre_archivo, mensaje_error).
    """
    s = _session()

    params = {
        "fechaInicial": fecha_inicial,
        "fechaFinal": fecha_final,
        "tiposEmpresa": tipos_empresa,
        "empresas": empresas,
        "tiposGeneracion": tipos_generacion,
        "central": central,
        "parametros": parametros,
        "tipo": tipo,
    }

    try:
        resp = s.get(URL_EXPORTAR, params=params, timeout=90)
    except requests.RequestException as e:
        return None, None, f"Error de conexión en Exportar: {e}"

    if resp.status_code != 200:
        return None, None, f"HTTP {resp.status_code} en Exportar: {resp.text[:300]}"

    content_type = resp.headers.get("Content-Type", "")
    if "text/html" in content_type:
        # El servidor devolvió una página (de error, de login, de aviso) en
        # vez del archivo -- por ejemplo, si algún parámetro no aplica para
        # ese mes/central en particular.
        return None, None, f"COES devolvió una página HTML en vez de un archivo: {resp.text[:300]}"

    if not resp.content or len(resp.content) < 100:
        return None, None, "El servidor devolvió un archivo vacío o demasiado pequeño."

    ext_default = FORMATO_EXTENSION.get(tipo, "xlsx")
    nombre_default = (
        f"MedidoresGeneracion_{fecha_inicial.replace('/', '')}_"
        f"{fecha_final.replace('/', '')}.{ext_default}"
    )
    nombre = _nombre_desde_content_disposition(resp, default=nombre_default)
    nombre = unicodedata.normalize("NFKD", nombre).encode("ascii", "ignore").decode()

    return resp.content, nombre, None


def dividir_en_meses(fecha_inicio: date, fecha_fin: date):
    """
    Divide [fecha_inicio, fecha_fin] en tramos que respetan el límite de
    1 mes que exige el servidor. Cada tramo va del día de inicio hasta el
    mismo día del mes siguiente menos uno (o el fin real, lo que sea antes).
    """
    tramos = []
    actual = fecha_inicio
    while actual <= fecha_fin:
        # Fin de tramo: mismo día un mes después, menos 1 día
        mes = actual.month + 1
        anio = actual.year
        if mes > 12:
            mes = 1
            anio += 1
        ultimo_dia_mes_destino = monthrange(anio, mes)[1]
        dia = min(actual.day, ultimo_dia_mes_destino)
        fin_tramo = date(anio, mes, dia) - timedelta(days=1)

        if fin_tramo > fecha_fin:
            fin_tramo = fecha_fin

        tramos.append((actual, fin_tramo))
        actual = fin_tramo + timedelta(days=1)

    return tramos


def _detectar_fila_en_vista(vista: pd.DataFrame) -> int:
    """Busca, entre las primeras filas de un archivo leído sin encabezado, cuál
    tiene más celdas que coinciden con palabras clave típicas del encabezado
    real de un reporte de COES. Igual estrategia que se usó para el reporte de
    Mantenimientos (dashboard_manttos.py), porque estos reportes suelen traer
    filas de título/fecha arriba de los encabezados reales, y la cantidad de
    esas filas puede variar."""
    palabras_clave = (
        "EMPRESA", "CENTRAL", "MEDIDOR", "FECHA", "PARAMETRO", "PARÁMETRO",
        "POTENCIA", "GENERACION", "GENERACIÓN", "HORA", "TIPO",
    )
    mejor_fila, mejor_score = 0, -1
    for i in range(len(vista)):
        valores = [str(v).upper().strip() for v in vista.iloc[i] if pd.notna(v)]
        score = sum(1 for v in valores if any(k in v for k in palabras_clave))
        if score > mejor_score:
            mejor_score, mejor_fila = score, i
    return mejor_fila


# COES ha usado más de un nombre para la misma empresa en distintos períodos
# (p.ej. "ENGIE" en registros de 2025, y la razón social completa "ENGIE
# ENERGIA PERU S.A.A." desde 2025-2026 en adelante -- mismas 24 unidades de
# generación en ambos casos). Sin esto, un rango que cruza ese cambio queda
# separado como si fueran dos empresas distintas. Las claves van en
# MAYÚSCULAS sin tildes para que la comparación sea insensible a acentos.
NORMALIZACION_EMPRESA = {
    "ENGIE": "ENGIE ENERGIA PERU S.A.A.",
}


def _normalizar_empresa(nombre):
    """Unifica nombres de empresa que representan la misma compañía pero
    aparecen distinto según el período (ver NORMALIZACION_EMPRESA)."""
    if nombre is None or (isinstance(nombre, float) and pd.isna(nombre)):
        return nombre
    clave = unicodedata.normalize("NFKD", str(nombre).strip().upper())
    clave = "".join(c for c in clave if not unicodedata.combining(c))
    return NORMALIZACION_EMPRESA.get(clave, nombre)


def _clave_normalizada(nombre) -> str:
    """Mayúsculas, sin tildes, sin espacios de más -- para comparar nombres
    (de empresa o central) de forma insensible a esas diferencias menores."""
    if nombre is None or (isinstance(nombre, float) and pd.isna(nombre)):
        return ""
    clave = unicodedata.normalize("NFKD", str(nombre).strip().upper())
    clave = "".join(c for c in clave if not unicodedata.combining(c))
    return " ".join(clave.split())


# Lista por defecto de centrales a incluir en el resultado final. Se puede
# editar libremente acá, y también se puede ajustar desde la app (el
# multiselect en "Opciones avanzadas" parte de esta lista, pero se puede
# cambiar sin tocar código). Una lista vacía == incluir todas las centrales.
CENTRALES_A_INCLUIR_DEFAULT = [
    "C.E. PUNTA LOMITAS",
    "C.E. PUNTA LOMITAS_EXP",
    "C.H. QUITARACSA",
    "C.H. YUNCAN",
    "C.S. INTIPAMPA",
    "C.S. EXPANSIÓN INTIPAMPA",
    "C.E. DUNA",
    "C.E. HUAMBOS",
]


def _aplanar_encabezado_vertical(df_final: pd.DataFrame, col_fecha_hora: tuple) -> pd.DataFrame:
    """Convierte el encabezado de 2 niveles (Empresa, Central) del formato
    Vertical a uno solo, plano: la primera columna se llama "Fecha/Hora" y
    las demás llevan directamente el nombre de la Central (sin la fila
    "Empresa"/"Central" que pandas agrega por defecto al exportar un
    DataFrame con columnas MultiIndex + índice con nombre, que queda
    confuso: "Fecha/Hora" termina en su propia fila, sin nada al lado).

    Si dos empresas distintas comparten el mismo nombre de central (caso
    raro, pero posible), esa central puntual se distingue como
    "Central (Empresa)" para no mezclar sus datos bajo un solo encabezado;
    el resto queda simplemente como "Central"."""
    conteo_centrales = {}
    for col in df_final.columns:
        if col == col_fecha_hora:
            continue
        _empresa, central = col
        conteo_centrales[central] = conteo_centrales.get(central, 0) + 1

    df_plano = df_final.copy()
    nuevas_columnas = []
    for col in df_plano.columns:
        if col == col_fecha_hora:
            nuevas_columnas.append("Fecha/Hora")
        else:
            empresa, central = col
            if conteo_centrales[central] > 1:
                nuevas_columnas.append(f"{central} ({empresa})")
            else:
                nuevas_columnas.append(central)
    df_plano.columns = nuevas_columnas
    return df_plano


def _parsear_vertical_coes(contenido: bytes) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Parser específico para el formato 'Excel Vertical' de COES.

    Este formato trae un encabezado JERÁRQUICO de 4 filas (código de medidor,
    empresa, central, unidad de generación) y una fila por cada intervalo de
    15 minutos, con una columna por cada unidad de generación.

    Esta función reproduce esa estructura ANCHA (filas = Fecha/Hora), pero
    SUMA las unidades de una misma central en una sola columna -- una central
    con varias unidades (p.ej. "C.E. PUNTA LOMITAS-BL1" y "-BL2") queda como
    una única columna con la generación total de la central. El encabezado
    final tiene 2 niveles (Empresa, Central), no 3 -- ya no aparece la
    Unidad.

    Como no todas las centrales generan todos los meses, el conjunto de
    columnas cambia de un mes a otro. Al consolidar varios meses con
    pd.concat(axis=0), las columnas se UNEN automáticamente. A diferencia de
    versiones anteriores, los huecos (central que no generó ese mes, o
    intervalo sin dato) quedan en 0 en vez de NaN -- Power BI no lee bien los
    NaN. Para no perder la información de "esta central no tuvo dato ese
    mes" (distinto de "generó cero"), esta función devuelve además una
    auditoría: (tabla, auditoria), donde auditoria tiene una fila por Central
    de este mes con una columna booleana "Operó".
    """
    vista = pd.read_excel(io.BytesIO(contenido), header=None)

    fila_fecha_hora = None
    for i in range(min(30, len(vista))):
        valores = [str(v).strip().upper() for v in vista.iloc[i] if pd.notna(v)]
        if any("FECHA/HORA" in v for v in valores):
            fila_fecha_hora = i
            break
    if fila_fecha_hora is None or fila_fecha_hora + 4 >= len(vista):
        raise RuntimeError(
            "No se encontró la fila 'FECHA/HORA' con el encabezado jerárquico "
            "esperado (formato Excel Vertical inesperado)."
        )

    fila_medidor = vista.iloc[fila_fecha_hora]
    fila_empresa = vista.iloc[fila_fecha_hora + 1].ffill()
    fila_central = vista.iloc[fila_fecha_hora + 2].ffill()
    fila_unidad = vista.iloc[fila_fecha_hora + 3]

    col_fecha_hora, col_total = None, None
    col_indices, columnas = [], []
    for col_idx in range(len(fila_medidor)):
        valor = fila_medidor.iloc[col_idx]
        etiqueta = str(valor).strip().upper() if pd.notna(valor) else ""
        if etiqueta == "FECHA/HORA":
            col_fecha_hora = col_idx
        elif etiqueta == "TOTAL":
            col_total = col_idx
        elif pd.notna(fila_unidad.iloc[col_idx]):
            col_indices.append(col_idx)
            empresa_raw = fila_empresa.iloc[col_idx] if pd.notna(fila_empresa.iloc[col_idx]) else ""
            columnas.append((
                _normalizar_empresa(empresa_raw) if empresa_raw != "" else "",
                fila_central.iloc[col_idx] if pd.notna(fila_central.iloc[col_idx]) else "",
                fila_unidad.iloc[col_idx],
            ))

    if col_fecha_hora is None or not col_indices:
        raise RuntimeError(
            "No se pudieron identificar las columnas de Fecha/Hora o de "
            "unidades de generación en el archivo (formato inesperado)."
        )

    datos = vista.iloc[fila_fecha_hora + 4:].reset_index(drop=True)
    datos = datos.dropna(how="all").reset_index(drop=True)
    fecha_hora = pd.to_datetime(datos.iloc[:, col_fecha_hora], errors="coerce", dayfirst=True)

    tabla = datos.iloc[:, col_indices].copy()
    tabla.columns = pd.MultiIndex.from_tuples(columnas, names=["Empresa", "Central", "Unidad"])
    if col_total is not None:
        tabla[("(TOTAL)", "(TOTAL)", "(TOTAL)")] = datos.iloc[:, col_total]

    # Agrupar las unidades de cada central, para sumarlas en una sola
    # columna (el encabezado final solo muestra Empresa, Central -- ya no
    # Unidad). Antes de sumar, se registra si la central tuvo AL MENOS UN
    # valor real (no nulo) en este mes -- esa es la auditoría de "operó/no
    # operó", y hay que calcularla ANTES de sumar porque la suma rellena con
    # 0 (no dejaría forma de distinguir "generó 0" de "no hay dato").
    grupos = {}
    for empresa, central, _unidad in tabla.columns:
        grupos.setdefault((empresa, central), []).append((empresa, central, _unidad))

    auditoria_filas = []
    tabla_por_central = pd.DataFrame(index=tabla.index)
    for (empresa, central), cols in grupos.items():
        sub = tabla[cols]
        opero = bool(sub.notna().any().any())
        if (empresa, central) != ("(TOTAL)", "(TOTAL)"):
            auditoria_filas.append({"Empresa": empresa, "Central": central, "Operó": opero})
        # Con al menos un dato real, se suma tratando los huecos como 0
        # (comportamiento normal de .sum()); si la central no operó nada
        # este mes, la fila igual queda en 0 -- ya está reflejado en la
        # auditoría de arriba, así que no hace falta dejarla en blanco.
        tabla_por_central[(empresa, central)] = sub.sum(axis=1)
    tabla_por_central.columns = pd.MultiIndex.from_tuples(
        tabla_por_central.columns, names=["Empresa", "Central"]
    )
    tabla = tabla_por_central
    auditoria = pd.DataFrame(auditoria_filas)

    tabla.index = fecha_hora
    tabla = tabla[tabla.index.notna()]  # descarta filas de resumen/nota al pie (sin fecha válida)
    tabla.index.name = "Fecha/Hora"
    return tabla.reset_index(), auditoria


def _quitar_filas_resumen_pie(df: pd.DataFrame) -> pd.DataFrame:
    """Los reportes de COES suelen traer, después de los datos reales de cada
    mes, un bloque de resumen y una nota al pie ('TOTAL ENERGÍA (MWh)',
    'TOTAL POTENCIA MÁXIMA (MW)', 'TOTAL POTENCIA MÍNIMA (MW)', 'Leyenda',
    '(*) Incluye a las centrales...'). Esas filas no son lecturas reales y
    quedan mezcladas con los datos si no se filtran. Se identifican porque no
    tienen una fecha válida en la columna 'FECHA' (tienen texto en su lugar),
    así que se descartan buscando esa columna y quedándose solo con las filas
    donde sí se puede interpretar como fecha."""
    col_fecha = next((c for c in df.columns if str(c).strip().upper() == "FECHA"), None)
    if col_fecha is None:
        return df
    fechas = pd.to_datetime(df[col_fecha], errors="coerce", dayfirst=True)
    return df[fechas.notna()].reset_index(drop=True)


def _parsear_contenido_coes(contenido: bytes, formato_val: str) -> tuple[pd.DataFrame, pd.DataFrame | None]:
    """Convierte los bytes crudos (Excel o CSV) devueltos por COES en un
    DataFrame limpio. Devuelve (df, auditoria).

    - 'Excel Vertical' (formato_val == '2'): tiene un encabezado jerárquico de
      varias filas (medidor/empresa/central/unidad) y una columna por unidad
      de generación, que varía mes a mes. Usa el parser dedicado
      _parsear_vertical_coes, que lo convierte a formato ancho sumado por
      central y devuelve también la auditoría de qué central operó ese mes.
    - 'Excel Horizontal' (formato_val == '1') y CSV: tienen un encabezado de
      una sola fila (FECHA, PUNTO MEDICIÓN, EMPRESA, CENTRAL, UNIDAD, ... y
      columnas de horario 00:15-24:00). Usan detección dinámica de la fila de
      encabezado (por si trae metadata arriba), descartan columnas líder
      vacías, y quitan las filas de resumen/pie de página que COES agrega
      después de los datos reales de cada mes (_quitar_filas_resumen_pie).
      Acá Central es un valor de fila (no un encabezado), así que la
      auditoría no aplica -- se devuelve None.
    """
    if formato_val == "2":  # Excel Vertical
        return _parsear_vertical_coes(contenido)

    if formato_val == "3":  # CSV
        vista = pd.read_csv(io.BytesIO(contenido), sep=None, engine="python", header=None, nrows=10)
        fila_header = _detectar_fila_en_vista(vista)
        df = pd.read_csv(io.BytesIO(contenido), sep=None, engine="python", header=fila_header)
    else:  # Excel Horizontal (formato_val == "1")
        vista = pd.read_excel(io.BytesIO(contenido), header=None, nrows=10)
        fila_header = _detectar_fila_en_vista(vista)
        df = pd.read_excel(io.BytesIO(contenido), header=fila_header)

    while (
        len(df.columns) > 0
        and str(df.columns[0]).startswith("Unnamed")
        and df[df.columns[0]].isna().all()
    ):
        df = df.drop(columns=df.columns[0])

    df = df.dropna(how="all").reset_index(drop=True)
    df = _quitar_filas_resumen_pie(df)

    col_empresa = next((c for c in df.columns if str(c).strip().upper() == "EMPRESA"), None)
    if col_empresa is not None:
        df[col_empresa] = df[col_empresa].map(_normalizar_empresa)

    return df, None


@st.cache_data(show_spinner=False, ttl=3600)
def _descargar_y_leer(ini, fin, empresas_val, tipo_gen_val, central_val, parametros_val, formato_val):
    """Descarga un tramo y lo convierte a DataFrame. Lanza excepción si algo falla.

    Cacheado por (ini, fin, filtros, formato): si este mes ya se descargó bien
    antes (en esta sesión, hasta 1 hora), no se vuelve a pedir al servidor.
    Esto es clave porque las descargas son secuenciales -así que si 5 de 7
    meses ya funcionaron y solo fallan 2, reintentar (o volver a apretar
    'Descargar') no debería re-descargar los 5 que ya estaban bien. Si la
    descarga falla (excepción), Streamlit NO cachea ese resultado, así que el
    reintento sí vuelve a pedirlo al servidor como corresponde."""
    rango_str = f"{ini.strftime('%d/%m/%Y')} - {fin.strftime('%d/%m/%Y')}"

    error_validacion = validar_antes_de_exportar(ini, fin, parametros_val, formato_val)
    if error_validacion:
        raise RuntimeError(error_validacion)

    contenido, nombre, error = descargar_medidores_generacion(
        fecha_inicial=ini.strftime("%d/%m/%Y"),
        fecha_final=fin.strftime("%d/%m/%Y"),
        empresas=empresas_val,
        tipos_generacion=tipo_gen_val,
        central=central_val,
        parametros=parametros_val,
        tipo=formato_val,
    )
    if error:
        raise RuntimeError(error)

    df, auditoria = _parsear_contenido_coes(contenido, formato_val)

    if df.empty:
        raise RuntimeError("El archivo se descargó pero no contiene filas (posible mezcla de tramos).")

    if isinstance(df.columns, pd.MultiIndex):
        # Formato Vertical (tabla ancha, encabezado de 2 niveles): no se le
        # agrega Tramo_Consultado porque no calza con columnas de varios
        # niveles, y de todas formas la columna Fecha/Hora ya identifica a
        # qué mes pertenece cada fila.
        if auditoria is not None and not auditoria.empty:
            auditoria = auditoria.copy()
            auditoria.insert(0, "Tramo", rango_str)
        return df, auditoria

    df.insert(0, "Tramo_Consultado", rango_str)
    return df, auditoria


def _procesar_tramos_secuencial(
    tramos, empresas_val, tipo_gen_val, central_val, parametros_val, formato_val,
    max_reintentos, progreso, estado,
):
    """
    Descarga cada mes UNO A LA VEZ (nada de hilos/paralelismo): se pide el
    archivo, se convierte a DataFrame apenas llega, y se guarda en una lista.
    Al final, todos los DataFrames se unen con pd.concat en un solo
    consolidado.

    Se hace estrictamente secuencial -y no en paralelo- porque el servidor de
    COES arma este reporte usando un único archivo fijo en disco (no uno por
    sesión): pedir varios meses al mismo tiempo hace que se pisen entre sí
    ('Error saving file...' / HTTP 500). Descargar de a uno es más lento,
    pero elimina ese choque de raíz en vez de solo mitigarlo.
    """
    dataframes = []
    auditorias = []
    errores = []
    total = len(tramos)

    for i, (ini, fin) in enumerate(tramos, start=1):
        rango_str = f"{ini.strftime('%d/%m/%Y')} - {fin.strftime('%d/%m/%Y')}"
        df_mes = None
        auditoria_mes = None
        ultimo_error = None

        for intento in range(max_reintentos + 1):
            sufijo = f" (reintento {intento}/{max_reintentos})" if intento else ""
            estado.text(f"Descargando {i}/{total} ({rango_str}){sufijo}...")
            try:
                df_mes, auditoria_mes = _descargar_y_leer(
                    ini, fin, empresas_val, tipo_gen_val, central_val, parametros_val, formato_val,
                )
                break
            except Exception as e:
                ultimo_error = str(e)
                if intento < max_reintentos:
                    time.sleep(2 * (intento + 1))  # pausa creciente antes de reintentar

        if df_mes is not None:
            dataframes.append(df_mes)
            if auditoria_mes is not None and not auditoria_mes.empty:
                auditorias.append(auditoria_mes)
        else:
            errores.append((rango_str, f"Falló tras {max_reintentos + 1} intento(s). Último error: {ultimo_error}"))

        progreso.progress(i / total)

    return dataframes, auditorias, errores


# ─────────────────────────────────────────────────────────────
#  INTERFAZ STREAMLIT
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="Medidores de Generación — COES", page_icon="⚡", layout="centered")

st.title("⚡ Medidores de Generación — COES")
st.caption("Descarga directa desde el portal COES (mediciones/medidoresgeneracion)")

st.subheader("Rango de fechas")
st.caption(
    "El portal solo permite consultar 1 mes por request; si eliges un rango "
    "mayor, la app descarga mes por mes y consolida todo en un solo Excel."
)
col1, col2 = st.columns(2)
with col1:
    fecha_inicial = st.date_input("Fecha inicial", value=date.today().replace(day=1))
with col2:
    fecha_final = st.date_input("Fecha final", value=date.today())

st.divider()
st.subheader(f"Empresa ({len(EMPRESAS)} agentes)")
todos_empresas = st.checkbox("TODOS los agentes", value=True, key="chk_empresas")
empresas_sel = st.multiselect(
    "Selecciona empresa(s)",
    options=sorted(EMPRESAS.keys()),
    default=[],
    disabled=todos_empresas,
    help="Escribe para filtrar por nombre.",
)

st.divider()
st.subheader("Tipo de Generación")
todos_tipo_gen = st.checkbox("TODOS los tipos", value=True, key="chk_tipo_gen")
tipo_gen_sel = st.multiselect(
    "Selecciona tipo(s) de generación",
    options=list(TIPOS_GENERACION.keys()),
    default=list(TIPOS_GENERACION.keys()),
    disabled=todos_tipo_gen,
)

st.divider()
st.subheader("Central / Alcance")
central_sel = st.radio(
    "Filtro de central",
    options=list(CENTRAL_OPCIONES.keys()),
    index=0,  # TODOS
    horizontal=True,
)

st.divider()
st.subheader("Parámetro (a exportar)")
todos_parametros = st.checkbox("TODOS los parámetros", value=True, key="chk_param")
parametros_sel = st.multiselect(
    "Selecciona parámetro(s)",
    options=list(PARAMETROS_EXPORTAR.keys()),
    default=list(PARAMETROS_EXPORTAR.keys()),
    disabled=todos_parametros,
)

st.divider()
st.subheader("Formato de salida")
formato_sel = st.radio(
    "Formato",
    options=list(FORMATOS.keys()),
    index=0,  # Excel Horizontal
    horizontal=True,
)

st.divider()
with st.expander("⚙️ Opciones avanzadas", expanded=False):
    st.caption(
        "Las descargas de varios meses se piden una por una (nunca en paralelo), "
        "porque el servidor de COES arma este reporte en un único archivo fijo en "
        "disco -no uno por sesión- y pedir varios meses al mismo tiempo los hace "
        "chocar entre sí (error 'Error saving file...' o HTTP 500)."
    )
    max_reintentos = st.slider(
        "Reintentos automáticos por tramo fallido", 0, 5, 2,
        help="Si un mes falla (por ejemplo por un error temporal del servidor), "
             "la app lo vuelve a intentar automáticamente hasta este número de veces "
             "antes de reportarlo como error definitivo.",
    )


st.divider()
submitted = st.button("📥 Descargar", use_container_width=True, type="primary")


if submitted:
    if fecha_inicial > fecha_final:
        st.error("La fecha inicial no puede ser posterior a la fecha final.")
        st.stop()

    if not todos_empresas and not empresas_sel:
        st.error("Selecciona al menos una empresa (o marca TODOS los agentes).")
        st.stop()
    if not todos_tipo_gen and not tipo_gen_sel:
        st.error("Selecciona al menos un tipo de generación (o marca TODOS).")
        st.stop()
    if not todos_parametros and not parametros_sel:
        st.error("Selecciona al menos un parámetro (o marca TODOS).")
        st.stop()

    empresas_val = (
        ",".join(EMPRESAS.values()) if todos_empresas
        else ",".join(EMPRESAS[nombre] for nombre in empresas_sel)
    )

    tipo_gen_val = (
        ",".join(TIPOS_GENERACION.values())
        if todos_tipo_gen
        else ",".join(TIPOS_GENERACION[t] for t in tipo_gen_sel)
    )

    central_val = CENTRAL_OPCIONES[central_sel]

    parametros_val = (
        ",".join(PARAMETROS_EXPORTAR.values())
        if todos_parametros
        else ",".join(PARAMETROS_EXPORTAR[p] for p in parametros_sel)
    )

    formato_val = FORMATOS[formato_sel]

    tramos = dividir_en_meses(fecha_inicial, fecha_final)

    if len(tramos) == 1:
        st.info("Consultando el portal COES (1 mes o menos, no requiere partir en tramos)...")
    else:
        st.info(
            f"El rango pedido abarca {len(tramos)} meses. El portal COES solo "
            f"permite descargar 1 mes a la vez, así que se descargarán uno por "
            f"uno (no en paralelo, para evitar choques del lado del servidor), "
            f"convirtiendo cada uno a DataFrame apenas llega, reintentando "
            f"automáticamente los que fallen, y al final se unen todos en un "
            f"solo consolidado."
        )

    progreso = st.progress(0.0)
    estado = st.empty()
    t0 = time.time()

    # Cada mes se descarga uno a la vez, se convierte a DataFrame de inmediato,
    # y se junta a una lista; recién al final se unen todos con pd.concat.
    # Antes esto se hacía con varios hilos en paralelo, pero el servidor de
    # COES arma el reporte en un único archivo fijo en disco (no uno por
    # sesión) y eso hacía que las descargas paralelas chocaran entre sí.
    dataframes, auditorias, errores = _procesar_tramos_secuencial(
        tramos, empresas_val, tipo_gen_val, central_val, parametros_val, formato_val,
        max_reintentos, progreso, estado,
    )

    elapsed = time.time() - t0
    estado.empty()
    progreso.empty()
    st.caption(f"⏱️ Tiempo total: {elapsed:.1f} s ({len(tramos)} tramo(s), secuencial)")


    if errores:
        with st.expander(f"⚠️ {len(errores)} tramo(s) fallaron tras los reintentos", expanded=True):
            for rango_str, msg in errores:
                st.write(f"- **{rango_str}**: {msg}")
            st.warning(
                "Puedes intentar descargar estos tramos individualmente cambiando "
                "el rango de fechas arriba, o volver a presionar 'Descargar' para "
                "reintentar todo el proceso."
            )

    if dataframes:
        df_final = pd.concat(dataframes, ignore_index=True, sort=False)
        es_formato_vertical = isinstance(df_final.columns, pd.MultiIndex)
        df_auditoria = None
        if es_formato_vertical:
            # Formato Vertical (tabla ancha): puede haber quedado más de una
            # fila para la misma Fecha/Hora si algún tramo se reintentó y se
            # sumó dos veces; y conviene ordenar cronológicamente.
            col_fecha_hora = ("Fecha/Hora", "")
            df_final = (
                df_final.drop_duplicates(subset=[col_fecha_hora])
                .sort_values(col_fecha_hora)
                .reset_index(drop=True)
            )
            # Los huecos que deja pd.concat (una central que no generó ese
            # mes, y por lo tanto no tenía esa columna en ese tramo) quedan
            # en 0 en vez de NaN -- Power BI no lee bien los NaN. La
            # auditoría (más abajo) es la que conserva la información de
            # qué central no tuvo dato en qué mes.
            columnas_dato = [c for c in df_final.columns if c != col_fecha_hora]
            df_final[columnas_dato] = df_final[columnas_dato].fillna(0)

            if auditorias:
                df_auditoria = (
                    pd.concat(auditorias, ignore_index=True)
                    .sort_values(["Empresa", "Central", "Tramo"])
                    .reset_index(drop=True)
                )

        if len(dataframes) == len(tramos):
            st.success(
                f"✅ Consolidado completo — {len(dataframes)} de {len(tramos)} tramos "
                f"({len(df_final):,} filas en total)."
            )
        else:
            st.warning(
                f"Consolidado parcial — {len(dataframes)} de {len(tramos)} tramos "
                f"({len(df_final):,} filas en total). Revisa los tramos fallidos arriba."
            )

        # ── Filtro de centrales ──────────────────────────────────────
        # Las opciones son las centrales REALMENTE detectadas en esta
        # descarga (no una lista fija), para no depender de que los nombres
        # coincidan letra por letra con CENTRALES_A_INCLUIR_DEFAULT. La
        # preselección sí parte de esa lista, comparando de forma insensible
        # a mayúsculas/tildes/espacios (ver _clave_normalizada), así que
        # variantes menores como "C.S. EXP. INTIPAMPA" vs "C.S. EXPANSIÓN
        # INTIPAMPA" igual quedan preseleccionadas.
        if es_formato_vertical:
            centrales_detectadas = sorted({
                central for _empresa, central in df_final.columns
                if central not in ("(TOTAL)", "")
            })
        else:
            col_central = next((c for c in df_final.columns if str(c).strip().upper() == "CENTRAL"), None)
            centrales_detectadas = sorted(df_final[col_central].dropna().unique()) if col_central else []

        if centrales_detectadas:
            claves_default = {_clave_normalizada(c) for c in CENTRALES_A_INCLUIR_DEFAULT}
            preseleccion = [c for c in centrales_detectadas if _clave_normalizada(c) in claves_default] or centrales_detectadas

            centrales_sel = st.multiselect(
                "Centrales a incluir en el resultado final",
                options=centrales_detectadas,
                default=preseleccion,
                help="Solo se muestran/exportan las centrales seleccionadas. Por defecto "
                     "vienen preseleccionadas las de CENTRALES_A_INCLUIR_DEFAULT (editable "
                     "en el código); puedes ajustar la selección libremente acá.",
            )

            if centrales_sel and set(centrales_sel) != set(centrales_detectadas):
                if es_formato_vertical:
                    columnas_mantener = [col_fecha_hora] + [
                        c for c in df_final.columns if c != col_fecha_hora and c[1] in centrales_sel
                    ]
                    df_final = df_final[columnas_mantener]
                else:
                    df_final = df_final[df_final[col_central].isin(centrales_sel)].reset_index(drop=True)
                if df_auditoria is not None:
                    df_auditoria = df_auditoria[df_auditoria["Central"].isin(centrales_sel)].reset_index(drop=True)

        if es_formato_vertical:
            # Encabezado plano de una sola fila: "Fecha/Hora" en la primera
            # columna, nombre de la Central en las demás (en vez del
            # encabezado de 2 niveles + fila de índice aparte que generaba
            # pandas por defecto, que quedaba confuso al exportar).
            df_final = _aplanar_encabezado_vertical(df_final, col_fecha_hora)

        with st.expander("👀 Ver columnas detectadas (para verificar que el encabezado se leyó bien)"):
            st.write(list(df_final.columns))

        if df_auditoria is not None:
            centrales_con_hueco = df_auditoria.loc[~df_auditoria["Operó"], ["Empresa", "Central", "Tramo"]]
            with st.expander(
                f"🔍 Auditoría: qué central operó cada mes "
                f"({len(centrales_con_hueco)} hueco(s) detectado(s))",
                expanded=len(centrales_con_hueco) > 0,
            ):
                st.caption(
                    "Los valores en la tabla están en 0 cuando una central no tuvo dato "
                    "ese mes (no NaN, para que Power BI los lea bien). Esta auditoría es "
                    "la forma de distinguir 'generó 0' de 'no hay dato ese mes'."
                )
                st.dataframe(df_auditoria, use_container_width=True)

        st.dataframe(df_final.head(300), use_container_width=True)

        nombre_base = (
            f"MedidoresGeneracion_{fecha_inicial.strftime('%Y%m%d')}_"
            f"{fecha_final.strftime('%Y%m%d')}_consolidado"
        )

        # Descarga xlsx
        buf_xlsx = io.BytesIO()
        with pd.ExcelWriter(buf_xlsx, engine="openpyxl") as writer:
            df_final.to_excel(writer, index=False, sheet_name="MedidoresGeneracion")
            if df_auditoria is not None:
                df_auditoria.to_excel(writer, index=False, sheet_name="Auditoria")
        st.download_button(
            "⬇️ Descargar Excel consolidado",
            data=buf_xlsx.getvalue(),
            file_name=f"{nombre_base}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        # Descarga csv
        buf_csv = df_final.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ Descargar CSV consolidado",
            data=buf_csv,
            file_name=f"{nombre_base}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    else:
        st.error("No se pudo descargar ningún tramo.")
