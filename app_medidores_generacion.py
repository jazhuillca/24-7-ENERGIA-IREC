# -*- coding: utf-8 -*-
"""
Streamlit — Descarga de "Medidores de Generación" (COES)

Ejecutar:
    pip install streamlit requests
    streamlit run app_medidores_generacion.py
"""

import re
import unicodedata
from datetime import date, timedelta
from pathlib import Path

import requests
import streamlit as st

# ─────────────────────────────────────────────────────────────
#  CONSTANTES DEL ENDPOINT / FORMULARIO (extraídas del HTML real)
# ─────────────────────────────────────────────────────────────
URL_PAGINA   = "https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion"
URL_EXPORTAR = "https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion/exportar"

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

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": URL_PAGINA,
    "Origin": "https://www.coes.org.pe",
    "X-Requested-With": "XMLHttpRequest",
    "Accept": "application/json, text/javascript, */*; q=0.01",
}


# ─────────────────────────────────────────────────────────────
#  LÓGICA DE DESCARGA (misma que el script standalone)
# ─────────────────────────────────────────────────────────────
def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    s.get(URL_PAGINA, timeout=30)
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
):
    """
    Devuelve (contenido_bytes, nombre_archivo, mensaje_error).
    """
    s = _session()

    payload = {
        "fechaInicial": fecha_inicial,
        "fechaFinal": fecha_final,
        "empresas": empresas,
        "tiposGeneracion": tipos_generacion,
        "central": central,
        "parametros": parametros,
        "tipo": tipo,
    }

    try:
        resp = s.post(URL_EXPORTAR, data=payload, timeout=90)
    except requests.RequestException as e:
        return None, None, f"Error de conexión: {e}"

    if resp.status_code != 200:
        return None, None, f"HTTP {resp.status_code}: {resp.text[:300]}"

    content_type = resp.headers.get("Content-Type", "")
    ext_default = FORMATO_EXTENSION.get(tipo, "xlsx")
    nombre_default = (
        f"MedidoresGeneracion_{fecha_inicial.replace('/', '')}_"
        f"{fecha_final.replace('/', '')}.{ext_default}"
    )

    # Caso A: respuesta binaria directa
    if "json" not in content_type.lower() and "text/html" not in content_type.lower():
        nombre = _nombre_desde_content_disposition(resp, default=nombre_default)
        nombre = unicodedata.normalize("NFKD", nombre).encode("ascii", "ignore").decode()
        return resp.content, nombre, None

    # Caso B: respuesta JSON -> segundo GET
    try:
        data = resp.json()
    except ValueError:
        return None, None, f"Respuesta inesperada (no binaria ni JSON): {resp.text[:300]}"

    url_archivo = data.get("archivo") or data.get("Url") or data.get("url") or data.get("ruta")
    if not url_archivo:
        return None, None, f"JSON sin campo de archivo reconocible: {data}"

    if not url_archivo.startswith("http"):
        url_archivo = "https://www.coes.org.pe" + (
            url_archivo if url_archivo.startswith("/") else "/" + url_archivo
        )

    resp2 = s.get(url_archivo, timeout=90)
    if resp2.status_code != 200:
        return None, None, f"HTTP {resp2.status_code} al descargar archivo final"

    nombre = _nombre_desde_content_disposition(resp2, default=Path(url_archivo).name)
    return resp2.content, nombre, None


# ─────────────────────────────────────────────────────────────
#  INTERFAZ STREAMLIT
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="Medidores de Generación — COES", page_icon="⚡", layout="centered")

st.title("⚡ Medidores de Generación — COES")
st.caption("Descarga directa desde el portal COES (mediciones/medidoresgeneracion)")

with st.form("filtros_form"):

    st.subheader("Rango de fechas")
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

    submitted = st.form_submit_button("📥 Descargar", use_container_width=True)


if submitted:
    if fecha_inicial > fecha_final:
        st.error("La fecha inicial no puede ser posterior a la fecha final.")
        st.stop()

    if not todos_empresas and not empresas_sel:
        st.error("Selecciona al menos una empresa (o marca TODOS los agentes).")
        st.stop()

    # Construir valores del payload a partir de la selección del usuario
    # Nota: el widget real del portal llama a checkAll() al cargar, es decir
    # "TODOS" envía la lista completa de IDs separados por coma (no un código
    # especial tipo -1, a diferencia del endpoint de mantenimientos).
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

    if not todos_tipo_gen and not tipo_gen_sel:
        st.error("Selecciona al menos un tipo de generación (o marca TODOS).")
        st.stop()
    if not todos_parametros and not parametros_sel:
        st.error("Selecciona al menos un parámetro (o marca TODOS).")
        st.stop()

    with st.spinner("Consultando el portal COES..."):
        contenido, nombre, error = descargar_medidores_generacion(
            fecha_inicial=fecha_inicial.strftime("%d/%m/%Y"),
            fecha_final=fecha_final.strftime("%d/%m/%Y"),
            empresas=empresas_val,
            tipos_generacion=tipo_gen_val,
            central=central_val,
            parametros=parametros_val,
            tipo=formato_val,
        )

    if error:
        st.error(f"No se pudo descargar: {error}")
    else:
        st.success(f"Listo — {nombre}")
        st.download_button(
            "⬇️ Guardar archivo",
            data=contenido,
            file_name=nombre,
            use_container_width=True,
        )
