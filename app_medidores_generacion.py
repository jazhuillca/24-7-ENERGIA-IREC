"""
Streamlit app: Descarga el reporte de Medidores de Generación del COES.

Página:   https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion
JS real:  /Portal/Areas/Mediciones/Content/Scripts/medidores.js (función exportarFormato)

Hallazgos clave del JS (confirmados leyendo el código fuente real):
- La exportación es un GET (no POST) a:
      /Portal/mediciones/medidoresgeneracion/Exportar?<query string>
  armado con $.param(modelo) y abierto con window.open(enlace, '_blank').
- Campos del query string: fechaInicial, fechaFinal, tiposEmpresa, empresas,
  tiposGeneracion, central, parametros, tipo.
- Validaciones que hace el propio JS antes de exportar (las replicamos aquí):
    * El rango [fechaInicial, fechaFinal] no puede superar 31 días.
    * Si tipo == '3' (CSV), solo se permite seleccionar 1 parámetro.
    * Debe haber al menos 1 parámetro seleccionado.

Ejecutar con:
    streamlit run app_medidores_generacion.py
"""

import io
import re
from datetime import date, timedelta

import requests
import streamlit as st

# ---------------------------------------------------------------------------
# Configuración fija
# ---------------------------------------------------------------------------
BASE_URL = "https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion/"
URL_EXPORTAR = BASE_URL + "Exportar"
REFERER = "https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion"

# Lista real de empresas, extraída directamente del <select id="cbEmpresas">
# de la página (confirmado por inspección en DevTools). Se usa como valor
# por defecto para no depender de la llamada AJAX en vivo, aunque igual
# puedes refrescarla con el botón "Cargar lista de empresas".
EMPRESAS_CONOCIDAS = [
    ("14260", "ACCIONA ENERGIA PERU S.A.C"),
    ("69", "ADINELSA ADN"),
    ("15214", "AGRO INDUSTRIAL PARAMONGA S.A."),
    ("11772", "AGROAURORA S.A.C."),
    ("12439", "AGROINDUSTRIAS SAN JACINTO S.A.A."),
    ("11777", "AGROLMOS SOCIEDAD ANONIMA - AGROLMOS S.A."),
    ("10481", "AGUAS Y ENERGIA PERU"),
    ("12056", "ANDEAN POWER S.A.C."),
    ("12362", "ASOCIACIÓN SANTA LUCIA DE CHACAS"),
    ("13196", "ATRIA ENERGIA S.A.C."),
    ("12896", "BIOENERGIA DEL CHIRA S.A."),
    ("10628", "CASA GRANDE S.A.A."),
    ("15707", "CELARIS ENERGY S.A."),
    ("10420", "CELEPSA"),
    ("12708", "CELEPSA RENOVABLES S.R.L."),
    ("13165", "CENTRALES SANTA ROSA S.A.C."),
    ("10901", "CHINANGO S.A.C."),
    ("15571", "COENERGY S.A.C."),
    ("12584", "COLCA SOLAR S.A.C."),
    ("11480", "COMPAÑIA MINERA ANTAPACCAY S.A."),
    ("11095", "CORPORACION MINERA DEL PERU S.A."),
    ("17", "EGASA"),
    ("11153", "EGEJUNIN"),
    ("58", "EGEMSA"),
    ("19", "EGESUR"),
    ("10684", "ELECTRICA YANAPAMPA SAC"),
    ("30", "ELECTRO ORIENTE"),
    ("27", "ELECTRO SUR ESTE"),
    ("40", "ELECTRO UCAYALI"),
    ("11228", "ELECTRO ZAÑA S.A.C."),
    ("23", "ELECTRONOROESTE S.A."),
    ("24", "ELECTRONORTE S.A."),
    ("2", "ELECTROPERU"),
    ("11389", "EMPRESA DE GENERACION ELECTRICA CANCHAYLLO SAC"),
    ("14173", "EMPRESA DE GENERACION ELECTRICA SANTA ANA S.A.C."),
    ("11412", "EMPRESA DE GENERACION HUALLAGA"),
    ("206", "EMPRESA DE GENERACION HUANZA"),
    ("11544", "EMPRESA ELECTRICA AGUA AZUL"),
    ("11058", "EMPRESA ELECTRICA RIO DOBLE"),
    ("12097", "ENEL GENERACION PIURA S.A."),
    ("10552", "ENERGÍA EÓLICA S.A."),
    ("11429", "ENERGIA RENOVABLE DEL SUR S.A."),
    ("12884", "ENERGIA RENOVABLE LA JOYA S.A."),
    ("15624", "ENGIE ENERGIA PERU S.A.A."),
    ("15392", "EOLICA CARAVELI S.A.C."),
    ("10725", "FENIX POWER PERÚ"),
    ("11527", "GENERACIÓN ANDINA S.A.C."),
    ("10647", "GENERADORA DE ENERGÍA DEL PERÚ"),
    ("14973", "GM OPERACIONES S.A.C."),
    ("11981", "GR CORTARRAMA SOCIEDAD ANONIMA CERRADA"),
    ("11840", "GR PAINO SOCIEDAD ANONIMA CERRADA"),
    ("11841", "GR TARUCA SOCIEDAD ANONIMA CERRADA"),
    ("12974", "GR VALE S.A.C."),
    ("10974", "HIDROCAÑETE S.A."),
    ("11258", "HIDROELECTRICA HUANCHOR S.A.C."),
    ("11644", "HIDROELECTRICA RAPAZ S.A.C."),
    ("11444", "HUAURA POWER GROUP S.A."),
    ("11940", "HYDRO GLOBAL PERÚ S.A.C."),
    ("12364", "HYDRO PATAPO S.A.C."),
    ("11149", "ILLAPU ENERGY"),
    ("11528", "INFRAESTRUCTURAS Y ENERGIAS DEL PERU S.A.C."),
    ("12634", "INLAND ENERGY SAC"),
    ("15849", "INTI JOYA S.A.C."),
    ("12624", "INVERSIONES SHAQSHA S.A.C."),
    ("13726", "JOYA SOLAR S.A.C."),
    ("12479", "KALLPA GENERACION S.A."),
    ("14807", "KONDU SAC"),
    ("11185", "LA VIRGEN"),
    ("13", "LUZ DEL SUR"),
    ("10916", "MAJA ENERGIA S.A.C."),
    ("13965", "MAJES ARCUS S.A.C."),
    ("67", "MINERA CERRO VERDE"),
    ("108", "MINERA CORONA"),
    ("11217", "MOQUEGUA FV S.A.C."),
    ("12480", "ORAZUL ENERGY PERÚ"),
    ("15259", "ORYGEN PERU S.A.A."),
    ("11102", "PANAMERICANA SOLAR SAC."),
    ("13120", "PARQUE EOLICO MARCONA S.A.C."),
    ("11218", "PARQUE EOLICO TRES HERMANAS S.A.C."),
    ("11063", "PETRAMAS"),
    ("149", "PETROPERU"),
    ("11323", "PLANTA DE RESERVA FRIA DE GENERACION DE ETEN S.A."),
    ("10784", "REFINERIA LA PAMPILLA S.A.A"),
    ("13966", "REPARTICIÓN ARCUS S.A.C."),
    ("15009", "SAMAY I S.A.C"),
    ("61", "SAN GABAN"),
    ("10913", "SDE PIURA"),
    ("10587", "SDF ENERGIA SAC"),
    ("8", "SHOUGESA"),
    ("138", "SINERSA"),
    ("12758", "STATKRAFT S.A"),
    ("11103", "TACNA SOLAR SAC."),
    ("15080", "TERMOCHILCA S.A.C."),
    ("10", "TERMOSELVA"),
    ("15683", "TRANSMISION ANDINA DE GENERACION S.A.C."),
    ("11410", "TRUPAL S.A."),
    ("14342", "UNACEM PERU S.A."),
    ("13984", "VARI ENERGIA S.A.C."),
    ("167", "YURA"),
    ("10422", "AGRO INDUSTRIAL PARAMONGA"),
    ("180", "CEMENTO ANDINO"),
    ("11146", "CERRO DEL AGUILA S.A."),
    ("15014", "COGENERACION OQUENDO SAC"),
    ("4", "EDEGEL"),
    ("5", "EEPSA"),
    ("9", "EGENOR"),
    ("76", "ELECTRICA SANTA ROSA"),
    ("11563", "EMPRESA CONCESIONARIA ENERGIA LIMPIA SAC"),
    ("10636", "EMPRESA DE GENERACIÓN ELÉCTRICA CHEVES S.A."),
    ("11129", "EMPRESA DE GENERACION ELECTRICA RIO BAÑOS S.A.C."),
    ("11509", "EMPRESA DE GENERACION ELECTRICA SANTA ANA"),
    ("12096", "ENEL GENERACION PERU S.A.A."),
    ("11395", "ENEL GREEN POWER PERU S.A."),
    ("13783", "ENEL GREEN POWER PERU S.A.C"),
    ("18", "ENERSUR"),
    ("48", "ENGIE"),
    ("10582", "HIDROELECTRICA SANTA CRUZ"),
    ("11064", "HIDROMARAÑON"),
    ("47", "KALLPA GENERACION"),
    ("11100", "MAJES ARCUS"),
    ("10755", "MAPLE ETANOL"),
    ("12190", "ORAZUL ENERGY"),
    ("11053", "PARQUE EOLICO MARCONA S.R.L."),
    ("10984", "PERUANA DE INVERSIONES EN ENERGIAS RENOVABLES S.A."),
    ("11101", "REPARTICIÓN ARCUS"),
    ("11486", "SAMAY I S.A."),
    ("6", "SN POWER"),
    ("11567", "STATKRAFT"),
    ("10767", "TERMOCHILCA"),
    ("11894", "UNION ANDINA DE CEMENTO"),
]

DEFAULT_TIPOS_GENERACION = "4,1,3,2"  # EÓLICA, HIDROELÉCTRICA, SOLAR, TERMOELÉCTRICA

# Parámetros reales vistos en el <select id="cbParametroExportar"> del modal
PARAMETROS_DISPONIBLES = {
    "Potencia Activa (MW)": "1",
    "Potencia Reactiva (MW)": "5",
    "Servicios Auxiliares": "3",
    "Potencia Reactiva Capacitiva (MVAR)": "2",
    "Potencia Reactiva Inductiva (MVAR)": "4",
}

# Valores del <select id="cbCentral">
CENTRAL_OPCIONES = {"TODOS": "0", "COES": "1", "GENERACION RER": "3"}

# Valores de los radio buttons rbFormato
FORMATOS = {
    "Excel Horizontal": "1",
    "Excel Vertical": "2",
    "CSV": "3",
}

MAX_DIAS_RANGO = 31  # límite real que impone el JS del COES


# ---------------------------------------------------------------------------
# Lógica de descarga (replicando exportarFormato del JS real)
# ---------------------------------------------------------------------------
def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Referer": REFERER,
        }
    )
    # Visitamos la página primero para obtener cualquier cookie de sesión necesaria
    s.get(REFERER, timeout=30)
    return s


def obtener_empresas(tipos_empresa: str = ""):
    """
    Replica cargarEmpresas(): POST a .../empresas, que devuelve el HTML
    (checkboxes u <option>) con la lista real de empresas disponibles.

    Devuelve una lista de tuplas (id, nombre) parseadas del HTML, y también
    el HTML crudo por si el parseo no encuentra el patrón esperado.
    """
    s = _session()
    resp = s.post(
        BASE_URL + "empresas",
        data={"tiposEmpresa": tipos_empresa},
        timeout=30,
    )
    resp.raise_for_status()
    html = resp.text

    # Intentamos varios patrones comunes: <option value="X">Nombre</option>
    # y checkboxes tipo <input ... value="X" ... > ... <label ...>Nombre</label>
    empresas = re.findall(r'<option[^>]*value=["\']?(\d+)["\']?[^>]*>([^<]+)</option>', html)

    if not empresas:
        # Fallback: pares value=".." seguidos de un texto de label cercano
        empresas = re.findall(
            r'value=["\']?(\d+)["\']?[^>]*>\s*(?:<[^>]*>)*\s*([^<>{}\n]{2,80})',
            html,
        )

    empresas = [(id_, nombre.strip()) for id_, nombre in empresas if nombre.strip()]
    return empresas, html


def descargar_medidores_generacion(
    fecha_inicial: str,
    fecha_final: str,
    empresas: str,
    tipos_generacion: str,
    parametros: str,
    tipo: str,
    tipos_empresa: str = "",
    central: str = "1",
):
    """
    fecha_inicial / fecha_final: strings en formato dd/mm/yyyy.
    parametros: string con los códigos separados por coma, ej. "1,5,3".
    tipo: "1" (Excel Horizontal), "2" (Excel Vertical) o "3" (CSV).

    Devuelve (contenido_bytes, content_type, nombre_archivo_sugerido) si tuvo éxito,
    o (None, mensaje_error, None) si falló.
    """
    s = _session()

    # El JS arma esto como querystring con $.param() y hace un GET (window.open)
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

    resp = s.get(URL_EXPORTAR, params=params, timeout=180)
    resp.raise_for_status()

    content_type = resp.headers.get("Content-Type", "")
    if "text/html" in content_type or "json" in content_type:
        mensaje = (
            f"Respuesta inesperada (Content-Type: {content_type!r}):\n\n"
            f"{resp.text[:800]}"
        )
        return None, mensaje, None

    disposition = resp.headers.get("Content-Disposition", "")
    nombre_archivo = "medidores_generacion.xlsx"
    if "filename=" in disposition:
        nombre_archivo = disposition.split("filename=")[-1].strip('"; ')

    return resp.content, content_type, nombre_archivo


# ---------------------------------------------------------------------------
# UI de Streamlit
# ---------------------------------------------------------------------------
st.set_page_config(page_title="COES · Medidores de Generación", page_icon="⚡")

st.title("⚡ Medidores de Generación — COES")
st.caption(
    "Descarga el reporte de Medidores de Generación desde el portal del COES "
    "(www.coes.org.pe)."
)

st.info(
    f"El propio formulario del COES limita cada exportación a un rango máximo "
    f"de **{MAX_DIAS_RANGO} días**, y si eliges formato **CSV** solo puedes "
    f"seleccionar **un** parámetro a la vez.",
    icon="ℹ️",
)

col_carga1, col_carga2 = st.columns([1, 3])
with col_carga1:
    cargar = st.button("🔄 Refrescar lista de empresas")
with col_carga2:
    n_empresas = len(st.session_state.get("empresas_disponibles", EMPRESAS_CONOCIDAS))
    st.caption(f"{n_empresas} empresas disponibles (lista precargada; puedes refrescarla).")

if "empresas_disponibles" not in st.session_state:
    st.session_state["empresas_disponibles"] = EMPRESAS_CONOCIDAS

if cargar:
    with st.spinner("Consultando lista de empresas en el COES..."):
        try:
            empresas_lista, html_crudo = obtener_empresas()
            if empresas_lista:
                st.session_state["empresas_disponibles"] = empresas_lista
                st.success(f"Se actualizaron {len(empresas_lista)} empresas.")
            else:
                st.warning(
                    "No se pudo interpretar la lista de empresas devuelta por el "
                    "COES; se mantiene la lista precargada. Revisa el HTML crudo "
                    "abajo si quieres ajustar el parseo."
                )
                with st.expander("Ver HTML crudo de la respuesta"):
                    st.code(html_crudo[:3000])
        except requests.exceptions.RequestException as exc:
            st.error(f"Error al refrescar empresas: {exc}. Se mantiene la lista precargada.")

with st.form("form_descarga"):
    col1, col2 = st.columns(2)
    hoy = date.today()
    with col1:
        fecha_inicial = st.date_input(
            "Fecha inicial", value=hoy.replace(day=1) - timedelta(days=1)
        )
    with col2:
        fecha_final = st.date_input("Fecha final", value=hoy)

    parametros_sel = st.multiselect(
        "Parámetro(s) a exportar",
        options=list(PARAMETROS_DISPONIBLES.keys()),
        default=["Potencia Activa (MW)"],
    )

    formato_label = st.selectbox("Formato de salida", list(FORMATOS.keys()))

    with st.expander("Parámetros avanzados"):
        empresas_cargadas = st.session_state["empresas_disponibles"]
        empresas_sel = st.multiselect(
            "Empresas",
            options=[nombre for _id, nombre in empresas_cargadas],
            default=[],
            help="Vacío = todas las empresas (igual que el comportamiento del formulario original).",
        )
        nombre_a_id = {nombre: id_ for id_, nombre in empresas_cargadas}
        empresas = ",".join(nombre_a_id[n] for n in empresas_sel)

        tipos_generacion = st.text_input(
            "Tipos de generación (IDs separados por coma)",
            value=DEFAULT_TIPOS_GENERACION,
        )
        central_label = st.selectbox(
            "Central", list(CENTRAL_OPCIONES.keys()), index=1  # COES por defecto
        )

    enviado = st.form_submit_button("Descargar reporte")

if enviado:
    dias = (fecha_final - fecha_inicial).days

    if fecha_inicial > fecha_final:
        st.error("La fecha inicial no puede ser posterior a la fecha final.")
    elif dias > MAX_DIAS_RANGO:
        st.error(
            f"El rango de fechas es de {dias} días, y el COES solo permite "
            f"hasta {MAX_DIAS_RANGO} días por exportación. Achica el rango."
        )
    elif not parametros_sel:
        st.error("Selecciona al menos un parámetro a exportar.")
    elif FORMATOS[formato_label] == "3" and len(parametros_sel) != 1:
        st.error("Para exportar en CSV solo puedes seleccionar un parámetro.")
    else:
        fi_str = fecha_inicial.strftime("%d/%m/%Y")
        ff_str = fecha_final.strftime("%d/%m/%Y")
        parametros_str = ",".join(PARAMETROS_DISPONIBLES[p] for p in parametros_sel)

        contenido = None
        content_type_o_error = None
        nombre_archivo = None
        error_conexion = None

        with st.spinner(f"Descargando medidores de generación {fi_str} – {ff_str} ..."):
            try:
                contenido, content_type_o_error, nombre_archivo = descargar_medidores_generacion(
                    fecha_inicial=fi_str,
                    fecha_final=ff_str,
                    empresas=empresas,
                    tipos_generacion=tipos_generacion,
                    central=CENTRAL_OPCIONES[central_label],
                    parametros=parametros_str,
                    tipo=FORMATOS[formato_label],
                )
            except requests.exceptions.Timeout:
                error_conexion = (
                    "El portal del COES tardó demasiado en responder (timeout). "
                    "Intenta un rango más corto o vuelve a intentarlo más tarde."
                )
            except requests.exceptions.RequestException as exc:
                error_conexion = f"Error de conexión con el portal del COES: {exc}"

        if error_conexion:
            st.error(error_conexion)
        elif contenido is None and content_type_o_error is not None:
            st.warning(content_type_o_error)
        elif contenido is not None:
            # Un .xlsx real es un ZIP por dentro: siempre empieza con esta firma.
            es_zip_valido = contenido[:4] == b"PK\x03\x04"

            if FORMATOS[formato_label] != "3" and not es_zip_valido:
                st.error(
                    "El servidor no devolvió un archivo Excel válido "
                    f"(Content-Type: {content_type_o_error!r}, "
                    f"{len(contenido)} bytes). Contenido recibido:"
                )
                try:
                    texto = contenido.decode("utf-8", errors="replace")
                except Exception:
                    texto = repr(contenido[:500])
                st.code(texto[:2000])
            else:
                st.success(f"Archivo listo: {nombre_archivo}")
                mime_por_defecto = (
                    "text/csv"
                    if FORMATOS[formato_label] == "3"
                    else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.download_button(
                    label="⬇️ Guardar archivo",
                    data=io.BytesIO(contenido),
                    file_name=nombre_archivo,
                    mime=content_type_o_error or mime_por_defecto,
                )
        else:
            st.error("No se pudo obtener el archivo por un motivo desconocido.")
