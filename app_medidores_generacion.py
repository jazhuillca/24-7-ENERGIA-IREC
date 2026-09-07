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
from datetime import date, timedelta

import requests
import streamlit as st

# ---------------------------------------------------------------------------
# Configuración fija
# ---------------------------------------------------------------------------
BASE_URL = "https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion/"
URL_EXPORTAR = BASE_URL + "Exportar"
REFERER = "https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion"

DEFAULT_EMPRESAS = (
    "14260,69,15214,11772,12439,11777,10481,12056,12362,13196,12896,10628,15707,"
    "10420,12708,13165,10901,15571,12584,11095,17,11153,58,19,10684,30,27,40,"
    "11228,23,24,2,11389,14173,11412,206,11544,11058,12097,10552,11429,12884,"
    "15624,15392,10725,11527,10647,14973,11981,11840,11841,12974,10974,11258,"
    "11644,11444,11940,12364,11149,11528,12634,15849,12624,13726,12479,14807,"
    "11185,13,10916,13965,67,108,11217,12480,15259,11102,13120,11218,11063,149,"
    "11323,13966,15009,61,10913,10587,8,138,12758,11103,15080,10,15683,14342,"
    "13984,167,10422,180,11146,15014,4,5,9,76,11563,10636,11129,11509,12096,"
    "11395,13783,18,48,10582,11064,47,11100,10755,12190,11053,10984,11101,11486,"
    "6,11567,10767,11894"
)
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
        empresas = st.text_area(
            "Empresas (IDs separados por coma)",
            value=DEFAULT_EMPRESAS,
            height=100,
        )
        tipos_generacion = st.text_input(
            "Tipos de generación (IDs separados por coma)",
            value=DEFAULT_TIPOS_GENERACION,
        )
        tipos_empresa = st.text_input(
            "Tipos de empresa (IDs separados por coma, puede dejarse vacío)",
            value="",
            help=(
                "El JS del COES también envía un campo 'tiposEmpresa' que proviene "
                "de un multiselect (#cbTipoEmpresa) que no vimos con sus opciones "
                "reales. Si tu exportación falla, revisa este filtro en la página "
                "y copia los valores que use."
            ),
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
                    tipos_empresa=tipos_empresa,
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
