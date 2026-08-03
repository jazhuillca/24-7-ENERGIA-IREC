"""
Streamlit app: Descarga el reporte de Medidores de Generación del COES.

Página:   https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion
Endpoint: POST https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion/exportar

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
URL_EXPORTAR = "https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion/exportar"
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
DEFAULT_TIPOS_GENERACION = "4,1,3,2"   # ver categorías reales en el <select> del formulario
DEFAULT_PARAMETROS = "1,2,4,3"          # ver checkboxes de parámetros en el formulario

FORMATOS = {"Excel (.xlsx)": "2", "CSV": "1"}  # ajustar según valores reales del <select>


# ---------------------------------------------------------------------------
# Lógica de descarga (adaptada del script original)
# ---------------------------------------------------------------------------
def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "X-Requested-With": "XMLHttpRequest",
            "Referer": REFERER,
            "Origin": "https://www.coes.org.pe",
        }
    )
    s.get(REFERER, timeout=30)
    return s


def descargar_medidores_generacion(
    fecha_inicial: str,
    fecha_final: str,
    empresas: str = DEFAULT_EMPRESAS,
    tipos_generacion: str = DEFAULT_TIPOS_GENERACION,
    central: str = "1",
    parametros: str = DEFAULT_PARAMETROS,
    tipo: str = "2",
    formato: str = "2",
):
    """
    fecha_inicial / fecha_final: strings en formato dd/mm/yyyy.

    Devuelve (contenido_bytes, content_type, nombre_archivo_sugerido) si tuvo éxito,
    o (None, mensaje_error, None) si falló.
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
        "formato": formato,
    }

    resp = s.post(URL_EXPORTAR, data=payload, timeout=180)
    resp.raise_for_status()

    content_type = resp.headers.get("Content-Type", "")
    if "json" in content_type:
        # El servidor devolvió JSON en vez del archivo -> probablemente es un
        # flujo en 2 pasos (generar + descargar), como en mantenimientos.
        mensaje = (
            f"Respuesta JSON inesperada (Content-Type: {content_type!r}):\n\n"
            f"{resp.text[:500]}\n\n"
            "Esto sugiere que 'exportar' solo genera el archivo y falta un "
            "segundo GET para descargarlo, como en el flujo de mantenimientos. "
            "Revisa la pestaña Network del navegador para encontrar esa segunda petición."
        )
        return None, mensaje, None

    # Intentar extraer un nombre de archivo sugerido por el servidor
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

with st.form("form_descarga"):
    col1, col2 = st.columns(2)
    hoy = date.today()
    with col1:
        fecha_inicial = st.date_input(
            "Fecha inicial", value=hoy.replace(day=1) - timedelta(days=1)
        )
    with col2:
        fecha_final = st.date_input("Fecha final", value=hoy)

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
        parametros = st.text_input(
            "Parámetros (IDs separados por coma)",
            value=DEFAULT_PARAMETROS,
        )
        central = st.text_input("Central", value="1")
        tipo = st.text_input("Tipo", value="2")
        formato_label = st.selectbox("Formato de salida", list(FORMATOS.keys()))

    enviado = st.form_submit_button("Descargar reporte")

if enviado:
    if fecha_inicial > fecha_final:
        st.error("La fecha inicial no puede ser posterior a la fecha final.")
    else:
        fi_str = fecha_inicial.strftime("%d/%m/%Y")
        ff_str = fecha_final.strftime("%d/%m/%Y")

        # Inicializamos siempre las 3 variables para evitar NameError
        # si el bloque try falla antes de asignarlas.
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
                    central=central,
                    parametros=parametros,
                    tipo=tipo,
                    formato=FORMATOS[formato_label],
                )
            except requests.exceptions.Timeout:
                error_conexion = (
                    "El portal del COES tardó demasiado en responder (timeout). "
                    "El reporte puede tardar más para rangos de fechas largos; "
                    "intenta un rango más corto o vuelve a intentarlo más tarde."
                )
            except requests.exceptions.RequestException as exc:
                error_conexion = f"Error de conexión con el portal del COES: {exc}"

        if error_conexion:
            st.error(error_conexion)
        elif contenido is None and content_type_o_error is not None:
            # content_type_o_error contiene el mensaje de error (caso JSON inesperado)
            st.warning(content_type_o_error)
        elif contenido is not None:
            st.success(f"Archivo listo: {nombre_archivo}")
            st.download_button(
                label="⬇️ Guardar archivo",
                data=io.BytesIO(contenido),
                file_name=nombre_archivo,
                mime=content_type_o_error or "application/octet-stream",
            )
        else:
            st.error("No se pudo obtener el archivo por un motivo desconocido.")
