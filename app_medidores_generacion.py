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

import pandas as pd
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


def dividir_en_tramos(fecha_inicio: date, fecha_fin: date, max_dias: int = MAX_DIAS_RANGO):
    """
    Divide [fecha_inicio, fecha_fin] en tramos consecutivos de a lo más
    `max_dias` días cada uno (inclusive), replicando el límite que impone
    el propio formulario del COES por cada exportación.
    """
    tramos = []
    actual = fecha_inicio
    while actual <= fecha_fin:
        fin_tramo = min(actual + timedelta(days=max_dias - 1), fecha_fin)
        tramos.append((actual, fin_tramo))
        actual = fin_tramo + timedelta(days=1)
    return tramos


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


def _normalizar_texto(valor) -> str:
    """
    Limpia texto de celdas de Excel de forma robusta: quita espacios NBSP
    (\\xa0), zero-width spaces (\\u200b) y similares que el sitio del COES
    a veces mete en su HTML/Excel (ya vimos uno en el propio <title> de la
    página: "Medi​dores"), colapsa espacios múltiples, y normaliza a
    mayúsculas sin espacios al borde.
    """
    s = str(valor)
    for invisible in ("\u200b", "\u200c", "\u200d", "\ufeff", "\xa0"):
        s = s.replace(invisible, " " if invisible == "\xa0" else "")
    s = re.sub(r"\s+", " ", s).strip().upper()
    return s


def parsear_a_dataframe(contenido: bytes, tipo: str, hoja: int | str = 0) -> pd.DataFrame:
    """
    Convierte los bytes crudos devueltos por el COES en un DataFrame "limpio".

    Soporta tanto un único exporte por archivo como varios exportes pegados
    uno tras otro dentro de la misma hoja (cada uno con su propio bloque de
    título/encabezado/datos/TOTAL) — esto último pasa si se concatenan
    exportaciones mensuales crudas en un solo archivo.

    Para cada bloque:
    - Ubica automáticamente la fila de encabezado real (la que contiene "FECHA"),
      sin importar cuántas filas de título/metadata haya arriba (normalmente
      empieza en B10, pero esto no depende de esa posición fija). La búsqueda
      normaliza el texto de cada celda (ver `_normalizar_texto`) para tolerar
      espacios invisibles que el COES a veces incluye.
    - Descarta la columna A (vacía) y cualquier columna basura sin nombre real.
    - Se queda solo con las columnas de datos conocidas (FECHA, PUNTO MEDICIÓN,
      EMPRESA, CENTRAL, UNIDAD) más las columnas de intervalos horarios
      (formato "HH:MM"), descartando la columna "TOTAL ENERGIA ACTIVA (MWh)"
      y cualquier otra columna de total.
    - Corta las filas justo antes de las filas de resumen del pie de página
      ("TOTAL ENERGÍA...", "TOTAL POTENCIA...", "Leyenda", etc.), o antes del
      siguiente bloque si lo hay.
    - Si el COES agrega más columnas de intervalo en el futuro, se incluyen
      automáticamente (calzan con el patrón HH:MM); si agrega un campo fijo
      nuevo, hay que sumarlo a `campos_fijos`.

    Todos los bloques encontrados se concatenan en un solo DataFrame.

    Si no se encuentra ninguna fila de encabezado, se levanta un ValueError
    cuyo mensaje incluye una vista previa cruda de las primeras filas/columnas
    del archivo, para poder diagnosticar por qué no calzó la detección.
    """
    if tipo == "3":
        return pd.read_csv(io.BytesIO(contenido))

    raw = pd.read_excel(io.BytesIO(contenido), sheet_name=hoja, header=None, engine="openpyxl")
    raw_norm = raw.map(_normalizar_texto)

    # Puede haber más de un bloque "FECHA...datos...TOTAL" pegado en la misma hoja
    filas_header = raw_norm.index[(raw_norm == "FECHA").any(axis=1)].tolist()
    if not filas_header:
        preview_rows = min(15, len(raw))
        preview_cols = min(6, raw.shape[1])
        preview = raw.iloc[:preview_rows, :preview_cols].to_string()
        raise ValueError(
            "No se encontró ninguna fila de encabezado ('FECHA') en el archivo.\n"
            f"Vista previa cruda (primeras {preview_rows} filas x {preview_cols} columnas):\n"
            f"{preview}"
        )

    campos_fijos = {"FECHA", "PUNTO MEDICIÓN", "EMPRESA", "CENTRAL", "UNIDAD"}
    patron_hora = re.compile(r"^\d{2}:\d{2}$")
    bloques = []

    for idx, fila_header in enumerate(filas_header):
        limite_bloque = filas_header[idx + 1] if idx + 1 < len(filas_header) else len(raw)

        encabezados_norm = raw_norm.iloc[fila_header]
        datos = raw.iloc[fila_header + 1 : limite_bloque].copy()
        datos.columns = encabezados_norm.values  # ya normalizados: "FECHA", "EMPRESA", "00:15", etc.
        datos = datos.loc[:, pd.Series(datos.columns).notna().values & (pd.Series(datos.columns) != "NAN").values]

        col_fecha_cand = [c for c in datos.columns if c == "FECHA"]
        if not col_fecha_cand:
            continue
        col_fecha = col_fecha_cand[0]

        marca = datos[col_fecha].map(_normalizar_texto)
        fin_datos = datos[col_fecha].isna() | marca.str.startswith("TOTAL") | marca.eq("LEYENDA") | marca.eq("NAN") | marca.eq("")
        if fin_datos.any():
            datos = datos.loc[: fin_datos.idxmax() - 1]

        cols_validas = [
            c
            for c in datos.columns
            if c in campos_fijos or patron_hora.match(str(c).strip())
        ]
        if not datos.empty and cols_validas:
            bloques.append(datos[cols_validas])

    if not bloques:
        raise ValueError("Se encontraron encabezados pero ningún bloque tenía datos válidos.")

    return pd.concat(bloques, ignore_index=True)


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
    f"El formulario del COES limita cada exportación a **{MAX_DIAS_RANGO} días**, "
    f"pero esta app **divide automáticamente** rangos más largos (incluso de "
    f"más de un año) en varios tramos y te los entrega juntos. Si eliges "
    f"formato **CSV**, solo puedes seleccionar **un** parámetro a la vez.",
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

    empresas_cargadas = st.session_state["empresas_disponibles"]
    empresas_sel = st.multiselect(
        "Empresa(s)",
        options=[nombre for _id, nombre in empresas_cargadas],
        default=[],
        help="Deja vacío para incluir todas las empresas.",
    )
    nombre_a_id = {nombre: id_ for id_, nombre in empresas_cargadas}
    empresas = ",".join(nombre_a_id[n] for n in empresas_sel)

    parametros_sel = st.multiselect(
        "Parámetro(s) a exportar",
        options=list(PARAMETROS_DISPONIBLES.keys()),
        default=["Potencia Activa (MW)"],
    )

    formato_label = st.selectbox(
        "Formato de salida",
        list(FORMATOS.keys()),
        index=1,  # "Excel Vertical" por defecto
        help=(
            "Para armar una base de datos que crezca mes a mes, usa "
            "'Excel Vertical': cada central/unidad es una FILA, así que "
            "cuando aparezcan centrales nuevas el encabezado no cambia. "
            "'Excel Horizontal' pone cada central como una COLUMNA con "
            "encabezados de varias filas combinadas — no está soportado "
            "por el parser automático de esta app todavía."
        ),
    )
    if formato_label == "Excel Horizontal":
        st.warning(
            "⚠️ El formato Horizontal tiene un encabezado de 4 filas con "
            "celdas combinadas (PUNTO MEDICIÓN / EMPRESA / CENTRAL / UNIDAD) "
            "que el parser automático de esta app aún no interpreta — fallará "
            "al intentar leerlo como tabla. Usa 'Excel Vertical' para que la "
            "descarga funcione de punta a punta."
        )

    with st.expander("Parámetros avanzados"):
        tipos_generacion = st.text_input(
            "Tipos de generación (IDs separados por coma)",
            value=DEFAULT_TIPOS_GENERACION,
        )
        central_label = st.selectbox(
            "Central", list(CENTRAL_OPCIONES.keys()), index=1  # COES por defecto
        )

    enviado = st.form_submit_button("Descargar reporte")

if enviado:
    if fecha_inicial > fecha_final:
        st.error("La fecha inicial no puede ser posterior a la fecha final.")
    elif not parametros_sel:
        st.error("Selecciona al menos un parámetro a exportar.")
    elif FORMATOS[formato_label] == "3" and len(parametros_sel) != 1:
        st.error("Para exportar en CSV solo puedes seleccionar un parámetro.")
    else:
        parametros_str = ",".join(PARAMETROS_DISPONIBLES[p] for p in parametros_sel)
        tipo_val = FORMATOS[formato_label]
        tramos = dividir_en_tramos(fecha_inicial, fecha_final)

        if len(tramos) > 1:
            st.info(
                f"El rango pedido supera los {MAX_DIAS_RANGO} días que permite el COES "
                f"por exportación, así que se descargará en **{len(tramos)} tramos** "
                f"y se armará un solo DataFrame con todo.",
                icon="📦",
            )

        dataframes = []  # DataFrames ya parseados, uno por tramo exitoso
        errores = []  # lista de (rango_str, mensaje)
        progreso = st.progress(0.0)
        estado = st.empty()

        for i, (t_ini, t_fin) in enumerate(tramos, start=1):
            fi_str = t_ini.strftime("%d/%m/%Y")
            ff_str = t_fin.strftime("%d/%m/%Y")
            estado.text(f"Descargando y leyendo tramo {i}/{len(tramos)}: {fi_str} – {ff_str} ...")

            try:
                contenido, content_type_o_error, _nombre = descargar_medidores_generacion(
                    fecha_inicial=fi_str,
                    fecha_final=ff_str,
                    empresas=empresas,
                    tipos_generacion=tipos_generacion,
                    central=CENTRAL_OPCIONES[central_label],
                    parametros=parametros_str,
                    tipo=tipo_val,
                )
            except requests.exceptions.Timeout:
                errores.append((f"{fi_str}–{ff_str}", "Timeout esperando respuesta del COES."))
                progreso.progress(i / len(tramos))
                continue
            except requests.exceptions.RequestException as exc:
                errores.append((f"{fi_str}–{ff_str}", f"Error de conexión: {exc}"))
                progreso.progress(i / len(tramos))
                continue

            if contenido is None:
                errores.append((f"{fi_str}–{ff_str}", content_type_o_error or "Respuesta vacía."))
            else:
                es_zip_valido = contenido[:4] == b"PK\x03\x04"
                if tipo_val != "3" and not es_zip_valido:
                    errores.append(
                        (f"{fi_str}–{ff_str}", "El servidor no devolvió un Excel válido.")
                    )
                else:
                    try:
                        df_tramo = parsear_a_dataframe(contenido, tipo_val)
                        df_tramo["_tramo_desde"] = fi_str
                        df_tramo["_tramo_hasta"] = ff_str
                        dataframes.append(df_tramo)
                    except Exception as exc:
                        errores.append(
                            (f"{fi_str}–{ff_str}", f"No se pudo leer como tabla: {exc}")
                        )

            progreso.progress(i / len(tramos))

        estado.empty()
        progreso.empty()

        if errores:
            with st.expander(f"⚠️ {len(errores)} tramo(s) fallaron", expanded=not dataframes):
                for rango, msg in errores:
                    st.markdown(f"**{rango}**")
                    st.code(msg)

        if dataframes:
            df_final = pd.concat(dataframes, ignore_index=True)

            st.success(
                f"Se combinaron {len(dataframes)} de {len(tramos)} tramo(s) en un "
                f"solo DataFrame: **{len(df_final):,} filas** × {len(df_final.columns)} columnas."
            )

            st.dataframe(df_final, use_container_width=True, height=400)

            col_desc1, col_desc2 = st.columns(2)
            with col_desc1:
                csv_bytes = df_final.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    label="⬇️ Descargar como CSV",
                    data=csv_bytes,
                    file_name=f"medidores_generacion_{fecha_inicial.isoformat()}_{fecha_final.isoformat()}.csv",
                    mime="text/csv",
                )
            with col_desc2:
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                    df_final.to_excel(writer, sheet_name="Medidores", index=False)
                excel_buffer.seek(0)
                st.download_button(
                    label="⬇️ Descargar como Excel",
                    data=excel_buffer,
                    file_name=f"medidores_generacion_{fecha_inicial.isoformat()}_{fecha_final.isoformat()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.error("No se pudo descargar ni leer ningún tramo. Revisa los errores arriba.")


