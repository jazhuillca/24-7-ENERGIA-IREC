# -*- coding: utf-8 -*-
"""
App Streamlit para consolidar la hoja 'Cmg_Barra' de los reportes de
Costos Marginales Revisados publicados por COES.

Ejecutar con:
    streamlit run app_cmg_barra.py
"""

import io
import re
import requests
import pandas as pd
import streamlit as st
from typing import Optional

st.set_page_config(page_title="Consolidador CMg Barra - COES", layout="wide")

# ---------------------------------------------------------------------------
# Configuración / utilidades
# ---------------------------------------------------------------------------

# El encabezado de la hoja 'Cmg_Barra' está en la fila 3 de Excel
# (fila 3 -> header=2, 0-indexado)
HEADER_ROW = 2

MESES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SETIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}

MESES_CAP = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Setiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}


def generate_url(year: int, month: int, nombre_archivo: Optional[str] = None) -> str:
    """
    Construye la URL de descarga del reporte de Costos Marginales Revisados
    para un año y mes dados.

    Si 'nombre_archivo' no se especifica, se asume el patrón:
        RptCostoMarginal_<Mes>.xlsx   (ej. RptCostoMarginal_Julio.xlsx)
    """
    mes_nombre = MESES[month]
    mes_cap = MESES_CAP[month]
    carpeta_mes = f"{month:02d}_{mes_nombre}"
    archivo_xlsx = nombre_archivo or f"RptCostoMarginal_{mes_cap}.xlsx"

    url = (
        f"https://www.coes.org.pe/portal/browser/download?url="
        f"Operaci%C3%B3n%2FCostos%20Marginales%20CP%2FRevisados%2F{year}%2F"
        f"{carpeta_mes}%2F02_Reportes%20Costos%20Marginales%20CP%2FFinal%2F{archivo_xlsx}"
    )
    return url


def descargar_archivo(year: int, month: int, nombre_archivo: Optional[str] = None):
    """
    Intenta descargar el reporte. Devuelve (bytes|None, url, error|None).

    IMPORTANTE: el portal de COES a veces responde HTTP 200 con una página de
    error (HTML) en vez de un 404 real cuando el archivo aún no existe (por
    ejemplo, meses futuros o reportes aún no publicados). Por eso, además del
    código de estado, se valida que el contenido sea realmente un .xlsx
    (todo archivo Excel real empieza con la firma binaria 'PK').
    """
    url = generate_url(year, month, nombre_archivo)
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        r = requests.get(url, headers=headers, timeout=30)

        if r.status_code != 200 or not r.content:
            return None, url, f"HTTP {r.status_code}"

        if not r.content.startswith(b"PK"):
            return None, url, (
                "El servidor respondió HTTP 200 pero el contenido no es un "
                "Excel válido (probablemente el reporte aún no está publicado "
                "para ese periodo)."
            )

        return r.content, url, None
    except Exception as e:
        return None, url, str(e)


def filtrar_columnas_soles(df: pd.DataFrame) -> pd.DataFrame:
    """
    Estructura real de la hoja 'Cmg_Barra': la celda que dice literalmente
    "S/./MWh" (fila 3) NO es el nombre de una barra ni un simple marcador de
    bloque -- es el encabezado de la columna A, que contiene los valores de
    Fecha/Hora. Es una particularidad de la plantilla de COES: usan esa
    celda como "título" de toda la tabla en vez de escribir "Fecha" ahí.

    A partir de la columna siguiente vienen los nombres de cada barra SIN
    sufijo de moneda (ej. "AGROLMOS 60", "AGUAYTIA 13.8", ...). Más a la
    derecha suele repetirse el mismo patrón con un marcador "USD/MWh" para
    el bloque en dólares.

    Esta función ubica el marcador "S/./MWh" por posición, lo trata como la
    columna de Fecha (renombrándola a "Fecha" para claridad) y toma todas
    las columnas de barras que vienen después, hasta el marcador de dólares
    (o hasta el final si no hay bloque en dólares). Como no depende de una
    letra de columna fija, se adapta automáticamente si el número de barras
    cambia de un mes a otro.
    """
    columnas = list(df.columns)

    def normalizar(texto):
        return re.sub(r"\s+", "", str(texto)).lower()

    idx_marcador_soles = next(
        (i for i, c in enumerate(columnas)
         if re.fullmatch(r"s/\.?/?mwh", normalizar(c))),
        None,
    )

    if idx_marcador_soles is None:
        # No se encontró el marcador "S/./MWh" -> se asume que la primera
        # columna es la de Fecha, como último recurso.
        idx_marcador_soles = 0

    idx_marcador_usd = next(
        (i for i, c in enumerate(columnas)
         if "usd" in normalizar(c) or "us$" in normalizar(c)),
        None,
    )

    col_fecha = columnas[idx_marcador_soles]
    inicio_barras = idx_marcador_soles + 1
    fin_barras = idx_marcador_usd if idx_marcador_usd is not None else len(columnas)

    columnas_barras = columnas[inicio_barras:fin_barras]
    return df[[col_fecha] + columnas_barras]


def leer_crudo(contenido: bytes, hoja: str, filas: int = 10) -> pd.DataFrame:
    """Lee las primeras filas sin encabezado, útil solo para verificación rápida."""
    return pd.read_excel(io.BytesIO(contenido), sheet_name=hoja, header=None, nrows=filas)


def leer_procesado(contenido: bytes, hoja: str) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(contenido), sheet_name=hoja, header=HEADER_ROW)
    df = df.dropna(how="all").reset_index(drop=True)
    # Normaliza nombres de columna: colapsa espacios múltiples y quita
    # espacios al inicio/final, para que la misma barra no se lea como
    # dos columnas distintas entre un mes y otro (ej. "BARRA X" vs "BARRA  X ").
    df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]
    df = filtrar_columnas_soles(df)

    # Fuerza tipos de dato consistentes por columna. Sin esto, si una celda
    # trae texto residual (por ejemplo restos de un encabezado repetido o una
    # nota al pie), la columna queda con tipos mezclados (texto + número) y
    # Streamlit/PyArrow falla al serializarla para mostrarla o exportarla.
    #
    # La primera columna es la de Fecha/Hora por POSICIÓN (aunque su nombre
    # literal sea "S/./MWh", no se renombra por pedido explícito), así que
    # se identifica por posición en vez de buscar "fecha" en el nombre.
    for i, col in enumerate(df.columns):
        if i == 0:
            df[col] = pd.to_datetime(df[col], errors="coerce")
        else:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


# ---------------------------------------------------------------------------
# Estado de sesión
# ---------------------------------------------------------------------------

st.session_state.setdefault("archivos_descargados", {})

# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

st.title("📊 Consolidador de Costos Marginales Revisados — hoja Cmg_Barra")
st.caption("Fuente: COES — Mercado Mayorista / Costos Marginales / Revisados")

hoja = st.text_input("Nombre de la hoja a extraer", value="Cmg_Barra")

st.markdown(
    "Se intenta descargar directamente desde el portal de COES usando el "
    "patrón de nombre `RptCostoMarginal_<Mes>.xlsx`. Si el nombre real del "
    "archivo es distinto, especifícalo manualmente abajo."
)

col1, col2 = st.columns(2)
with col1:
    anios = st.multiselect("Año(s)", options=list(range(2020, 2028)), default=[2026])
with col2:
    meses_sel = st.multiselect(
        "Mes(es)", options=list(MESES.keys()),
        format_func=lambda m: MESES_CAP[m], default=[7],
    )

nombre_personalizado = st.text_input(
    "Nombre de archivo exacto (opcional, aplica a todos los periodos elegidos)",
    placeholder="Ej: RptCostoMarginal_Julio_2026.xlsx",
)

if st.button("Descargar periodos seleccionados"):
    for year in anios:
        for month in meses_sel:
            clave = f"{MESES_CAP[month]} {year}"
            contenido, url, err = descargar_archivo(
                year, month, nombre_personalizado or None
            )
            if err:
                st.error(f"❌ {clave}: {err}\n\nURL intentada: {url}")
            else:
                st.session_state["archivos_descargados"][clave] = contenido
                st.success(f"✅ {clave}: descargado correctamente")

if st.session_state["archivos_descargados"]:
    st.info(
        "Archivos descargados: "
        + ", ".join(st.session_state["archivos_descargados"].keys())
    )
    if st.button("🗑️ Limpiar descargas"):
        st.session_state["archivos_descargados"] = {}
        st.rerun()

# ---------------------------------------------------------------------------
# Vista previa + consolidación
# ---------------------------------------------------------------------------

archivos = st.session_state["archivos_descargados"]

if archivos:
    st.divider()
    st.subheader("🧩 Consolidación")
    st.caption(
        "Se lee el encabezado desde la fila 3 de la hoja y se conservan solo "
        "la columna de Fecha (incluye fecha y hora) y los costos marginales en S/./MWh."
    )

    with st.expander("🔍 Ver vista previa cruda de un archivo (opcional)"):
        nombre_ref = st.selectbox("Archivo de referencia", options=list(archivos.keys()))
        try:
            st.dataframe(leer_crudo(archivos[nombre_ref], hoja), width='stretch')
        except Exception as e:
            st.error(f"No se pudo leer la hoja '{hoja}' en '{nombre_ref}': {e}")

    if st.button("Consolidar todos los archivos disponibles", type="primary"):
        dfs = []
        columnas_vistas = set()
        columnas_por_mes = {}
        col_fecha_nombre = None

        for nombre, contenido in archivos.items():
            try:
                df = leer_procesado(contenido, hoja)

                if col_fecha_nombre is None:
                    col_fecha_nombre = df.columns[0]  # primera columna = Fecha/Hora

                nuevas = [c for c in df.columns if c not in columnas_vistas]
                if columnas_vistas and nuevas:
                    st.info(f"ℹ️ '{nombre}' agrega {len(nuevas)} barra(s) nueva(s): {nuevas}")
                columnas_vistas.update(df.columns)
                columnas_por_mes[nombre] = set(df.columns)

                df["Archivo_Origen"] = nombre
                dfs.append(df)
            except Exception as e:
                st.error(f"Error leyendo '{nombre}': {e}")

        if dfs:
            # Unión de columnas: pd.concat alinea automáticamente por nombre
            # de columna. Si una barra no existe en un mes dado, esa celda
            # queda en NaN para ese mes, en vez de perder la barra o
            # descartarla de los meses donde sí existe.
            consolidado = pd.concat(dfs, ignore_index=True)

            # Avisa si alguna barra del set total no aparece en TODOS los meses
            # (útil para detectar barras que se dieron de baja o entraron a
            # mitad de camino).
            todas_las_barras = columnas_vistas - {"Archivo_Origen", col_fecha_nombre}
            for barra in sorted(todas_las_barras):
                meses_con_barra = [m for m, cols in columnas_por_mes.items() if barra in cols]
                if len(meses_con_barra) < len(columnas_por_mes):
                    meses_sin_barra = [m for m in columnas_por_mes if m not in meses_con_barra]
                    st.caption(f"⚠️ '{barra}' no aparece en: {', '.join(meses_sin_barra)} (quedará NaN ahí).")

            st.success(f"Consolidado generado: {len(consolidado)} filas, {len(consolidado.columns)} columnas.")
            st.dataframe(consolidado.head(100), width='stretch')

            buffer = io.BytesIO()
            consolidado.to_excel(buffer, index=False)
            buffer.seek(0)

            st.download_button(
                "⬇️ Descargar consolidado en Excel",
                data=buffer,
                file_name="CMG_BARRA_Consolidado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.info("Descarga al menos un archivo para comenzar.")
