# -*- coding: utf-8 -*-
"""
Descarga el reporte de Medidores de Generación del COES.
Página: https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion
Endpoint (visto en medidores.js, función exportarFormato):
    POST https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion/exportar
"""

import requests

# ---------------------------------------------------------------------------
RUTA_ARCHIVO = r"C:\Users\GZ6710\OneDrive - ENGIE\Escritorio\ENGIE\2026\Plexos\SCRIPT\_tmp_medidores_generacion.xlsx"

URL_EXPORTAR = "https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion/exportar"
REFERER = "https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion"

# Lista de empresas vista en tu payload de ejemplo (ajusta según necesites)
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


def _session():
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": REFERER,
        "Origin": "https://www.coes.org.pe",
    })
    s.get(REFERER, timeout=30)
    return s


def descargar_medidores_generacion(
    fecha_inicial,
    fecha_final,
    empresas=DEFAULT_EMPRESAS,
    tipos_generacion=DEFAULT_TIPOS_GENERACION,
    central="1",
    parametros=DEFAULT_PARAMETROS,
    tipo="2",
    formato="2",
    ruta_salida=RUTA_ARCHIVO,
):
    """
    fecha_inicial / fecha_final: strings en formato dd/mm/yyyy.
    Devuelve la ruta del archivo descargado, o None si falló.
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

    print(f"Descargando medidores de generación {fecha_inicial} - {fecha_final} ...")
    resp = s.post(URL_EXPORTAR, data=payload, timeout=60)
    resp.raise_for_status()

    content_type = resp.headers.get("Content-Type", "")

    if "json" in content_type:
        # El servidor devolvió JSON en vez del archivo -> probablemente
        # es un flujo en 2 pasos (generar + descargar), como en mantenimientos.
        print(f"Respuesta JSON inesperada (Content-Type: {content_type!r}):")
        print(resp.text[:500])
        print(
            "Esto sugiere que 'exportar' solo genera el archivo y falta un "
            "segundo GET para descargarlo, como en el flujo de mantenimientos. "
            "Revisa Network para encontrar esa segunda petición."
        )
        return None

    with open(ruta_salida, "wb") as f:
        f.write(resp.content)

    print(f"Archivo guardado en: {ruta_salida}")
    return ruta_salida


if __name__ == "__main__":
    descargar_medidores_generacion("01/05/2026", "31/05/2026")
