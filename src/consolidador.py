from config import (
    CARPETA_ENTRADA,
    HOJA_ANIO
)

from excel_utils import (
    leer_tabla,
    ultima_fila,
    escribir_dataframe,
    obtener_zona,
    construir_dataframe_destino
)


# ==========================================================
# CONSOLIDAR AÑO 2026
# ==========================================================

def consolidar_anio(libro):

    print("=" * 60)
    print("SERVITRAVEL")
    print("CONSOLIDACIÓN AÑO 2026")
    print("=" * 60)

    hoja_destino = libro.sheets[HOJA_ANIO]

    archivos = sorted(
        CARPETA_ENTRADA.glob("*.xlsx")
    )

    if not archivos:

        print("No existen archivos para procesar.")
        return

    total_registros = 0

    # ------------------------------------------------------
    for archivo in archivos:

        print(f"\n📄 Procesando : {archivo.name}")

        zona = obtener_zona(archivo.name)

        print(f"📍 Zona       : {zona}")

        # Leer archivo origen
        df_origen, _ = leer_tabla(
            archivo
        )

        # Construir DataFrame destino
        df_destino = construir_dataframe_destino(
            df_origen,
            zona
        )
        # Buscar última fila disponible
        fila = ultima_fila(hoja_destino) + 1

        # Escribir información
        escribir_dataframe(
            hoja_destino,
            fila,
            df_destino
        )

        cantidad = len(df_destino)

        total_registros += cantidad

        print(f"✅ Registros agregados : {cantidad}")
        print(f"📊 Total acumulado     : {total_registros}")

    # ------------------------------------------------------

    print("\n" + "=" * 60)
    print(f"TOTAL REGISTROS AGREGADOS : {total_registros}")
    print("=" * 60)