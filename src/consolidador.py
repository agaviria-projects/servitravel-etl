from config import (
    CARPETA_ENTRADA,
    HOJA_ANIO,
    HOJA_VIATICOS,
    HOJA_PARQUEADEROS
)

from excel_utils import (
    leer_tabla,
    ultima_fila,
    escribir_dataframe,
    obtener_zona,
    construir_dataframe_destino,
    construir_dataframe_viaticos,
    construir_dataframe_parqueaderos,
    COLUMNAS_ANIO,
    COLUMNAS_VIATICOS,
    COLUMNAS_PARQUEADEROS
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

        # ==========================================
        # Determinar hoja origen
        # ==========================================

        nombre_hoja = archivo.stem.upper()

        if nombre_hoja == "METROPOLITANO":
            nombre_hoja = "METRO-SUR"

        # ==========================================
        # Leer archivo origen
        # ==========================================

        df_origen, _ = leer_tabla(
            archivo,
            nombre_hoja,
            COLUMNAS_ANIO
        )

        # Si no pudo leer la hoja, continuar con el siguiente archivo
        if df_origen is None:

            print(f"❌ No fue posible leer la hoja '{nombre_hoja}'.")

            continue

        # ==========================================
        # Construir DataFrame destino
        # ==========================================

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

# ==========================================================
# CONSOLIDAR VIATICOS
# ==========================================================

def consolidar_viaticos(libro):

    print("\n" + "=" * 60)
    print("SERVITRAVEL")
    print("CONSOLIDACIÓN VIATICOS")
    print("=" * 60)

    hoja_destino = libro.sheets[HOJA_VIATICOS]

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

        # ==========================================
        # METROPOLITANO NO TIENE HOJA VIATICOS
        # ==========================================

        if archivo.stem.upper() == "METROPOLITANO":

            print("ℹ VIATICOS ............. No aplica para esta zona.")

            continue

        # ==========================================
        # Leer hoja VIATICOS
        # ==========================================

        df_origen, _ = leer_tabla(
            archivo,
            HOJA_VIATICOS,
            COLUMNAS_VIATICOS
        )

        if df_origen is None:

            print("❌ No fue posible leer la hoja VIATICOS.")

            continue

        # ==========================================
        # Construir DataFrame destino
        # ==========================================

        df_destino = construir_dataframe_viaticos(
            df_origen,
            zona
        )

        # ==========================================
        # Escribir consolidado
        # ==========================================

        fila = ultima_fila(hoja_destino) + 1

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

# ==========================================================
# CONSOLIDAR PARQUEADEROS
# ==========================================================

def consolidar_parqueaderos(libro):

    print("\n" + "=" * 60)
    print("SERVITRAVEL")
    print("CONSOLIDACIÓN PARQUEADEROS")
    print("=" * 60)

    hoja_destino = libro.sheets[HOJA_PARQUEADEROS]

    archivos = sorted(
        archivo
        for archivo in CARPETA_ENTRADA.glob("*.xlsx")
        if not archivo.name.startswith("~$")
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

        # ==========================================
        # METROPOLITANO NO TIENE HOJA PARQUEADEROS
        # ==========================================

        if archivo.stem.upper() == "METROPOLITANO":

            print("ℹ PARQUEADEROS ............. No aplica para esta zona.")

            continue

        # ==========================================
        # Leer hoja PARQUEADEROS
        # ==========================================

        df_origen, _ = leer_tabla(
            archivo,
            HOJA_PARQUEADEROS,
            COLUMNAS_PARQUEADEROS
        )

        if df_origen is None:

            print("❌ No fue posible leer la hoja PARQUEADEROS.")

            continue

        # ==========================================
        # Construir DataFrame destino
        # ==========================================

        df_destino = construir_dataframe_parqueaderos(
            df_origen,
            zona
        )

        # ==========================================
        # Escribir consolidado
        # ==========================================

        fila = ultima_fila(hoja_destino) + 1

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