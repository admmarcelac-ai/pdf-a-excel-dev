import streamlit as st
from PyPDF2 import PdfReader
import pandas as pd
from io import BytesIO
import re

st.title("PDF a Excel - Multi Proveedor")

archivos = st.file_uploader(
    "Subí uno o varios PDFs",
    type="pdf",
    accept_multiple_files=True
)

def procesar_pdf(archivo):

    filas = []

    reader = PdfReader(archivo)
    texto = reader.pages[0].extract_text()

    if not texto:
        return filas

    # =========================
    # ✅ DATOS GENERALES
    # =========================

    fecha = re.search(r"\d{2}/\d{2}/\d{4}", texto)
    fecha = fecha.group(0) if fecha else ""

    tipo = re.search(r"FACTURA\s+([ABC])", texto)
    tipo = tipo.group(1) if tipo else ""

    cuits = re.findall(r"\d{11}", texto)
    cuit_emisor = cuits[0] if len(cuits) > 0 else ""
    cuit_receptor = cuits[1] if len(cuits) > 1 else ""

    # =========================
    # ✅ RAZONES SOCIALES
    # =========================

    razon_emisor = ""
    razon_receptor = ""

    for linea in texto.split("\n"):
        if ("SRL" in linea or "S.A" in linea or "SA" in linea):

            linea_limpia = linea.strip()

            if not razon_emisor:
                razon_emisor = linea_limpia

            elif not razon_receptor and linea_limpia != razon_emisor:
                if cuit_receptor in linea_limpia:
                    razon_receptor = linea_limpia.replace(cuit_receptor, "").strip()
                else:
                    razon_receptor = linea_limpia

    # =========================
    # ✅ PV + NUMERO
    # =========================

    match = re.search(r"Punto de Venta:\s*Comp\.?\s*Nro:\s*(\d+)\s*(\d+)", texto)

    punto_venta = match.group(1) if match else ""
    numero = match.group(2) if match else ""

    # =========================
    # ✅ DETALLE (MULTI FORMATO)
    # =========================

    if "Código Producto" in texto:
        texto = texto.split("Código Producto", 1)[1]

    lineas = [l.strip() for l in texto.split("\n") if l.strip()]

    for linea in lineas:

        if "unidades" in linea:

            numeros = re.findall(r"\d+,\d+", linea)

            # ✅ CANTIDAD (sirve para todos los formatos)
            if numeros:
                cantidad = int(float(numeros[0].replace(",", ".")))
            else:
                cantidad = 0

            # ✅ IMPORTES
            if len(numeros) >= 4:
                precio = float(numeros[1].replace(",", "."))
                subtotal = float(numeros[3].replace(",", "."))
                total = float(numeros[-1].replace(",", "."))
            elif len(numeros) >= 3:
                precio = float(numeros[1].replace(",", "."))
                subtotal = float(numeros[2].replace(",", "."))
                total = float(numeros[-1].replace(",", "."))
            else:
                precio = subtotal = total = 0

            # ✅ PRODUCTO (todo antes del primer número)
            producto = re.split(r"\d+,\d+", linea)[0].strip()

            filas.append({
                "Fecha": fecha,
                "Tipo": tipo,
                "CUIT Emisor": cuit_emisor,
                "Razón Emisor": razon_emisor,
                "CUIT Receptor": cuit_receptor,
                "Razón Receptor": razon_receptor,
                "Punto de Venta": punto_venta,
                "Número": numero,
                "Producto": producto,
                "Cantidad": cantidad,
                "Precio Unitario": precio,
                "Subtotal": subtotal,
                "Total c/ IVA": total
            })

    return filas


# =========================
# ✅ PROCESAR TODOS LOS PDFS
# =========================

if archivos:

    todas_filas = []

    for pdf in archivos:
        todas_filas.extend(procesar_pdf(pdf))

    if todas_filas:
        df = pd.DataFrame(todas_filas)

        st.subheader("📊 Resultado combinado")
        st.dataframe(df)

        buffer = BytesIO()
        df.to_excel(buffer, index=False, engine="openpyxl")

        st.download_button(
            "⬇️ Descargar Excel",
            buffer.getvalue(),
            "facturas_combinadas.xlsx"
        )
    else:
        st.warning("No se detectaron datos.")
