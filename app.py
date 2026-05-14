import streamlit as st
from PyPDF2 import PdfReader
import pandas as pd
from io import BytesIO
import re

st.title("PDF a Excel - PRO FINAL")

archivos = st.file_uploader(
    "Subí PDFs",
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
    # ✅ RAZÓN EMISOR
    # =========================

    razon_emisor = ""

    match_rs = re.search(
        r"Razón Social:\s*([A-Z0-9 .]+?)\s*(Condición|Domicilio|CUIT|FACTURA)",
        texto
    )

    if match_rs:
        razon_emisor = match_rs.group(1).strip()

    if not razon_emisor:
        for linea in texto.split("\n"):
            if ("SRL" in linea or "S.A" in linea or "SA" in linea):
                razon_emisor = linea.strip()
                break

    # =========================
    # ✅ RECEPTOR
    # =========================

    razon_receptor = ""

    for linea in texto.split("\n"):
        if cuit_receptor in linea:
            razon_receptor = linea.replace(cuit_receptor, "").strip()
            break

    # =========================
    # ✅ PV + Nº
    # =========================

    m = re.search(
        r"Punto de Venta:\s*Comp\.?\s*Nro:\s*(\d+)\s*(\d+)", texto
    )

    punto_venta = m.group(1) if m else ""
    numero = m.group(2) if m else ""

    # =========================
    # ✅ DETALLE
    # =========================

    if "Código Producto" in texto:
        texto = texto.split("Código Producto", 1)[1]

    lineas = [l.strip() for l in texto.split("\n") if l.strip()]
    buffer_desc = []

    for linea in lineas:

        if "unidades" in linea:

            numeros = re.findall(r"\d+,\d+", linea)

            # ✅ PRIORIDAD 1: PAPUS
            match_x = re.search(r"X\s*(\d+)", linea)

            if match_x:
                numero_raw = match_x.group(1)

                # 🔥 CLAVE: tomar solo últimos 2 dígitos
                cantidad = int(numero_raw[-2:])

                producto = " ".join(buffer_desc).strip()

            else:
                # ✅ IRISITA
                match_cant = re.search(r"(\d+),\d+\s+unidades", linea)

                if match_cant:
                    cantidad = int(match_cant.group(1))
                else:
                    cantidad = 0

                producto = re.split(r"\d+,\d+", linea)[0].strip()

            # =========================
            # ✅ IMPORTES
            # =========================

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

            producto = re.sub(r"\s+", " ", producto).strip()

            if len(producto) < 3:
                producto = "SIN DESCRIPCIÓN"

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

            buffer_desc = []

        else:
            if not any(x in linea for x in ["Código", "Subtotal", "IVA", "CAE", "%"]):
                buffer_desc.append(linea)

    return filas


# =========================
# ✅ GLOBAL
# =========================

if archivos:

    todas = []

    for pdf in archivos:
        todas.extend(procesar_pdf(pdf))

    if todas:
        df = pd.DataFrame(todas)

        st.dataframe(df)

        buffer = BytesIO()
        df.to_excel(buffer, index=False, engine="openpyxl")

        st.download_button(
            "Descargar Excel",
            buffer.getvalue(),
            "facturas_combinadas.xlsx"
        )
``
