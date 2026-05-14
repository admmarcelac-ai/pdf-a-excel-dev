import streamlit as st
from PyPDF2 import PdfReader
import pandas as pd
from io import BytesIO
import re

st.title("PDF a Excel")

archivo = st.file_uploader("Subí un PDF", type="pdf")

if archivo:
    reader = PdfReader(archivo)
    texto = reader.pages[0].extract_text()

    if texto:
        # ✅ FECHA
        fecha_match = re.search(r"\d{2}/\d{2}/\d{4}", texto)
        fecha = fecha_match.group(0) if fecha_match else ""

        # ✅ PV + NUMERO (correcto AFIP)
        match = re.search(r"Punto de Venta:\s*Comp\.?\s*Nro:\s*(\d+)\s*(\d+)", texto)

        if match:
            punto_venta = match.group(1)
            numero = match.group(2)
        else:
            punto_venta = ""
            numero = ""

        # ✅ cortar encabezado
        if "Código Producto" in texto:
            texto = texto.split("Código Producto", 1)[1]
        elif "Producto" in texto:
            texto = texto.split("Producto", 1)[1]

        st.text_area("Texto", texto, height=300)

        lineas = [l.strip() for l in texto.split("\n") if l.strip()]

        filas = []
        descripcion = []

        for linea in lineas:

            if "unidades" in linea:

                match_cant = re.search(r"X\s*(\d+),", linea)
                if match_cant:
                    numero_raw = match_cant.group(1)
                    cantidad = int(numero_raw[-2:])
                else:
                    cantidad = 0

                numeros = re.findall(r"\d+,\d+", linea)

                if len(numeros) >= 4:
                    precio = float(numeros[1].replace(",", "."))
                    subtotal = float(numeros[3].replace(",", "."))
                    total = float(numeros[-1].replace(",", "."))
                else:
                    precio = subtotal = total = 0

                producto = " ".join(descripcion).strip()

                if "Subtotal" in producto:
                    producto = producto.split("Subtotal")[-1]

                filas.append({
                    "Fecha": fecha,
                    "Punto de Venta": punto_venta,
                    "Número": numero,
                    "Producto": producto,
                    "Cantidad": cantidad,
                    "Precio Unitario": precio,
                    "Subtotal": subtotal,
                    "Total c/ IVA": total
                })

                descripcion = []

            else:
                if "Código" not in linea and "Subtotal" not in linea:
                    descripcion.append(linea)

        if filas:
            df = pd.DataFrame(filas)
            st.dataframe(df)

            buffer = BytesIO()
            df.to_excel(buffer, index=False, engine="openpyxl")

            st.download_button(
                "Descargar Excel",
                buffer.getvalue(),
                "facturas.xlsx"
            )
        else:
            st.warning("No se detectaron productos.")
