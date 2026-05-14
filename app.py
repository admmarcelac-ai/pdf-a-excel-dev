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
        # ✅ EXTRAER DATOS GENERALES (ANTES DE CORTAR)
        fecha = re.search(r"\d{2}/\d{2}/\d{4}", texto)
        fecha = fecha.group(0) if fecha else ""

        pv = re.search(r"Punto de Venta:\s*(\d+)", texto)
        punto_venta = pv.group(1) if pv else ""

        nro = re.search(r"Comp\.?\s*Nro:\s*\d+\s*(\d+)", texto)
        numero = nro.group(1) if nro else ""

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

            # ✅ detectar líneas reales de producto
            if "unidades" in linea:

                # ✅ cantidad corregida
                match = re.search(r"X\s*(\d+),", linea)
                if match:
                    numero_raw = match.group(1)
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
