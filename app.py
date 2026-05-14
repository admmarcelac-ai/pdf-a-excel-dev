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

            # ✅ detectar líneas de producto reales
            if "unidades" in linea:

                # ✅ cantidad corregida (148 → 48)
                match = re.search(r"X\s*(\d+),", linea)

                if match:
                    numero = match.group(1)
                    cantidad = int(numero[-2:])  # ← CLAVE
                else:
                    cantidad = 0

                # ✅ extraer números
                numeros = re.findall(r"\d+,\d+", linea)

                if len(numeros) >= 4:
                    precio = float(numeros[1].replace(",", "."))
                    subtotal = float(numeros[3].replace(",", "."))
                    total = float(numeros[-1].replace(",", "."))
                else:
                    precio = subtotal = total = 0

                producto = " ".join(descripcion).strip()

                # ✅ limpiar encabezado basura
                if "Subtotal" in producto:
                    producto = producto.split("Subtotal")[-1]

                filas.append({
                    "Producto": producto,
                    "Cantidad": cantidad,
                    "Precio Unitario": precio,
                    "Subtotal": subtotal,
                    "Total c/ IVA": total
                })

                descripcion = []

            else:
                # ✅ evitar encabezado
                if "Código" not in linea and "Subtotal" not in linea:
                    descripcion.append(linea)

        # ✅ salida
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
