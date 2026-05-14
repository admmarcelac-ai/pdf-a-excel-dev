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

        # =========================
        # ✅ DATOS GENERALES
        # =========================

        # Fecha
        fecha_match = re.search(r"\d{2}/\d{2}/\d{4}", texto)
        fecha = fecha_match.group(0) if fecha_match else ""

        # Tipo factura
        tipo_match = re.search(r"FACTURA\s+([ABC])", texto)
        tipo = tipo_match.group(1) if tipo_match else ""

        # CUITs
        cuits = re.findall(r"\d{11}", texto)
        cuit_emisor = cuits[0] if len(cuits) > 0 else ""
        cuit_receptor = cuits[1] if len(cuits) > 1 else ""

        # =========================
        # ✅ RAZONES SOCIALES
        # =========================

        # Emisor
        emisor_match = re.search(r"Razón Social:\s*([A-Z0-9 .]+)", texto)
        razon_emisor = emisor_match.group(1).strip() if emisor_match else ""

        # Receptor (línea que contiene el CUIT receptor)
        razon_receptor = ""
        lineas_texto = texto.split("\n")

        for i, l in enumerate(lineas_texto):
            if cuit_receptor in l and i > 0:
                razon_receptor = lineas_texto[i-1].strip()
                break

        # =========================
        # ✅ PV + NÚMERO
        # =========================

        match = re.search(r"Punto de Venta:\s*Comp\.?\s*Nro:\s*(\d+)\s*(\d+)", texto)

        if match:
            punto_venta = match.group(1)
            numero = match.group(2)
        else:
            punto_venta = ""
            numero = ""

        # =========================
        # ✅ DETALLE PRODUCTOS
        # =========================

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
                    num = match_cant.group(1)
                    cantidad = int(num[-2:])
                else:
                    cantidad = 0

                numeros = re.findall(r"\d+,\d+", linea)

