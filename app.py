import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

try:
    import pytesseract
    from pdf2image import convert_from_bytes
    from PIL import Image
    OCR_DISPONIBLE = True
except:
    OCR_DISPONIBLE = False

st.set_page_config(page_title="Extractor DGII Universal", page_icon="🧾", layout="wide")
st.title("🧾 Extractor DGII 100% Universal")
st.caption("Claro, Altice, La Sirena, Ferreterías, Colmados - PDF normal o escaneado")

uploaded_files = st.file_uploader("Arrastra todas tus facturas aquí", type="pdf", accept_multiple_files=True)

def extraer_texto_pdf(file_bytes):
    texto = ""
    try:
        with pdfplumber.open(BytesIO(file_bytes)) as doc:
            texto = "\n".join([p.extract_text() or "" for p in doc.pages])
    except:
        pass
    
    if len(texto.strip()) < 50 and OCR_DISPONIBLE:
        try:
            st.info("Factura escaneada detectada, usando OCR...")
            imagenes = convert_from_bytes(file_bytes, dpi=300)
            for img in imagenes:
                texto += "\n" + pytesseract.image_to_string(img, lang="spa")
        except Exception as e:
            texto += f"\n[OCR error: {e}]"
    return texto

def extraer_datos_universal(texto):
    txt = texto.replace(",", "").replace("$", "").replace("RD", "")
    
    rnc = re.search(r'RNC[:\s-]*([0-9-]{9,13})', txt, re.I)
    ncf = re.search(r'(?:e-NCF|NCF|No\.\s*Comprobante)[:\s]*([EB][0-9A-Z]{10,19})', txt, re.I)
    fecha = re.search(r'(\d{2}[/-]\d{2}[/-]\d{4})', txt)

    # FIX: No confundir SUBTOTAL con TOTAL
    m_itbis = re.search(r'\bITBIS\b[^\d]{0,25}(\d+\.\d{2})', txt, re.I)
    m_sub = re.search(r'\bSUBTOTAL\b[^\d]{0,25}(\d+\.\d{2})', txt, re.I)
    m_total = re.search(r'(?<!SUB)(?:\bTOTAL\s*A\s*PAGAR\b|\bMONTO\s*TOTAL\b|\bTOTAL\s*FACTURA\b|\bTOTAL\b)[^\d]{0,30}(\d+\.\d{2})', txt, re.I)

    itbis = float(m_itbis.group(1)) if m_itbis else 0.0
    subtotal = float(m_sub.group(1)) if m_sub else 0.0
    total = float(m_total.group(1)) if m_total else 0.0

    if total == 0.0:
        montos = re.findall(r'(\d+\.\d{2})', txt)
        if montos:
            montos_f = sorted([float(x) for x in montos])
            total = montos_f[-1]

    if itbis == 0.0 and total > 0:
        if subtotal > 0 and subtotal < total:
            itbis = total - subtotal
        else:
            calc = total - (total / 1.18)
            if calc >= 1:
                itbis = calc
                if subtotal == 0:
                    subtotal = total - itbis

    return {
        "RNC": rnc.group(1).replace("-", "").strip() if rnc else "",
        "NCF": ncf.group(1).strip() if ncf else "",
        "Fecha": fecha.group(1) if fecha else "",
        "Subtotal": f"{subtotal:.2f}" if subtotal else "",
        "ITBIS": f"{itbis:.2f}",
        "Total": f"{total:.2f}",
        "Archivo": ""
    }

if uploaded_files:
    lista = []
    for pdf in uploaded_files:
        bytes_pdf = pdf.read()
        texto = extraer_texto_pdf(bytes_pdf)
        if not texto.strip():
            st.error(f"{pdf.name}: No pude leer texto.")
            continue
        datos = extraer_datos_universal(texto)
        datos["Archivo"] = pdf.name
        lista.append(datos)
        with st.expander(f"Ver texto de {pdf.name}"):
            st.text(texto[:3000])

    if lista:
        df = pd.DataFrame(lista)
        st.success(f"¡Listo! {len(df)} facturas procesadas")
        st.dataframe(df, use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="DGII")
            ws = writer.sheets["DGII"]
            for col_cells in ws.columns:
                max_len = max(len(str(c.value) if c.value else "") for c in col_cells)
                ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 45)
            for col in ["A","B"]:
                for cell in ws[col]:
                    cell.number_format = '@'

        st.download_button("📥 Descargar Excel Universal DGII", output.getvalue(), file_name="reporte_dgii_universal.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown("---")
if not OCR_DISPONIBLE:
    st.warning("Para facturas escaneadas como FOTO instala: pip install pytesseract pdf2image pillow")
else:
    st.success("OCR Activado")