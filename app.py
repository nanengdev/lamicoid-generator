import streamlit as st
import pandas as pd
import math
import zipfile
import io
import os
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont

st.set_page_config(page_title='Generador Lamicoides Pro', layout='wide')
st.title('🛠️ Generador de Etiquetas (Centrado Absoluto y Escala Real)')

# --- CONFIGURACIÓN EN SIDEBAR ---
st.sidebar.header('Parámetros de Producción')
dpi = st.sidebar.number_input('DPI (LaserGRBL)', value=600)
sep_mm = st.sidebar.slider('Separación entre etiquetas (mm)', 0.0, 10.0, 3.0)
margen_hoja = st.sidebar.slider('Margen de la plancha (mm)', 0.0, 20.0, 5.0)
cols_num = st.sidebar.number_input('Columnas por fila', min_value=1, value=3)
margen_seg = st.sidebar.slider('Margen seguridad texto (mm)', 0.0, 5.0, 1.5)
tam_fuente_max = st.sidebar.slider('Tamaño de fuente Máximo (pt)', 10, 150, 40)

def mm_a_px(mm): return int(round(mm * dpi / 25.4))
def pts_a_px(pts): return max(1, int(round(pts * dpi / 72.0)))

def cargar_fuente_escalable(size_pt):
    # Intentar cargar fuentes comunes en sistemas Linux (Streamlit Cloud)
    rutas = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSans.ttf"
    ]
    for r in rutas:
        if os.path.exists(r):
            return ImageFont.truetype(r, pts_a_px(size_pt))
    return ImageFont.load_default()

def ajustar_fuente(draw, lineas, w_px, h_px):
    esp_mm = 1.0
    esp_px = mm_a_px(esp_mm)
    for pt in range(tam_fuente_max, 5, -1):
        fnt = cargar_fuente_escalable(pt)
        max_w, total_h = 0, 0
        for l in lineas:
            bbox = draw.textbbox((0,0), l, font=fnt)
            max_w = max(max_w, bbox[2]-bbox[0])
            total_h += (bbox[3]-bbox[1])
        total_h += esp_px * (len(lineas)-1)
        if max_w <= w_px and total_h <= h_px: 
            return fnt, esp_px
    return cargar_fuente_escalable(6), esp_px

def distribuir(etiquetas):
    if not etiquetas: return 0, 0
    cols = min(cols_num, len(etiquetas))
    filas = math.ceil(len(etiquetas) / cols)
    anchos_cols = [0.0] * cols
    altos_filas = [0.0] * filas
    for i, et in enumerate(etiquetas):
        c, f = i % cols, i // cols
        anchos_cols[c] = max(anchos_cols[c], et['ancho_mm'])
        altos_filas[f] = max(altos_filas[f], et['alto_mm'])
    px, py = [], []
    cx, cy = margen_hoja, margen_hoja
    for a in anchos_cols: px.append(cx); cx += a + sep_mm
    for a in altos_filas: py.append(cy); cy += a + sep_mm
    for i, et in enumerate(etiquetas):
        c, f = i % cols, i // cols
        et['x_mm'] = px[c] + (anchos_cols[c] - et['ancho_mm'])/2
        et['y_mm'] = py[f] + (altos_filas[f] - et['alto_mm'])/2
    return cx - sep_mm + margen_hoja, cy - sep_mm + margen_hoja

uploaded = st.file_uploader('Carga tu Excel (.xlsx)', type=['xlsx'])

if uploaded:
    if st.button('🚀 Procesar y Generar ZIP'):
        wb = load_workbook(uploaded, data_only=True)
        zip_buf = io.BytesIO()

        with zipfile.ZipFile(zip_buf, 'w') as zf:
            for sh_name in wb.sheetnames:
                ws = wb[sh_name]
                header = {str(c.value).strip(): i for i, c in enumerate(ws[1], 1) if c.value}
                if 'Texto1' not in header: continue

                etiquetas = []
                for r in range(2, ws.max_row + 1):
                    t1 = str(ws.cell(r, header.get('Texto1')).value or "").strip()
                    t2 = str(ws.cell(r, header.get('Texto2', 99)).value or "").strip()
                    t3 = str(ws.cell(r, header.get('Texto3', 99)).value or "").strip()
                    if not t1 and not t2 and not t3: continue

                    lineas_finales = []
                    for t in [t1, t2, t3]:
                        if t and t != 'None':
                            lineas_finales.extend([l.strip() for l in t.replace('\\\\n', '\\n').split('\\n') if l.strip()])

                    ancho = float(ws.cell(r, header.get('Ancho_mm', 99)).value or 50)
                    alto = float(ws.cell(r, header.get('Alto_mm', 99)).value or 20)
                    cant = int(ws.cell(r, header.get('Cantidad', 99)).value or 1)
                    for _ in range(cant):
                        etiquetas.append({'lineas': lineas_finales, 'ancho_mm': ancho, 'alto_mm': alto})

                if etiquetas:
                    w_h, h_h = distribuir(etiquetas)
                    img = Image.new(\"L\", (mm_a_px(w_h), mm_a_px(h_h)), 255)
                    draw = ImageDraw.Draw(img)

                    for et in etiquetas:
                        x_et_px, y_et_px = mm_a_px(et['x_mm']), mm_a_px(et['y_mm'])
                        w_et_px, h_et_px = mm_a_px(et['ancho_mm']), mm_a_px(et['alto_mm'])
                        m_px = mm_a_px(margen_seg)
                        fnt, esp = ajustar_fuente(draw, et['lineas'], w_et_px - 2*m_px, h_et_px - 2*m_px)

                        info_lineas = []
                        alto_total = 0
                        for l in et['lineas']:
                            bbox = draw.textbbox((0,0), l, font=fnt)
                            wl, hl = bbox[2]-bbox[0], bbox[3]-bbox[1]
                            info_lineas.append({'texto': l, 'w': wl, 'h': hl, 'ox': bbox[0], 'oy': bbox[1]})
                            alto_total += hl
                        alto_total += esp * (len(et['lineas'])-1)

                        y_cursor = y_et_px + (h_et_px - alto_total) / 2
                        for item in info_lineas:
                            x_cursor = x_et_px + (w_et_px - item['w']) / 2
                            draw.text((x_cursor - item['ox'], y_cursor - item['oy']), item['texto'], font=fnt, fill=0)
                            y_cursor += item['h'] + esp

                    svg = [f'<?xml version=\"1.0\"?><svg width=\"{w_h}mm\" height=\"{h_h}mm\" viewBox=\"0 0 {w_h} {h_h}\" xmlns=\"http://www.w3.org/2000/svg\">']
                    svg.append(f'<rect x=\"0\" y=\"0\" width=\"{w_h}\" height=\"{h_h}\" fill=\"none\" stroke=\"none\"/>')
                    for et in etiquetas:
                        svg.append(f'<rect x=\"{et[\"x_mm\"]}\" y=\"{et[\"y_mm\"]}\" width=\"{et[\"ancho_mm\"]}\" height=\"{et[\"alto_mm\"]}\" fill=\"none\" stroke=\"red\" stroke-width=\"0.1\"/>')
                    svg.append('</svg>')

                    png_io = io.BytesIO()
                    img.save(png_io, format='PNG', dpi=(dpi, dpi))
                    zf.writestr(f'{sh_name}/grabado.png', png_io.getvalue())
                    zf.writestr(f'{sh_name}/corte.svg', "\\n".join(svg))

        st.success('✅ Generado con fuentes escalables!')
        st.download_button('🎁 Descargar ZIP Final', zip_buf.getvalue(), 'lamicoides_final.zip')
