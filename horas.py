import streamlit as st
import pandas as pd
import os
from io import BytesIO

# --- 1. CONFIGURACIÓN ---
st.set_page_config(page_title="Control de Horas y Costes", layout="wide", initial_sidebar_state="expanded")
CARPETA_DATOS = "registros_horas"

if not os.path.exists(CARPETA_DATOS):
    os.makedirs(CARPETA_DATOS)

# --- 2. ESTILOS CSS ---
st.markdown("""
    <style>
    .scroll-container { overflow-x: auto; width: 100%; border: 1px solid #ddd; border-radius: 5px; }
    table { border-collapse: collapse; width: 100%; font-family: sans-serif; }
    th { background-color: #f8f9fa; position: sticky; top: 0; padding: 4px; border: 1px solid #eee; min-width: 32px; text-align: center; font-size: 0.6rem; }
    .col-nombre { min-width: 150px; max-width: 150px; text-align: left; position: sticky; left: 0; background-color: white; z-index: 10; border-right: 2px solid #ddd; padding: 4px 6px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    td { padding: 2px; border: 1px solid #eee; text-align: center; vertical-align: middle; }
    .v-box { border-radius: 2px; padding: 1px 3px; color: white; font-weight: bold; font-size: 0.65rem; display: inline-block; min-width: 25px; }
    .pos { background-color: #2ecc71; } .neg { background-color: #e74c3c; } .neu { background-color: #95a5a6; }
    .punto { color: #eee; font-size: 0.7rem; }
    .info-mini { font-size: 0.55rem; color: #777; display: block; line-height: 1.1; margin-top: 1px; }
    .nombre-txt { font-size: 0.7rem; font-weight: bold; display: block; }
    .importe-total { color: #2c3e50; font-weight: bold; font-size: 0.6rem; border-top: 1px solid #eee; margin-top: 2px; padding-top: 1px; }
    </style>
    """, unsafe_allow_html=True)

# --- 3. FUNCIÓN DE EXCEL CORREGIDA ---
def generar_excel_total(df_completo, precio_hora):
    output = BytesIO()
    # Importante: engine='xlsxwriter' debe estar instalado
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # 1. Filtramos bases únicas
        bases = df_completo['code_base'].dropna().unique()
        
        for base in sorted(bases):
            # Filtrar datos de la base actual
            df_b = df_completo[df_completo['code_base'] == base].copy()
            cols_d = [c for c in df_b.columns if c.startswith('D')]
            
            # Crear lista de resumen para esta base
            resumen_data = []
            for _, fila in df_b.iterrows():
                v_dias = pd.to_numeric(fila[cols_d], errors='coerce').fillna(0)
                excesos = v_dias.apply(lambda x: x - 8 if x > 8 else 0).sum()
                
                resumen_data.append({
                    "Empleado": fila["nombre_operador"],
                    "Base": base,
                    "Total Efectivo": fila.get('ti_totalefectivo', 0),
                    "Horas Extra": excesos,
                    "Importe (€)": round(excesos * precio_hora, 2)
                })
            
            # Convertir a DataFrame y limpiar nombre de pestaña
            df_hoja = pd.DataFrame(resumen_data)
            # Excel no permite: : \ / ? * [ ] en nombres de pestaña y max 31 caracteres
            nombre_limpio = str(base).replace(':','').replace('/','').replace('\\','').replace('*','').replace('?','').replace('[','').replace(']','')[:31]
            if not nombre_limpio: nombre_limpio = f"Base_{base}"
            
            df_hoja.to_excel(writer, index=False, sheet_name=nombre_limpio)
            
    return output.getvalue()

# --- 4. INTERFAZ Y CARGA ---
st.title("⏱️ Control de Horas e Importes")
archivos = [f for f in os.listdir(CARPETA_DATOS) if f.endswith('.csv')]

with st.sidebar:
    st.header("💰 Configuración")
    precio_hora_extra = st.number_input("Precio Hora Extra (€)", value=13.89, step=0.01)
    
    st.divider()
    st.header("📁 Archivos")
    nuevo_archivo = st.file_uploader("Subir CSV", type=["csv"])
    if nuevo_archivo:
        with open(os.path.join(CARPETA_DATOS, nuevo_archivo.name), "wb") as f:
            f.write(nuevo_archivo.getbuffer())
        st.success("Guardado")
        st.rerun()

    if archivos:
        archivo_sel = st.selectbox("📅 Mes", sorted(archivos, reverse=True))
        # Carga del CSV
        df = pd.read_csv(os.path.join(CARPETA_DATOS, archivo_sel), sep=';', decimal=',', encoding='latin-1')
        df.columns = [c.strip() for c in df.columns]
        
        lista_bases = sorted(df['code_base'].dropna().unique())
        base_sel = st.selectbox("📍 Ver Base", lista_bases)
        
        # BOTÓN DE EXCEL
        st.divider()
        try:
            excel_bin = generar_excel_total(df, precio_hora_extra)
            st.download_button(
                label="📥 Descargar Excel Todas las Bases",
                data=excel_bin,
                file_name=f"Resumen_{archivo_sel.replace('.csv', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error generando Excel: {e}")
    else:
        df = None

# --- 5. TABLA VISUAL ---
if df is not None:
    df_f = df[df['code_base'] == base_sel].copy()
    col_dias = [c for c in df.columns if c.startswith('D')]

    html = '<div class="scroll-container"><table><thead><tr><th class="col-nombre">Empleado</th>'
    for d in col_dias: html += f'<th>{d.replace("D", "")[:2]}</th>'
    html += '</tr></thead><tbody>'

    for _, fila in df_f.iterrows():
        v_dias = pd.to_numeric(fila[col_dias], errors='coerce').fillna(0)
        extras = v_dias.apply(lambda x: x - 8 if x > 8 else 0).sum()
        color_bal = "#2ecc71" if extras > 0 else "#95a5a6"

        html += f'<tr><td class="col-nombre">'
        html += f'<span class="nombre-txt">{fila["nombre_operador"]}</span>'
        html += f'<span class="info-mini"><b style="color:{color_bal}">Extras: +{extras:.2f}h</b></span>'
        if extras > 0: html += f'<div class="importe-total">{(extras * precio_hora_extra):.2f} €</div>'
        html += '</td>'

        for i, d in enumerate(col_dias):
            v = v_dias.iloc[i]
            if v > 0:
                exc = v - 8
                clase = "pos" if exc > 0 else "neg" if exc < 0 else "neu"
                html += f'<td><span class="v-box {clase}">{exc:g}</span></td>'
            else: html += '<td><span class="punto">·</span></td>'
        html += '</tr>'
    
    html += '</tbody></table></div>'
    st.write(html, unsafe_allow_html=True)