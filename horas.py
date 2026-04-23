import streamlit as st
import pandas as pd
import os
from io import BytesIO

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Control de Horas y Costes", layout="wide", initial_sidebar_state="expanded")
CARPETA_DATOS = "registros_horas"

if not os.path.exists(CARPETA_DATOS):
    os.makedirs(CARPETA_DATOS)

# --- ESTILOS CSS ---
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

st.title("⏱️ Control de Horas e Importes")

# --- FUNCIONES DE APOYO ---
def to_excel(df_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Resumen_Costes')
    return output.getvalue()

# --- BARRA LATERAL ---
archivos = [f for f in os.listdir(CARPETA_DATOS) if f.endswith('.csv')]

with st.sidebar:
    st.header("💰 Configuración de Costes")
    precio_hora_extra = st.number_input("Precio Hora Extra (€)", value=13.89, step=0.01, format="%.2f")
    
    st.divider()
    st.header("📁 Gestión de Archivos")
    nuevo_archivo = st.file_uploader("Añadir nuevo mes (CSV)", type=["csv"])
    if nuevo_archivo:
        with open(os.path.join(CARPETA_DATOS, nuevo_archivo.name), "wb") as f:
            f.write(nuevo_archivo.getbuffer())
        st.success("Guardado correctamente")
        st.rerun()

    if archivos:
        archivo_sel = st.selectbox("📅 Seleccionar Mes", sorted(archivos, reverse=True))
        df = pd.read_csv(os.path.join(CARPETA_DATOS, archivo_sel), sep=';', decimal=',', encoding='latin-1')
        df.columns = [c.strip() for c in df.columns]
        bases = sorted(df['code_base'].dropna().unique())
        base_sel = st.selectbox("📍 Base", bases)
    else:
        df = None

# --- VISUALIZACIÓN ---
if df is not None:
    try:
        df_filtrado = df[df['code_base'] == base_sel].copy()
        col_dias = [c for c in df.columns if c.startswith('D')]

        st.subheader(f"Vista: {archivo_sel} | Base: {base_sel}")

        # Lista para guardar los datos que irán al Excel
        datos_para_excel = []

        html = '<div class="scroll-container"><table>'
        html += '<thead><tr><th class="col-nombre">Empleado / Totales</th>'
        for dia in col_dias:
            num = dia.replace('D', '')[:2]
            html += f'<th>{num}</th>'
        html += '</tr></thead><tbody>'

        for _, fila in df_filtrado.iterrows():
            valores_dias = pd.to_numeric(fila[col_dias], errors='coerce').fillna(0)
            horas_extra_totales = valores_dias.apply(lambda x: x - 8 if x > 8 else 0).sum()
            
            real_total = float(fila['ti_totalefectivo'])
            dias_activos = (valores_dias > 0).sum()
            previsto_total = dias_activos * 8
            importe_pagar = horas_extra_totales * precio_hora_extra
            
            # Guardar en la lista para Excel
            datos_para_excel.append({
                "Empleado": fila["nombre_operador"],
                "Horas Reales": real_total,
                "Horas Previstas": previsto_total,
                "Horas Extra": horas_extra_totales,
                "Importe Extra (€)": round(importe_pagar, 2)
            })

            # Color y diseño HTML
            color_bal = "#2ecc71" if horas_extra_totales > 0 else "#95a5a6"
            html += f'<tr><td class="col-nombre" title="{fila["nombre_operador"]}">'
            html += f'<span class="nombre-txt">{fila["nombre_operador"]}</span>'
            html += f'<span class="info-mini">Prev:{int(previsto_total)}h | Real:{real_total}h</span>'
            html += f'<span class="info-mini"><b style="color:{color_bal}">Extras: +{horas_extra_totales:.2f}h</b></span>'
            if horas_extra_totales > 0:
                html += f'<div class="importe-total">Total Extra: {importe_pagar:.2f} €</div>'
            html += f'</td>'

            for i, dia in enumerate(col_dias):
                v_dia = valores_dias.iloc[i]
                if v_dia > 0:
                    exceso = v_dia - 8
                    clase = "pos" if exceso > 0 else "neg" if exceso < 0 else "neu"
                    txt = f"+{exceso:g}" if exceso > 0 else f"{exceso:g}"
                    html += f'<td><span class="v-box {clase}">{txt}</span></td>'
                else:
                    html += '<td><span class="punto">·</span></td>'
            html += '</tr>'

        html += '</tbody></table></div>'
        st.write(html, unsafe_allow_html=True)

        # --- BOTÓN DE EXCEL EN LA SIDEBAR (debajo de todo) ---
        df_excel = pd.DataFrame(datos_para_excel)
        excel_data = to_excel(df_excel)
        st.sidebar.divider()
        st.sidebar.download_button(
            label="📥 Descargar Resumen Excel",
            data=excel_data,
            file_name=f"Resumen_{base_sel}_{archivo_sel.replace('.csv', '')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error al procesar los datos: {e}")