import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

st.set_page_config(
    page_title="Pharmadvisor | Validador Metodológico",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .stApp { 
        background-color: #2D3346; 
        color: #FFFFFF; 
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .ph-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding-bottom: 15px;
        border-bottom: 1px solid rgba(255, 255, 255, 0.1);
        margin-bottom: 20px;
    }
    .ph-title {
        color: #E6007E;
        font-size: 28px;
        font-weight: bold;
        margin: 0;
    }
    .section-divider {
        margin-top: 40px;
        margin-bottom: 40px;
        border: 0;
        height: 1px;
        background: rgba(255, 255, 255, 0.2);
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="ph-header">
        <div>
            <h1 class="ph-title">Validador Metodológico</h1>
            <span style="color: #9AA5B1; font-size: 13px;">Auditoría Universal de Autorrepetición y Calidad de Registro | Pharmadvisor</span>
        </div>
        <div style="text-align: right;">
            <span style="color: #E6007E; font-weight: bold; font-size: 20px;">Pharm<span style="color: #FFFFFF;">ADVISOR</span></span>
        </div>
    </div>
""", unsafe_allow_html=True)

st.sidebar.subheader("1. Carga del Archivo")
uploaded_file = st.sidebar.file_uploader("Cargar Listado de Visitas (.xlsx)", type=["xlsx"], key="visitas_validador")

@st.cache_data
def cargar_datos_hoja(source, hoja_nombre):
    try:
        df = pd.read_excel(source, sheet_name=hoja_nombre)
        df.columns = df.columns.astype(str).str.strip()
        return df
    except Exception as e:
        return None

df_raw = None
source_file = uploaded_file if uploaded_file is not None else ('listado_visitas_2026-09-07_11-44-08.xlsx' if os.path.exists('listado_visitas_2026-09-07_11-44-08.xlsx') else None)

if source_file is not None:
    try:
        xls = pd.ExcelFile(source_file)
        hoja_seleccionada = st.sidebar.selectbox("Seleccione la Hoja del Excel", options=xls.sheet_names, key="select_hoja_excel")
        df_raw = cargar_datos_hoja(source_file, hoja_seleccionada)
    except Exception as e:
        df_raw = None

df_visitas = None

if df_raw is not None:
    st.sidebar.markdown("---")
    st.sidebar.subheader("2. Mapeo de Columnas (Configuración Universal)")
    st.sidebar.markdown("<span style='color: #9AA5B1; font-size: 11px;'>Asocia los campos de tu archivo con las variables del auditor.</span>", unsafe_allow_html=True)
    
    col_nombres = list(df_raw.columns)
    
    def buscar_defecto(lista, palabras_clave):
        for idx, col in enumerate(lista):
            if any(p in col.lower() for p in palabras_clave):
                return idx
        return 0

    map_rep = st.sidebar.selectbox("Columna: Representante", options=col_nombres, index=buscar_defecto(col_nombres, ['representante', 'asesor', 'ejecutivo', 'usuario']), key="map_rep")
    map_reg = st.sidebar.selectbox("Columna: Región / Zona", options=col_nombres, index=buscar_defecto(col_nombres, ['región', 'region', 'zona', 'coordinación', 'gerencia']), key="map_reg")
    map_lin = st.sidebar.selectbox("Columna: Línea", options=col_nombres, index=buscar_defecto(col_nombres, ['línea', 'linea', 'portafolio', 'ciclo']), key="map_lin")
    map_vis = st.sidebar.selectbox("Columna: ID Visita", options=col_nombres, index=buscar_defecto(col_nombres, ['cod. visita', 'id visita', 'visita', 'id_visita', 'id']), key="map_vis")
    map_fec = st.sidebar.selectbox("Columna: Fecha", options=col_nombres, index=buscar_defecto(col_nombres, ['fecha', 'date', 'dia']), key="map_fec")
    map_med = st.sidebar.selectbox("Columna: Médico / Cliente", options=col_nombres, index=buscar_defecto(col_nombres, ['médicos', 'medico', 'cliente', 'doctor', 'farmacia']), key="map_med")
    map_obj = st.sidebar.selectbox("Columna: Objetivo", options=col_nombres, index=buscar_defecto(col_nombres, ['objetivo', 'propósito', 'meta']), key="map_obj")
    map_com = st.sidebar.selectbox("Columna: Comentario", options=col_nombres, index=buscar_defecto(col_nombres, ['comentario', 'observación', 'observacion', 'detalle', 'feedback']), key="map_com")

    df_visitas = df_raw.rename(columns={
        map_rep: 'Representante',
        map_reg: 'Región',
        map_lin: 'Línea',
        map_vis: 'Cod. visita',
        map_fec: 'Fecha visita',
        map_med: 'Médicos',
        map_obj: 'Objetivo',
        map_com: 'Comentario'
    })

if df_visitas is not None:
    st.sidebar.markdown("---")
    st.sidebar.subheader("3. Filtros en Cascada")
    
    regiones = sorted(df_visitas['Región'].dropna().unique()) if 'Región' in df_visitas.columns else []
    selected_regiones = st.sidebar.multiselect("Región / Coordinación", options=regiones, default=regiones, key="filtro_reg")
    df_f1 = df_visitas[df_visitas['Región'].isin(selected_regiones)] if regiones else df_visitas
    
    lineas = sorted(df_f1['Línea'].dropna().unique()) if 'Línea' in df_f1.columns else []
    selected_lineas = st.sidebar.multiselect("Línea Estratégica", options=lineas, default=lineas, key="filtro_lin")
    df_f2 = df_f1[df_f1['Línea'].isin(selected_lineas)] if lineas else df_f1
    
    representantes = sorted(df_f2['Representante'].dropna().unique()) if 'Representante' in df_f2.columns else []
    selected_reps = st.sidebar.multiselect("Representante", options=representantes, default=representantes, key="filtro_rep")
    df_filtered = df_f2[df_f2['Representante'].isin(selected_reps)] if representantes else df_f2
    
    df_unique = df_filtered.drop_duplicates(subset=['Cod. visita']).copy()
    
    df_unique['Comentario_Clean'] = df_unique['Comentario'].fillna('').astype(str).str.strip().str.lower()
    df_unique['Rep_Comentario_Count'] = df_unique.groupby(['Representante', 'Comentario_Clean'])['Cod. visita'].transform('count')
    
    def check_repetido_individual(row):
        com = row['Comentario_Clean']
        if com in ['', 'nan', 'none', '-']:
            return False
        return row['Rep_Comentario_Count'] > 1

    df_unique['Es_Repetido'] = df_unique.apply(check_repetido_individual, axis=1)
    
    st.subheader("📋 3. Auditoría de Calidad: Índice de Autorrepetición por Representante")
    st.markdown("<span style='color: #9AA5B1;'>Evaluación estricta de cuántas veces cada representante recicla sus propios comentarios entre sus visitas.</span>", unsafe_allow_html=True)
    st.markdown("---")
    
    total_comentarios = len(df_unique)
    comentarios_repetidos = int(df_unique['Es_Repetido'].sum())
    pct_copia = (comentarios_repetidos / total_comentarios * 100) if total_comentarios > 0 else 0
    
    ckpi1, ckpi2, ckpi3 = st.columns(3)
    ckpi1.metric("Visitas Únicas Evaluadas", f"{total_comentarios:,}")
    ckpi2.metric("Comentarios Autorrepetidos", f"{comentarios_repetidos:,}")
    ckpi3.metric("% Índice Global de Autorrepetición", f"{pct_copia:.1f}%")
    
    st.markdown("---")
    
    rep_copia = df_unique.groupby('Representante').agg(
        Total_Visitas=('Cod. visita', 'count'),
        Comentarios_Repetidos=('Es_Repetido', lambda x: int(x.sum()))
    ).reset_index()
    
    rep_copia['Pct_Copia'] = (rep_copia['Comentarios_Repetidos'] / rep_copia['Total_Visitas'] * 100).round(1)
    rep_copia = rep_copia.sort_values(by='Pct_Copia', ascending=True)
    
    altura_grafico = max(450, len(rep_copia) * 25)
    
    fig_bar_copia = px.bar(
        rep_copia, x='Pct_Copia', y='Representante', text='Pct_Copia',
        template='plotly_dark', title="<b>Índice de Autorrepetición (%) por Representante (Copy-Paste Interno)</b>",
        color='Pct_Copia', color_continuous_scale=[[0.0, '#2ECC71'], [0.05, '#2ECC71'], [0.30, '#F39C12'], [0.31, '#E74C3C'], [1.0, '#C0392B']],
        range_color=[0, 100],
        orientation='h'
    )
    fig_bar_copia.update_traces(texttemplate='%{text}%', textposition='outside', textfont_size=11)
    fig_bar_copia.update_layout(
        paper_bgcolor='#1C202C', plot_bgcolor
