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
            <span style="color: #9AA5B1; font-size: 13px;">Auditoría de Autorrepetición y Calidad de Registro (Médicos y Farmacias) | Pharmadvisor</span>
        </div>
        <div style="text-align: right;">
            <span style="color: #E6007E; font-weight: bold; font-size: 20px;">Pharm<span style="color: #FFFFFF;">ADVISOR</span></span>
        </div>
    </div>
""", unsafe_allow_html=True)

st.sidebar.subheader("1. Carga del Archivo")
uploaded_file = st.sidebar.file_uploader("Cargar Listado de Visitas (.xlsx)", type=["xlsx"], key="visitas_validador")

# Función limpia sin widgets adentro
@st.cache_data
def cargar_datos_hoja(source, hoja_nombre):
    try:
        df = pd.read_excel(source, sheet_name=hoja_nombre)
        df.columns = df.columns.astype(str).str.strip()
        return df
    except Exception as e:
        return None

df_visitas = None
source_file = uploaded_file if uploaded_file is not None else ('listado_visitas_2026-09-07_11-44-08.xlsx' if os.path.exists('listado_visitas_2026-09-07_11-44-08.xlsx') else None)

if source_file is not None:
    try:
        xls = pd.ExcelFile(source_file)
        hoja_seleccionada = st.sidebar.selectbox("Seleccione la Hoja del Excel", options=xls.sheet_names, key="select_hoja_excel")
        df_visitas = cargar_datos_hoja(source_file, hoja_seleccionada)
    except Exception as e:
        df_visitas = None

if df_visitas is not None:
    st.sidebar.markdown("---")
    st.sidebar.subheader("2. Filtros en Cascada")
    
    regiones = sorted(df_visitas['Región'].dropna().unique()) if 'Región' in df_visitas.columns else []
    selected_regiones = st.sidebar.multiselect("Región / Coordinación", options=regiones, default=regiones, key="filtro_reg")
    df_f1 = df_visitas[df_visitas['Región'].isin(selected_regiones)] if regiones else df_visitas
    
    ciclos = sorted(df_f1['Ciclo'].dropna().unique()) if 'Ciclo' in df_f1.columns else []
    selected_ciclos = st.sidebar.multiselect("Ciclo", options=ciclos, default=ciclos, key="filtro_ciclo")
    df_f2 = df_f1[df_f1['Ciclo'].isin(selected_ciclos)] if ciclos else df_f1
    
    lineas = sorted(df_f2['Línea'].dropna().unique()) if 'Línea' in df_f2.columns else []
    selected_lineas = st.sidebar.multiselect("Línea Estratégica", options=lineas, default=lineas, key="filtro_lin")
    df_f3 = df_f2[df_f2['Línea'].isin(selected_lineas)] if lineas else df_f2
    
    representantes = sorted(df_f3['Representante'].dropna().unique()) if 'Representante' in df_f3.columns else []
    selected_reps = st.sidebar.multiselect("Representante", options=representantes, default=representantes, key="filtro_rep")
    df_filtered = df_f3[df_f3['Representante'].isin(selected_reps)] if representantes else df_f3
    
    df_unique = df_filtered.drop_duplicates(subset=['Cod. visita']).copy()
    
    df_unique['Comentario_Clean'] = df_unique['Comentario'].fillna('').astype(str).str.strip().str.lower()
    df_unique['Rep_Comentario_Count'] = df_unique.groupby(['Representante', 'Comentario_Clean'])['Cod. visita'].transform('count')
    
    def check_repetido_individual(row):
        com = row['Comentario_Clean']
        if com in ['', 'nan', 'none', '-']:
            return False
        return row['Rep_Comentario_Count'] > 1

    df_unique['Es_Repetido'] = df_unique.apply(check_repetido_individual, axis=1)
    
    # ==========================================
    # GRÁFICA 3: ÍNDICE DE COPIADO POR REPRESENTANTE
    # ==========================================
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
        paper_bgcolor='#1C202C', plot_bgcolor='#2D3346', height=altura_grafico,
        xaxis_title="Índice de Autorrepetición (%)", yaxis_title="Representante",
        xaxis=dict(range=[0, 115]),
        yaxis={'categoryorder': 'total ascending'},
        margin=dict(t=50, b=50, l=150, r=20)
    )
    st.plotly_chart(fig_bar_copia, use_container_width=True)
    
    with st.expander("Ver listado de visitas únicas con comentarios repetidos"):
        st.dataframe(df_unique[df_unique['Es_Repetido'] == True][['Región', 'Representante', 'Fecha visita', 'Médicos', 'Comentario']].head(50), use_container_width=True, hide_index=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ==========================================
    # GRÁFICA 4: AUDITORÍA DE CALIDAD METODOLÓGICA
    # ==========================================
    st.subheader("🎓 4. Auditoría de Calidad Metodológica (Técnica de Ventas Pharmadvisor - Visitas Únicas)")
    st.markdown("<span style='color: #9AA5B1;'>Evaluación inteligente de la Fase 2 (Comentarios) y Fase 1/3 (Objetivos) con penalización automática a Alerta ante registros con copy-paste.</span>", unsafe_allow_html=True)
    st.markdown("---")
    
    if 'Comentario' in df_filtered.columns and 'Objetivo' in df_filtered.columns:
        df_audit_tec = df_unique.copy()
        
        df_audit_tec['Com_Text'] = df_audit_tec['Comentario'].fillna('').astype(str).str.strip()
        df_audit_tec['Obj_Text'] = df_audit_tec['Objetivo'].fillna('').astype(str).str.strip()
        
        palabras_prohibidas_com = ['', '-', 'nan', 'none', 'nat', '0', 'ok', 'bien', 'excelente', 'sin novedad', 'atendió bien']
        
        def calificar_y_justificar_comentario_flexible(txt, es_rep):
            t_low = txt.lower()
            if t_low in palabras_prohibidas_com or len(txt) < 8:
                return '🔴 Alerta: Vacío o Genérico', 'El comentario está vacío o usa expresiones genéricas ("bien", "ok", "sin novedad").'
            
            if es_rep:
                return '🔴 Alerta: Penalizado por Copy-Paste (Clonación)', 'El texto contiene elementos teóricos correctos, pero al estar repetido idénticamente en múltiples visitas, se invalida por falta de exploración individual genuina.'
            
            palabras_alta_calidad = [
                'acepta', 'indiferente', 'objeción', 'objecion', 'escepticismo', 'evasivo', 'acuerdo', 
                'compromiso', 'diferencia', 'valor', 'claro', 'dudas', 'explica', 'explicó', 'habla', 'habló', 
                'revisa', 'revisó', 'conoce', 'conoció', 'prescribe', 'prescribirá'
            ]
            if any(w in t_low for w in palabras_alta_calidad):
