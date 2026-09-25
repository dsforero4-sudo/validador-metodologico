import streamlit as st
import pandas as pd
import plotly.express as px
import os

# Configuración de página
st.set_page_config(
    page_title="AuditMaster | Auditoría Metodológica de Visitas",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS ejecutivos
st.markdown("""
    <style>
    .stApp { 
        background-color: #1E1E2F; 
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
        color: #00D26A;
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

# Encabezado
st.markdown("""
    <div class="ph-header">
        <div>
            <h1 class="ph-title">AuditMaster | Validador Universal de Visitas</h1>
            <span style="color: #9AA5B1; font-size: 13px;">Auditoría de Calidad Metodológica de Comentarios y Objetivos Comerciales</span>
        </div>
        <div style="text-align: right;">
            <span style="color: #00D26A; font-weight: bold; font-size: 20px;">Sales<span style="color: #FFFFFF;">AUDIT</span></span>
        </div>
    </div>
""", unsafe_allow_html=True)

# --- BARRA LATERAL: CARGA Y MAPEO DE COLUMNAS ---
st.sidebar.subheader("1. Carga del Archivo")
uploaded_file = st.sidebar.file_uploader("Cargar Listado de Visitas (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = st.sidebar.selectbox("Seleccione la Hoja del Excel", options=xls.sheet_names)
        df_raw = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        df_raw.columns = df_raw.columns.astype(str).str.strip()
        
        st.sidebar.markdown("---")
        st.sidebar.subheader("2. Mapeo de Columnas")
        st.sidebar.markdown("<span style='color: #9AA5B1; font-size: 12px;'>Indique qué columna de su archivo corresponde a cada campo de auditoría.</span>", unsafe_allow_html=True)
        
        columnas_disponibles = df_raw.columns.tolist()
        
        # Intentar autodetectar por nombres comunes
        def guess_col(keywords):
            for col in columnas_disponibles:
                if any(kw in col.lower() for kw in keywords):
                    return col
            return columnas_disponibles[0] if columnas_disponibles else None

        col_rep = st.sidebar.selectbox("Columna de Representante", options=columnas_disponibles, index=columnas_disponibles.index(guess_col(['representante', 'rep', 'asesor', 'ejecutivo'])) if guess_col(['representante', 'rep', 'asesor', 'ejecutivo']) in columnas_disponibles else 0)
        col_com = st.sidebar.selectbox("Columna de Comentarios", options=columnas_disponibles, index=columnas_disponibles.index(guess_col(['comentario', 'comentarios', 'observacion', 'detalle'])) if guess_col(['comentario', 'comentarios', 'observacion', 'detalle']) in columnas_disponibles else 0)
        col_obj = st.sidebar.selectbox("Columna de Objetivos", options=columnas_disponibles, index=columnas_disponibles.index(guess_col(['objetivo', 'objetivos', 'meta'])) if guess_col(['objetivo', 'objetivos', 'meta']) in columnas_disponibles else 0)
        col_vis = st.sidebar.selectbox("Columna de ID / Código de Visita", options=columnas_disponibles, index=columnas_disponibles.index(guess_col(['visita', 'id', 'codigo', 'cod'])) if guess_col(['visita', 'id', 'codigo', 'cod']) in columnas_disponibles else 0)
        col_med = st.sidebar.selectbox("Columna de Médico / Cliente", options=columnas_disponibles, index=columnas_disponibles.index(guess_col(['medico', 'médico', 'cliente', 'nombre', 'doctor'])) if guess_col(['medico', 'médico', 'cliente', 'nombre', 'doctor']) in columnas_disponibles else 0)

        # Filtro opcional por representante
        st.sidebar.markdown("---")
        st.sidebar.subheader("3. Filtros")
        reps_disponibles = sorted(df_raw[col_rep].dropna().unique())
        selected_reps = st.sidebar.multiselect("Filtrar por Representante", options=reps_disponibles, default=reps_disponibles)
        
        df_filtered = df_raw[df_raw[col_rep].isin(selected_reps)].copy()
        
        # Procesar visitas únicas basadas en el ID seleccionado
        df_unique = df_filtered.drop_duplicates(subset=[col_vis]).copy()

        # ==========================================
        # SECCIÓN DE AUDITORÍA METODOLÓGICA
        # ==========================================
        st.subheader("🎓 Auditoría de Calidad Metodológica (Comentarios y Objetivos)")
        st.markdown("<span style='color: #9AA5B1;'>Evaluación inteligente de la calidad de registro comercial con justificación teórica adaptativa.</span>", unsafe_allow_html=True)
        st.markdown("---")
        
        df_audit_tec = df_unique.copy()
        df_audit_tec['Com_Text'] = df_audit_tec[col_com].fillna('').astype(str).str.strip()
        df_audit_tec['Obj_Text'] = df_audit_tec[col_obj].fillna('').astype(str).str.strip()
        
        palabras_prohibidas_com = ['', '-', 'nan', 'none', 'nat', '0', 'ok', 'bien', 'excelente', 'sin novedad', 'atendió bien']
        
        def calificar_y_justificar_comentario(txt):
            t_low = txt.lower()
            if t_low in palabras_prohibidas_com or len(txt) < 8:
                return '🔴 Alerta: Vacío o Genérico', 'El comentario está vacío, usa un guion o expresiones genéricas ("bien", "ok", "sin novedad") que no evidencian exploración ni identificación de actitud.'
            
            palabras_alta_calidad = [
                'acepta', 'indiferente', 'objeción', 'objecion', 'escepticismo', 'evasivo', 'acuerdo', 
                'compromiso', 'diferencia', 'valor', 'claro', 'dudas', 'explica', 'explicó', 'habla', 'habló', 
                'revisa', 'revisó', 'conoce', 'conoció', 'prescribe', 'prescribirá'
            ]
            if any(w in t_low for w in palabras_alta_calidad):
                return '🟢 Alta Calidad (Técnica Aplicada)', 'El comentario evidencia de forma sólida la argumentación comercial: presenta la diferencia de la propuesta, aborda consultas y valida la comprensión del cliente.'
            else:
                return '🟡 Regular (Superficial / Sin Actitud Clara)', 'El texto relata que se realizó la visita pero es demasiado superficial; no detalla la argumentación diferencial ni la respuesta del profesional.'

        palabras_actividades = ['entregar', 'visitar', 'saludar', 'llamar', 'dejar', 'muestra', 'material', 'obsequio']
        
        def calificar_y_justificar_objetivo(txt):
            t_low = txt.lower()
            if t_low in palabras_prohibidas_com or len(txt) < 8:
                return '🔴 Alerta: Sin Objetivo Definido', 'El campo de objetivo está vacío o no especifica el comportamiento esperado para la próxima visita.'
            elif any(t_low.startswith(act) for act in palabras_actividades):
                return '🔴 Alerta: Confunde Actividad con Objetivo', 'El texto describe una tarea administrativa o logística ("entregar", "visitar", "dejar muestras") en lugar de definir un comportamiento comercial SMART.'
            elif any(w in t_low for w in ['iniciar', 'reiniciar', 'aumentar', 'sostener', 'mantener', 'reemplazar', 'posicionar', 'evaluar', 'prescripción', 'uso', 'venta', 'compra']):
                return '🟢 Alta Calidad (Comportamental SMART)', 'El objetivo está formulado correctamente como un cambio de comportamiento comercial o de prescripción medible.'
            else:
                return '🟡 Regular (Objetivo Poco Específico)', 'El objetivo menciona una intención hacia la cuenta pero carece de la precisión comportamental requerida.'

        res_com = df_audit_tec['Com_Text'].apply(calificar_y_justificar_comentario)
        df_audit_tec['Calidad_Comentario'] = [r[0] for r in res_com]
        df_audit_tec['Justificacion_Comentario'] = [r[1] for r in res_com]

        res_obj = df_audit_tec['Obj_Text'].apply(calificar_y_justificar_objetivo)
        df_audit_tec['Calidad_Objetivo'] = [r[0] for r in res_obj]
        df_audit_tec['Justificacion_Objetivo'] = [r[1] for r in res_obj]
        
        def calificacion_global(row):
            c = row['Calidad_Comentario']
            o = row['Calidad_Objetivo']
            if 'Alerta' in c or 'Alerta' in o:
                return '🔴 Riesgo Metodológico (Alerta)'
            elif 'Regular' in c or 'Regular' in o:
                return '🟡 En Proceso de Apropiación'
            else:
                return '🟢 Visita Sobresaliente (Metodología Dominada)'

        df_audit_tec['Estado_Metodologico'] = df_audit_tec.apply(calificacion_global, axis=1)
        
        total_v_audit = len(df_audit_tec)
        sobresalientes = len(df_audit_tec[df_audit_tec['Estado_Metodologico'] == '🟢 Visita Sobresaliente (Metodología Dominada)'])
        alertas = len(df_audit_tec[df_audit_tec['Estado_Metodologico'] == '🔴 Riesgo Metodológico (Alerta)'])
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Visitas Únicas Evaluadas", f"{total_v_audit:,}")
        m2.metric("Visitas Metodológicamente Sobresalientes", f"{sobresalientes:,}", f"{(sobresalientes/total_v_audit*100 if total_v_audit > 0 else 0):.1f}%")
        m3.metric("Visitas en Alerta (Riesgo)", f"{alertas:,}", f"{(alertas/total_v_audit*100 if total_v_audit > 0 else 0):.1f}%", delta_color="inverse")
        
        st.markdown("---")
        
        df_rep_metodo = df_audit_tec.groupby([col_rep, 'Estado_Metodologico'], as_index=False).agg(
            Total=(col_vis, 'count')
        )
        
        fig_metodo = px.bar(
            df_rep_metodo, x='Total', y=col_rep, color='Estado_Metodologico', barmode='stack',
            template='plotly_dark', title="<b>Adopción de la Técnica de Ventas por Representante (Visitas Únicas)</b>",
            color_discrete_map={
                '🟢 Visita Sobresaliente (Metodología Dominada)': '#2ECC71',
                '🟡 En Proceso de Apropiación': '#F39C12',
                '🔴 Riesgo Metodológico (Alerta)': '#E74C3C'
            },
            orientation='h'
        )
        fig_metodo.update_layout(
            paper_bgcolor='#1E1E2F', plot_bgcolor='#2D2D44', height=max(450, len(reps_disponibles)*25),
            xaxis_title="Cantidad de Visitas Únicas", yaxis_title="Representante",
            yaxis={'categoryorder': 'total ascending'},
            legend_title="Nivel Metodológico",
            margin=dict(t=50, b=50, l=150, r=40)
        )
        st.plotly_chart(fig_metodo, use_container_width=True)
        
        with st.expander("🔍 Ver detalle completo de auditoría (por Visita Única) con razones y justificaciones teóricas"):
            st.markdown("<span style='color: #9AA5B1; font-size: 13px;'>Esta tabla evalúa cada visita única mostrando el desglose exacto de por qué el comentario y el objetivo obtuvieron su valoración.</span>", unsafe_allow_html=True)
            st.dataframe(
                df_audit_tec[[
                    col_rep, col_vis, col_med, 
                    col_com, 'Calidad_Comentario', 'Justificacion_Comentario', 
                    col_obj, 'Calidad_Objetivo', 'Justificacion_Objetivo', 
                    'Estado_Metodologico'
                ]],
                use_container_width=True, hide_index=True
            )
            
    except Exception as e:
        st.error(f"⚠️ Error al procesar el archivo cargado: {e}")
else:
    st.info("👋 **Por favor carga un archivo de visitas en la barra lateral** para iniciar el análisis metodológico.")
