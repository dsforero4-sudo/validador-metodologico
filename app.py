import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import io

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
                return '🟢 Alta Calidad (Técnica Aplicada)', 'El comentario evidencia de forma sólida la Fase 2 y es un registro único y personalizado.'
            else:
                return '🟡 Regular (Superficial / Sin Actitud Clara)', 'El texto relata la visita pero carece de profundidad y argumentación diferencial.'

        palabras_actividades = ['entregar', 'visitar', 'saludar', 'llamar', 'dejar', 'muestra', 'material', 'obsequio']
        
        def calificar_y_justificar_objetivo(txt):
            t_low = txt.lower()
            if t_low in palabras_prohibidas_com or len(txt) < 8:
                return '🔴 Alerta: Sin Objetivo Definido', 'El campo de objetivo está vacío o no especifica el comportamiento esperado.'
            elif any(t_low.startswith(act) for act in palabras_actividades):
                return '🔴 Alerta: Confunde Actividad con Objetivo', 'Describe una tarea logística en lieu de definir un comportamiento clínico SMART.'
            elif any(w in t_low for w in ['iniciar', 'reiniciar', 'aumentar', 'sostener', 'mantener', 'reemplazar', 'posicionar', 'evaluar', 'prescripción', 'uso']):
                return '🟢 Alta Calidad (Comportamental SMART)', 'El objetivo está formulado correctamente como un comportamiento prescriptivo.'
            else:
                return '🟡 Regular (Objetivo Poco Específico)', 'El objetivo menciona una intención pero carece de la precisión requerida.'

        res_com = [calificar_y_justificar_comentario_flexible(row['Com_Text'], row['Es_Repetido']) for _, row in df_audit_tec.iterrows()]
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
        m1.metric("Visitas Únicas Evaluadas (Técnica)", f"{total_v_audit:,}")
        m2.metric("Visitas Metodológicamente Sobresalientes", f"{sobresalientes:,}", f"{(sobresalientes/total_v_audit*100):.1f}%")
        m3.metric("Visitas en Alerta (Riesgo)", f"{alertas:,}", f"{(alertas/total_v_audit*100):.1f}%", delta_color="inverse")
        
        st.markdown("---")
        
        df_rep_metodo = df_audit_tec.groupby(['Representante', 'Estado_Metodologico'], as_index=False).agg(
            Total=('Cod. visita', 'count')
        )
        df_rep_totales = df_rep_metodo.groupby('Representante', as_index=False).agg(Total_Rep=('Total', 'sum'))
        df_rep_metodo = pd.merge(df_rep_metodo, df_rep_totales, on='Representante')
        df_rep_metodo['Porcentaje'] = (df_rep_metodo['Total'] / df_rep_metodo['Total_Rep'] * 100).round(1)
        
        def formato_etiqueta(row):
            if row['Porcentaje'] >= 5.0:
                return f"{row['Total']} ({row['Porcentaje']}%)"
            return ""

        df_rep_metodo['Texto_Barra'] = df_rep_metodo.apply(formato_etiqueta, axis=1)
        
        fig_metodo = px.bar(
            df_rep_metodo, x='Total', y='Representante', color='Estado_Metodologico', barmode='stack',
            text='Texto_Barra',
            template='plotly_dark', title="<b>Adopción de la Técnica de Ventas por Representante (Penalización Estricta por Copia)</b>",
            color_discrete_map={
                '🟢 Visita Sobresaliente (Metodología Dominada)': '#2ECC71',
                '🟡 En Proceso de Apropiación': '#F39C12',
                '🔴 Riesgo Metodológico (Alerta)': '#E74C3C'
            },
            orientation='h'
        )
        fig_metodo.update_traces(textposition='inside', insidetextanchor='middle', textfont_size=11)
        fig_metodo.update_layout(
            paper_bgcolor='#1C202C', plot_bgcolor='#2D3346', height=max(450, len(representantes)*25),
            xaxis_title="Cantidad de Visitas Únicas", yaxis_title="Representante",
            yaxis={'categoryorder': 'total ascending'},
            legend_title="Nivel Metodológico",
            margin=dict(t=50, b=50, l=150, r=40)
        )
        st.plotly_chart(fig_metodo, use_container_width=True)
        
        with st.expander("🔍 Ver detalle completo de auditoría (por Visita Única) con filtros de calidad y representante"):
            st.markdown("<span style='color: #9AA5B1; font-size: 13px;'>Filtra el detalle de visitas únicas según el nivel metodológico o el representante de interés.</span>", unsafe_allow_html=True)
            
            col_f_niv, col_f_rep = st.columns(2)
            niveles_disponibles = ['Todos'] + sorted(df_audit_tec['Estado_Metodologico'].unique().tolist())
            with col_f_niv:
                filtro_nivel_sel = st.selectbox("Filtrar por Nivel Metodológico", options=niveles_disponibles, key="select_filtro_nivel")
                
            reps_disponibles_audit = ['Todos'] + sorted(df_audit_tec['Representante'].dropna().unique().tolist())
            with col_f_rep:
                filtro_rep_sel = st.selectbox("Filtrar por Representante", options=reps_disponibles_audit, key="select_filtro_rep_audit")
                
            df_tabla_filtrada = df_audit_tec.copy()
            if filtro_nivel_sel != 'Todos':
                df_tabla_filtrada = df_tabla_filtrada[df_tabla_filtrada['Estado_Metodologico'] == filtro_nivel_sel]
            if filtro_rep_sel != 'Todos':
                df_tabla_filtrada = df_tabla_filtrada[df_tabla_filtrada['Representante'] == filtro_rep_sel]
                
            st.markdown(f"<span style='color: #00D26A; font-size: 13px;'>Mostrando {len(df_tabla_filtrada):,} registros filtrados.</span>", unsafe_allow_html=True)
            
            st.dataframe(
                df_tabla_filtrada[[
                    'Representante', 'Cod. visita', 'Fecha visita', 'Médicos', 
                    'Comentario', 'Calidad_Comentario', 'Justificacion_Comentario', 
                    'Objetivo', 'Calidad_Objetivo', 'Justificacion_Objetivo', 
                    'Estado_Metodologico'
                ]],
                use_container_width=True, hide_index=True
            )

        # ==========================================
        # MÓDULO DE REPORTE EJECUTIVO CON GRÁFICAS EN PDF
        # ==========================================
        st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
        st.subheader("📄 Generación de Reporte Ejecutivo Gerencial con Gráficas")
        st.markdown("<span style='color: #9AA5B1;'>Genera un informe analítico completo incrustando las visualizaciones comerciales para gerentes de distrito y línea.</span>", unsafe_allow_html=True)
        
        if st.button("Generar Informe Ejecutivo con Gráficas"):
            import base64
            
            # Convertir gráficos Plotly a imágenes PNG en Base64 para el HTML
            img_copia_bytes = fig_bar_copia.to_image(format="png", width=900, height=max(450, len(rep_copia) * 22), scale=2)
            img_copia_b64 = base64.b64encode(img_copia_bytes).decode("utf-8")
            
            img_metodo_bytes = fig_metodo.to_image(format="png", width=900, height=max(450, len(representantes) * 22), scale=2)
            img_metodo_b64 = base64.b64encode(img_metodo_bytes).decode("utf-8")
            
            # Análisis gerencial automatizado
            pct_alertas = (alertas / total_v_audit * 100) if total_v_audit > 0 else 0
            pct_sobresalientes = (sobresalientes / total_v_audit * 100) if total_v_audit > 0 else 0
            
            rep_resumen = df_audit_tec.groupby('Representante').agg(
                Visitas=('Cod. visita', 'count'),
                Alertas=('Estado_Metodologico', lambda x: (x == '🔴 Riesgo Metodológico (Alerta)').sum()),
                Copia=('Es_Repetido', 'sum')
            ).reset_index()
            rep_resumen['Pct_Riesgo'] = (rep_resumen['Alertas'] / rep_resumen['Visitas'] * 100).round(1)
            criticos = rep_resumen.sort_values(by='Pct_Riesgo', ascending=False).head(5)
            
            reporte_html = f"""
            <html>
            <head>
                <meta charset="utf-8">
                <style>
                    body {{ font-family: 'Segoe UI', Arial, sans-serif; color: #2C3E50; margin: 30px; line-height: 1.5; }}
                    h1 {{ color: #E6007E; border-bottom: 3px solid #E6007E; padding-bottom: 8px; font-size: 24px; }}
                    h2 {{ color: #2D3346; margin-top: 35px; border-bottom: 1px solid #BDC3C7; padding-bottom: 5px; font-size: 18px; }}
                    .metrics-container {{ display: flex; justify-content: space-between; margin-bottom: 20px; }}
                    .metric-card {{ background: #f8f9fa; border-left: 4px solid #E6007E; padding: 12px; width: 22%; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
                    .metric-title {{ font-size: 11px; color: #7F8C8D; text-transform: uppercase; font-weight: bold; }}
                    .metric-value {{ font-size: 18px; color: #2C3E50; font-weight: bold; margin-top: 5px; }}
                    table {{ width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 25px; }}
                    th, td {{ border: 1px solid #DDDDDD; padding: 10px; text-align: left; font-size: 12px; }}
                    th {{ background-color: #2D3346; color: white; }}
                    tr:nth-child(even) {{ background-color: #f9f9f9; }}
                    .chart-container {{ text-align: center; margin: 25px 0; }}
                    .chart-img {{ max-width: 100%; height: auto; border: 1px solid #ddd; border-radius: 4px; }}
                    .recommendation-box {{ background: #fdf2f7; border-left: 4px solid #E6007E; padding: 15px; border-radius: 4px; margin-top: 20px; }}
                </style>
            </head>
            <body>
                <h1>PHARMADVISOR | INFORME EJECUTIVO DE AUDITORÍA METODOLÓGICA</h1>
                <p><b>Fecha de Emisión:</b> {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')} | <b>Segmento:</b> Visitas a Médicos y Farmacias</p>
                
                <h2>1. Resumen Ejecutivo del Ciclo</h2>
                <div class="metrics-container">
                    <div class="metric-card">
                        <div class="metric-title">Visitas Únicas</div>
                        <div class="metric-value">{total_v_audit:,}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-title">Índice de Clonación</div>
                        <div class="metric-value">{pct_copia:.1f}%</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-title">Visitas Sobresalientes</div>
                        <div class="metric-value">{sobresalientes:,} ({pct_sobresalientes:.1f}%)</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-title">Riesgo / Alertas</div>
                        <div class="metric-value">{alertas:,} ({pct_alertas:.1f}%)</div>
                    </div>
                </div>
                
                <h2>2. Hallazgos Analíticos y Visuales del Ciclo</h2>
                <p>El análisis automatizado permite identificar patrones de autorrepetición en los comentarios y desviaciones en la formulación de objetivos comerciales. A continuación se presentan las visualizaciones de auditoría correspondientes al alcance seleccionado:</p>
                
                <div class="chart-container">
                    <p><b>Figura 1: Índice de Autorrepetición (%) por Representante (Copy-Paste Interno)</b></p>
                    <img class="chart-img" src="data:image/png;base64,{img_copia_b64}" />
                </div>
                
                <div class="chart-container">
                    <p><b>Figura 2: Adopción de la Técnica de Ventas por Representante (Penalización por Copia)</b></p>
                    <img class="chart-img" src="data:image/png;base64,{img_metodo_b64}" />
                </div>
                
                <h2>3. Representantes con Mayor Oportunidad de Acompañamiento (Top Riesgos)</h2>
                <table>
                    <tr>
                        <th>Representante</th>
                        <th>Visitas Evaluadas</th>
                        <th>Registros en Alerta</th>
                        <th>% de Riesgo Metodológico</th>
                    </tr>
            """
            for _, r in criticos.iterrows():
                reporte_html += f"<tr><td>{r['Representante']}</td><td>{r['Visitas']}</td><td>{r['Alertas']}</td><td>{r['Pct_Riesgo']}%</td></tr>"
            
            reporte_html += f"""
                </table>
                
                <div class="recommendation-box">
                    <h3 style="margin-top:0; color: #E6007E;">4. Recomendaciones de Acción para Gerentes de Distrito y Línea</h3>
                    <ol>
                        <li><b>Retroalimentación 1 a 1:</b> Programar sesiones de coaching con los asesores identificados con mayores índices de clonación para fomentar descripciones personalizadas de las objeciones del médico/farmacia.</li>
                        <li><b>Alineación en Objetivos SMART:</b> Reforzar en la planeación del siguiente ciclo que los objetivos redactados reflejen un comportamiento clínico o de prescripción y no tareas logísticas rutinarias.</li>
                        <li><b>Monitoreo Preventivo:</b> Utilizar este informe semanalmente para corregir desvíos antes del cierre oficial de ciclo.</li>
                    </ol>
                </div>
            </body>
            </html>
            """
            
            st.success("¡Informe ejecutivo con gráficas generado exitosamente!")
            st.download_button(
                label="📥 Descargar Informe Ejecutivo Completo (HTML / Imprimible a PDF)",
                data=reporte_html,
                file_name=f"Informe_Gerencial_Graficas_Pharmadvisor_{pd.Timestamp.now().strftime('%Y%m%d')}.html",
                mime="text/html"
            )
            st.info("💡 **Impresión a PDF:** Haz clic en el botón de descarga, abre el archivo HTML en tu navegador web, presiona `Ctrl + P` (o `Cmd + P` en Mac) y selecciona **'Guardar como PDF'**. Las gráficas y tablas se imprimen en alta resolución.")
    else:
        st.warning("⚠️ No se pudieron localizar las columnas 'Comentario' y/o 'Objetivo' en el archivo cargado.")

else:
    st.info("👋 **Por favor carga el archivo de visitas** en la barra lateral para visualizar el validador metodológico.")
