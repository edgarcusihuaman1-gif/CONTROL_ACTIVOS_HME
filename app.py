import streamlit as st
import pandas as pd
import altair as alt
import io
import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Configuración de la página
st.set_page_config(
    page_title="Sistema de Gestión - HME Drive Thru",
    page_icon="🚗",
    layout="wide"
)

# Estilo visual adaptado
st.markdown("""
    <style>
        .stMetric {
            background-color: rgba(128, 128, 128, 0.1);
            padding: 15px;
            border-radius: 10px;
            border: 1px solid rgba(128, 128, 128, 0.2);
        }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "DRIVE TRHU BASE.xlsx"

def cargar_datos(file_path):
    if not os.path.exists(file_path):
        st.error(f"No se encuentra el archivo {file_path}. Asegúrate de que esté en la carpeta del proyecto.")
        return None, None, None, None
        
    xls = pd.ExcelFile(file_path)
    df_hme = pd.read_excel(xls, 'Drive HME')
    df_at_2025 = pd.read_excel(xls, 'Atenciones 2025')
    df_at_2026 = pd.read_excel(xls, 'Atenciones 2026')
    df_cotizaciones = pd.read_excel(xls, 'Cotizaciones')
    
    # Formatear fechas si existen en las hojas
    for df in [df_at_2025, df_at_2026]:
        if 'FECHA DE SOLICITUD' in df.columns:
            df['FECHA DE SOLICITUD'] = pd.to_datetime(df['FECHA DE SOLICITUD'], errors='coerce').dt.strftime('%Y-%m-%d')
            
    return df_hme, df_at_2025, df_at_2026, df_cotizaciones

# Inicializar estado de autenticación y datos
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

if 'df_hme' not in st.session_state:
    df_h, df_25, df_26, df_cot = cargar_datos(EXCEL_FILE)
    st.session_state.df_hme = df_h
    st.session_state.df_at_2025 = df_25
    st.session_state.df_at_2026 = df_26
    st.session_state.df_cotizaciones = df_cot

# --- PANTALLA DE LOGIN ---
if not st.session_state.autenticado:
    st.title("🔐 Acceso al Sistema - HME Drive Thru")
    st.markdown("---")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("### Por favor, ingrese sus credenciales")
        with st.form("form_login"):
            usuario_ingresado = st.text_input("Usuario")
            clave_ingresada = st.text_input("Contraseña", type="password")
            btn_login = st.form_submit_button("Ingresar al Sistema", use_container_width=True)
            
            if btn_login:
                USUARIO_VALIDO = "admin"
                CLAVE_VALIDA = "hme2026"
                
                if usuario_ingresado == USUARIO_VALIDO and clave_ingresada == CLAVE_VALIDA:
                    st.session_state.autenticado = True
                    st.success("¡Acceso exitoso! Cargando sistema...")
                    st.rerun()
                else:
                    st.error("Usuario o contraseña incorrectos. Intente nuevamente.")
                    
    st.stop()

# --- FUNCIONES DE EXPORTACIÓN (PDF / EXCEL) ---
def convertir_df_a_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Reporte')
    return output.getvalue()

def generar_pdf_resumen(titulo, df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor('#1f77b4'), spaceAfter=15
    )
    
    elements.append(Paragraph(f"<b>Reporte Ejecutivo: {titulo}</b>", title_style))
    elements.append(Spacer(1, 10))
    
    if df.empty:
        elements.append(Paragraph("No hay registros para los filtros seleccionados.", styles['Normal']))
    else:
        cols_a_mostrar = list(df.columns[:6])
        data_tabla = [cols_a_mostrar]
        for _, row in df.head(30).iterrows():
            data_tabla.append([str(row[c])[:25] for c in cols_a_mostrar])
            
        t = Table(data_tabla, colWidths=[80, 80, 90, 90, 85, 85])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f77b4')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f8f9fa')),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#d6d6d6')),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
        ]))
        elements.append(t)
        
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()


# --- APLICACIÓN PRINCIPAL (POST-LOGIN) ---
if 'df_hme' in st.session_state:
    st.title("📊 Sistema de Gestión y Control - HME Drive Thru")
    
    st.sidebar.header("👤 Sesión de Usuario")
    st.sidebar.success("Conectado como: **admin**")
    
    # Botón para recargar datos directamente desde el archivo Excel modificado
    if st.sidebar.button("🔄 Sincronizar / Recargar Excel", use_container_width=True):
        df_h, df_25, df_26, df_cot = cargar_datos(EXCEL_FILE)
        st.session_state.df_hme = df_h
        st.session_state.df_at_2025 = df_25
        st.session_state.df_at_2026 = df_26
        st.session_state.df_cotizaciones = df_cot
        st.success("¡Datos sincronizados correctamente desde el Excel!")
        st.rerun()

    if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

    st.sidebar.markdown("---")
    st.sidebar.header("🔍 Panel de Navegación")
    menu_opcion = st.sidebar.radio("Ir a la sección:", [
        "🏢 Estado Drive HME", 
        "🛠️ Atenciones 2025", 
        "🛠️ Atenciones 2026", 
        "💰 Cotizaciones",
        "➕ Registrar Nuevo Caso (2026)"
    ])
    
    st.sidebar.markdown("---")
    st.sidebar.info("💡 Edita tu Excel en la PC y haz clic en 'Sincronizar' para ver los cambios al instante.")

    # Pestaña 1: Estado Drive HME
    if menu_opcion == "🏢 Estado Drive HME":
        c1, c2 = st.columns(2)
        with c1:
            marcas = ["Todas"] + list(st.session_state.df_hme["MARCA"].unique())
            filtro_marca = st.selectbox("Filtrar por Marca:", marcas)
        with c2:
            busqueda_tienda = st.text_input("Buscar Tienda / Código / Ubicación:", "")

        df_hme_filtered = st.session_state.df_hme.copy()
        if filtro_marca != "Todas":
            df_hme_filtered = df_hme_filtered[df_hme_filtered["MARCA"] == filtro_marca]
        if busqueda_tienda:
            df_hme_filtered = df_hme_filtered[
                df_hme_filtered["TIENDA"].astype(str).str.contains(busqueda_tienda, case=False, na=False) |
                df_hme_filtered["UBICACIÓN"].astype(str).str.contains(busqueda_tienda, case=False, na=False)
            ]

        col_title, col_btn1, col_btn2 = st.columns([3, 1.2, 1.2])
        with col_title:
            st.subheader("🏢 Estado General de Equipos en Tiendas HME")
        with col_btn1:
            st.download_button("📊 Excel", convertir_df_a_excel(df_hme_filtered), "Estado_Drive_HME_Filtrado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_btn2:
            st.download_button("📄 PDF", generar_pdf_resumen("Estado Drive HME", df_hme_filtered), "Estado_Drive_HME_Filtrado.pdf", "application/pdf", use_container_width=True)
        
        st.markdown("---")
        
        total_tiendas = len(df_hme_filtered)
        total_headphone_op = df_hme_filtered["HEADPHONE OPERATIVOS"].sum() if not df_hme_filtered.empty else 0
        total_headphone_av = df_hme_filtered["HEADPHONE AVERIADOS"].sum() if not df_hme_filtered.empty else 0
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Tiendas Monitoreadas", total_tiendas)
        col2.metric("Headphones Operativos", int(total_headphone_op))
        col3.metric("Headphones Averiados", int(total_headphone_av))
        
        st.markdown("---")

        if not df_hme_filtered.empty:
            st.markdown("### 📊 Comparativa de Headphones por Tienda")
            df_chart = df_hme_filtered[["TIENDA", "HEADPHONE OPERATIVOS", "HEADPHONE AVERIADOS"]].melt(
                id_vars=["TIENDA"], value_vars=["HEADPHONE OPERATIVOS", "HEADPHONE AVERIADOS"],
                var_name="Estado", value_name="Cantidad"
            )
            chart_hme = alt.Chart(df_chart).mark_bar().encode(
                x=alt.X('TIENDA:N', sort=None),
                y='Cantidad:Q',
                color=alt.Color('Estado:N', scale=alt.Scale(range=['#2ca02c', '#d62728'])),
                tooltip=['TIENDA', 'Estado', 'Cantidad']
            ).properties(height=300)
            st.altair_chart(chart_hme, use_container_width=True)

        st.dataframe(df_hme_filtered, use_container_width=True)

    # Pestaña 2: Atenciones 2025
    elif menu_opcion == "🛠️ Atenciones 2025":
        df_25_filtered = st.session_state.df_at_2025.copy()
        
        c1, c2 = st.columns(2)
        with c1:
            if "TIPOLOGIA" in df_25_filtered.columns:
                tipologias_25 = ["Todas"] + list(df_25_filtered["TIPOLOGIA"].unique())
                f_tipo_25 = st.selectbox("Filtrar por Tipología (2025):", tipologias_25)
                if f_tipo_25 != "Todas":
                    df_25_filtered = df_25_filtered[df_25_filtered["TIPOLOGIA"] == f_tipo_25]
        with c2:
            busqueda_25 = st.text_input("Buscar en Atenciones 2025:", "")
            if busqueda_25:
                df_25_filtered = df_25_filtered[df_25_filtered.astype(str).apply(lambda x: x.str.contains(busqueda_25, case=False)).any(axis=1)]

        col_title, col_btn1, col_btn2 = st.columns([3, 1.2, 1.2])
        with col_title:
            st.subheader("🛠️ Registro de Atenciones - 2025")
        with col_btn1:
            st.download_button("📊 Excel", convertir_df_a_excel(df_25_filtered), "Atenciones_2025_Filtrado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_btn2:
            st.download_button("📄 PDF", generar_pdf_resumen("Atenciones 2025", df_25_filtered), "Atenciones_2025_Filtrado.pdf", "application/pdf", use_container_width=True)
        
        st.markdown("---")
        st.metric("Total de Atenciones Filtradas (2025)", len(df_25_filtered))
        st.markdown("---")
        st.dataframe(df_25_filtered, use_container_width=True)

    # Pestaña 3: Atenciones 2026
    elif menu_opcion == "🛠️ Atenciones 2026":
        df_26_filtered = st.session_state.df_at_2026.copy()
        
        c1, c2 = st.columns(2)
        with c1:
            if "TIPOLOGIA" in df_26_filtered.columns:
                tips_26 = ["Todas"] + list(df_26_filtered["TIPOLOGIA"].unique())
                ft_26 = st.selectbox("Filtrar por Tipología (2026):", tips_26)
                if ft_26 != "Todas":
                    df_26_filtered = df_26_filtered[df_26_filtered["TIPOLOGIA"] == ft_26]
        with c2:
            busqueda_26 = st.text_input("Buscar en Atenciones 2026:", "")
            if busqueda_26:
                df_26_filtered = df_26_filtered[df_26_filtered.astype(str).apply(lambda x: x.str.contains(busqueda_26, case=False)).any(axis=1)]

        col_title, col_btn1, col_btn2 = st.columns([3, 1.2, 1.2])
        with col_title:
            st.subheader("🛠️ Registro de Atenciones - 2026")
        with col_btn1:
            st.download_button("📊 Excel", convertir_df_a_excel(df_26_filtered), "Atenciones_2026_Filtrado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_btn2:
            st.download_button("📄 PDF", generar_pdf_resumen("Atenciones 2026", df_26_filtered), "Atenciones_2026_Filtrado.pdf", "application/pdf", use_container_width=True)
        
        st.markdown("---")
        
        total_atenciones_26 = len(df_26_filtered)
        total_costo_26 = df_26_filtered["COSTO DE ATENCION"].sum() if "COSTO DE ATENCION" in df_26_filtered.columns else 0
        
        col1, col2 = st.columns(2)
        col1.metric("Atenciones Filtradas (2026)", total_atenciones_26)
        col2.metric("Costo Total Acumulado", f"S/. {total_costo_26:,.2f}")
        
        st.markdown("---")

        if "MES" in df_26_filtered.columns and not df_26_filtered.empty:
            st.markdown("### 📈 Tendencia de Atenciones por Mes y Costos por Tipología")
            c_g1, c_g2 = st.columns(2)
            with c_g1:
                conteo_mes = df_26_filtered["MES"].value_counts().reset_index()
                conteo_mes.columns = ["Mes", "Cantidad"]
                chart_mes = alt.Chart(conteo_mes).mark_bar(color="#1f77b4", cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(
                    x=alt.X('Mes', sort=None), y='Cantidad', tooltip=['Mes', 'Cantidad']
                ).properties(height=250, title="Atenciones por Mes")
                st.altair_chart(chart_mes, use_container_width=True)
            with c_g2:
                costo_tipo = df_26_filtered.groupby("TIPOLOGIA")["COSTO DE ATENCION"].sum().reset_index()
                chart_tipo = alt.Chart(costo_tipo).mark_arc(innerRadius=50).encode(
                    theta='COSTO DE ATENCION:Q', color='TIPOLOGIA:N', tooltip=['TIPOLOGIA', 'COSTO DE ATENCION']
                ).properties(height=250, title="Costo por Tipología (S/.)")
                st.altair_chart(chart_tipo, use_container_width=True)

        st.dataframe(df_26_filtered, use_container_width=True)

    # Pestaña 4: Cotizaciones
    elif menu_opcion == "💰 Cotizaciones":
        col_title, col_btn1, col_btn2 = st.columns([3, 1.2, 1.2])
        with col_title:
            st.subheader("💰 Cotizaciones HME")
        with col_btn1:
            st.download_button("📊 Excel", convertir_df_a_excel(st.session_state.df_cotizaciones), "Cotizaciones_HME.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_btn2:
            st.download_button("📄 PDF", generar_pdf_resumen("Cotizaciones HME", st.session_state.df_cotizaciones), "Cotizaciones_HME.pdf", "application/pdf", use_container_width=True)
        
        st.markdown("---")
        st.markdown("Listado oficial de cotizaciones vigentes para componentes y sistemas completos.")
        st.dataframe(st.session_state.df_cotizaciones, use_container_width=True)

    # Pestaña 5: Registrar Nuevo Caso (2026)
    elif menu_opcion == "➕ Registrar Nuevo Caso (2026)":
        st.subheader("➕ Agregar Nueva Atención / Caso al Sistema (2026)")
        st.markdown("Completa los campos para registrar un nuevo caso y actualizar automáticamente los reportes y costos.")
        
        with st.form("form_nuevo_caso"):
            c1, c2, c3 = st.columns(3)
            with c1:
                marca_input = st.text_input("Marca", "DELOSI")
                tienda_input = st.text_input("Tienda", "KFC ")
                ubicacion_input = st.text_input("Ubicación", "SAN ISIDRO")
            with c2:
                mes_input = st.selectbox("Mes", ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"])
                fecha_solicitud = st.date_input("Fecha de Solicitud")
                solicitante_input = st.text_input("Solicitante", "EDGAR CUSIHUAMAN")
            with c3:
                tipologia_input = st.selectbox("Tipología", ["T1", "T2", "T3", "T4"])
                horario_input = st.selectbox("Horario", ["HORARIO DE OFICINA", "NOCTURNO"])
                costo_input = st.number_input("Costo de Atención (S/.)", min_value=0.0, value=50.0, step=5.0)
            
            solicitud_input = st.text_area("Descripción de la Solicitud", "REVISIÓN TÉCNICA DE EQUIPOS HME")
            detalle_input = st.text_area("Detalle de la Atención", "SE REALIZÓ SOPORTE TÉCNICO Y VALIDACIÓN DE SISTEMA.")
            
            submit_button = st.form_submit_button(label="💾 Guardar Caso en la Base de Sesión")
            
            if submit_button:
                nuevo_item = len(st.session_state.df_at_2026) + 1
                fecha_formateada = pd.to_datetime(fecha_solicitud).strftime('%Y-%m-%d')
                
                nuevo_registro = {
                    "ITEM": nuevo_item,
                    "MARCA": marca_input,
                    "TIENDA": tienda_input,
                    "UBICACIÓN ": ubicacion_input,
                    "MES": mes_input,
                    "FECHA DE SOLICITUD": fecha_formateada,
                    "SOLICITANTE": solicitante_input,
                    "SOLICITUD": solicitud_input,
                    "DETALLE": detalle_input,
                    "TIPOLOGIA": tipologia_input,
                    "DIURNO/NOCTURNO": horario_input,
                    "COSTO DE ATENCION": costo_input
                }
                
                nuevo_df = pd.DataFrame([nuevo_registro])
                st.session_state.df_at_2026 = pd.concat([st.session_state.df_at_2026, nuevo_df], ignore_index=True)
                st.success("¡Caso registrado correctamente! Ya se refleja en las estadísticas, gráficos y opciones de exportación.")
