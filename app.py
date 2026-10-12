import streamlit as st
import pandas as pd
from datetime import datetime
import io

# Importaciones de ReportLab para exportación nativa a PDF
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Configuración de página web
st.set_page_config(
    page_title="Panel HME Autoservicio",
    page_icon="🚗",
    layout="wide"
)

# Estilo CSS para ocultar menú nativo y comprimir encabezados/tablas
estilo_oculto = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    h1 { font-size: 1.6rem !important; padding-top: 0rem !important; }
    h2 { font-size: 1.3rem !important; }
    h3 { font-size: 1.1rem !important; }
    .stHeader { padding-top: 0.5rem !important; }
    [data-testid="stTable"] th, [data-testid="stDataEditor"] th {
        font-size: 13px !important;
        white-space: nowrap !important;
    }
    </style>
"""
st.markdown(estilo_oculto, unsafe_allow_html=True)

EXCEL_FILE = "DRIVE TRHU BASE.xlsx"

# -------------------------------------------------------------
# BASE DE DATOS DE USUARIOS Y ROLES
# Roles disponibles: 'admin', 'visor_exportador', 'solo_vista'
# -------------------------------------------------------------
USUARIOS = {
    "admin": {"password": "123", "rol": "admin", "nombre": "Administrador"},
    "reportes": {"password": "123", "rol": "visor_exportador", "nombre": "Usuario Exportación"},
    "invitado": {"password": "123", "rol": "solo_vista", "nombre": "Usuario Solo Lectura"}
}

# Control de Sesión / Login
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario_actual = None
    st.session_state.rol_actual = None

def login():
    st.title("🔒 Control de Acceso - HME System")
    with st.form("form_login"):
        user = st.text_input("Usuario:").strip().lower()
        pwd = st.text_input("Contraseña:", type="password")
        submit = st.form_submit_button("Iniciar Sesión")

        if submit:
            if user in USUARIOS and USUARIOS[user]["password"] == pwd:
                st.session_state.autenticado = True
                st.session_state.usuario_actual = USUARIOS[user]["nombre"]
                st.session_state.rol_actual = USUARIOS[user]["rol"]
                st.rerun()
            else:
                st.error("⚠️ Usuario o contraseña incorrectos.")

def logout():
    st.session_state.autenticado = False
    st.session_state.usuario_actual = None
    st.session_state.rol_actual = None
    st.rerun()

if not st.session_state.autenticado:
    login()
    st.stop()

# -------------------------------------------------------------
# CARGA Y PERSISTENCIA DE DATOS
# -------------------------------------------------------------
@st.cache_data
def cargar_datos_iniciales():
    xls = pd.ExcelFile(EXCEL_FILE)
    df_hme = pd.read_excel(xls, sheet_name="Drive HME")
    df_2025 = pd.read_excel(xls, sheet_name="Atenciones 2025")
    df_2026 = pd.read_excel(xls, sheet_name="Atenciones 2026")
    df_cotiz = pd.read_excel(xls, sheet_name="Cotizaciones")
    return df_hme, df_2025, df_2026, df_cotiz

if 'df_hme' not in st.session_state:
    df_hme_ini, df_2025_ini, df_2026_ini, df_cotiz_ini = cargar_datos_iniciales()
    st.session_state.df_hme = df_hme_ini
    st.session_state.df_2025 = df_2025_ini
    st.session_state.df_2026 = df_2026_ini
    st.session_state.df_cotiz = df_cotiz_ini

RENOMBRAR_GLOBAL = {
    'HEADPHONE OPERATIVOS': 'Headph. Operat.',
    'HEADPHONE AVERIADOS': 'Headph. Averiados',
    'TOTAL HEADPHONE': 'Total Headph.',
    'CARGADOR DE BATERIAS': 'Cargador Bat.',
    'CARGADOR OPERATIVO': 'Estado Cargador',
    'UBICACIÓN': 'Ubicación',
    'TIENDA': 'Tienda',
    'TIPO DE ATENCION': 'Tipo Atenc.',
    'COSTO DE ATENCION': 'Costo ($)',
    'DIAGNOSTICO': 'Diagnóstico',
    'SOLUCION': 'Solución',
    'TECNICO': 'Técnico',
    'FECHA': 'Fecha'
}

REVERSO_GLOBAL = {v: k for k, v in RENOMBRAR_GLOBAL.items()}

def acortar_columnas(df):
    columnas_existentes = {k: v for k, v in RENOMBRAR_GLOBAL.items() if k in df.columns}
    return df.rename(columns=columnas_existentes)

def guardar_hoja_excel(df, nombre_hoja):
    try:
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=nombre_hoja, index=False)
        return True
    except Exception:
        return False

# -------------------------------------------------------------
# FUNCIONES AUXILIARES DE EXPORTACIÓN (EXCEL Y PDF REAL)
# -------------------------------------------------------------
def exportar_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Reporte')
    return output.getvalue()

def generar_pdf(titulo, df):
    """Genera un archivo PDF ajustando automáticamente el ancho de tabla y texto."""
    df_corto = acortar_columnas(df)
    buffer = io.BytesIO()
    
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        rightMargin=15,
        leftMargin=15,
        topMargin=15,
        bottomMargin=15
    )
    
    elementos = []
    estilos = getSampleStyleSheet()
    
    estilo_titulo = ParagraphStyle(
        'TituloPDF',
        parent=estilos['Heading1'],
        fontSize=14,
        leading=16,
        textColor=colors.HexColor('#1E3A8A'),
        alignment=1,
        spaceAfter=6
    )
    
    estilo_fecha = ParagraphStyle(
        'FechaPDF',
        parent=estilos['Normal'],
        fontSize=8,
        leading=10,
        textColor=colors.HexColor('#666666'),
        alignment=2,
        spaceAfter=10
    )

    elementos.append(Paragraph("🚗 Sistema HME Drive-Thru", estilo_titulo))
    elementos.append(Paragraph(f"<b>{titulo}</b>", estilos['Heading2']))
    elementos.append(Paragraph(f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}", estilo_fecha))
    
    estilo_celda = ParagraphStyle('Celda', parent=estilos['Normal'], fontSize=7, leading=9)
    estilo_encabezado = ParagraphStyle('HeaderCelda', parent=estilos['Normal'], fontSize=7, leading=9, textColor=colors.white, fontName='Helvetica-Bold')
    
    datos_tabla = []
    headers = [Paragraph(str(col), estilo_encabezado) for col in df_corto.columns]
    datos_tabla.append(headers)
    
    for _, fila in df_corto.iterrows():
        fila_texto = [Paragraph(str(val) if pd.notna(val) else "", estilo_celda) for val in fila]
        datos_tabla.append(fila_texto)
    
    num_cols = len(df_corto.columns)
    ancho_pagina_usable = 762
    ancho_columna = ancho_pagina_usable / max(num_cols, 1)
    
    tabla = Table(datos_tabla, colWidths=[ancho_columna] * num_cols, repeatRows=1)
    tabla.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')])
    ]))
    
    elementos.append(tabla)
    doc.build(elementos)
    
    return buffer.getvalue()

# -------------------------------------------------------------
# MENÚ DE NAVEGACIÓN Y PERFIL
# -------------------------------------------------------------
st.sidebar.title("🚗 HME Control")
st.sidebar.caption(f"👤 **Usuario:** {st.session_state.usuario_actual}")

opciones_menu = ["📊 Panel General", "📋 Inventario de Tiendas", "🛠️ Histórico Atenciones", "💵 Cotizaciones"]

if st.session_state.rol_actual == "admin":
    opciones_menu.insert(3, "📝 Nueva Atención")

if st.session_state.rol_actual in ["admin", "visor_exportador"]:
    opciones_menu.append("📥 Exportar Datos")

opcion = st.sidebar.radio("Menú Principal:", opciones_menu)

if st.sidebar.button("🚪 Cerrar Sesión"):
    logout()

# -------------------------------------------------------------
# MÓDULO 1: PANEL GENERAL
# -------------------------------------------------------------
if opcion == "📊 Panel General":
    st.title("📊 Panel General HME")
    st.caption("Indicadores clave e inventario general")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Tiendas", len(st.session_state.df_hme))
    c2.metric("Auric. Operativos", int(st.session_state.df_hme['HEADPHONE OPERATIVOS'].sum()))
    c3.metric("Auric. Averiados", int(st.session_state.df_hme['HEADPHONE AVERIADOS'].sum()))
    
    costo_total = st.session_state.df_2026['COSTO DE ATENCION'].sum() if 'COSTO DE ATENCION' in st.session_state.df_2026.columns else 0
    c4.metric("Gasto 2026", f"${costo_total:,.2f} USD")

    st.markdown("---")
    col_left, col_right = st.columns([1.2, 1])
    
    with col_left:
        st.subheader("📶 Estado de Auriculares")
        st.bar_chart(st.session_state.df_hme.set_index('TIENDA')[['HEADPHONE OPERATIVOS', 'HEADPHONE AVERIADOS']])

    with col_right:
        st.subheader("⚠️ Equipos en Estado Crítico")
        criticos = st.session_state.df_hme[
            (st.session_state.df_hme['CARGADOR OPERATIVO'] != 'OPERATIVO') | 
            (st.session_state.df_hme['HEADPHONE AVERIADOS'] > 0)
        ][['TIENDA', 'UBICACIÓN', 'HEADPHONE AVERIADOS', 'CARGADOR OPERATIVO']]
        st.dataframe(acortar_columnas(criticos), use_container_width=True, hide_index=True)

# -------------------------------------------------------------
# MÓDULO 2: INVENTARIO DE TIENDAS
# -------------------------------------------------------------
elif opcion == "📋 Inventario de Tiendas":
    st.title("📋 Inventario de Tiendas")

    if st.session_state.rol_actual == "admin":
        st.info("💡 Edite los datos directamente en la tabla y presione 'Guardar Cambios':")
        df_editado = st.data_editor(
            acortar_columnas(st.session_state.df_hme), 
            num_rows="dynamic", 
            use_container_width=True,
            key="editor_tiendas"
        )
        col_btn1, col_btn2, col_btn3 = st.columns([2, 1, 1])
        with col_btn1:
            if st.button("💾 Guardar Cambios"):
                df_restaurado = df_editado.rename(columns=REVERSO_GLOBAL)
                st.session_state.df_hme = df_restaurado
                if guardar_hoja_excel(df_restaurado, "Drive HME"):
                    st.success("✅ ¡Inventario actualizado en la app y en Excel!")
                else:
                    st.warning("⚠️ Guardado en memoria. Cierre el Excel físico si lo tiene abierto.")
        with col_btn2:
            st.download_button(
                "📊 Exportar Excel", 
                exportar_excel(st.session_state.df_hme), 
                f"Inventario_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col_btn3:
            st.download_button(
                "📄 Exportar PDF", 
                generar_pdf("Reporte Inventario", st.session_state.df_hme), 
                f"Inventario_{datetime.now().strftime('%Y%m%d')}.pdf", 
                mime="application/pdf"
            )
    
    elif st.session_state.rol_actual == "visor_exportador":
        st.dataframe(acortar_columnas(st.session_state.df_hme), use_container_width=True, hide_index=True)
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "📊 Exportar Excel", 
                exportar_excel(st.session_state.df_hme), 
                f"Inventario_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col2:
            st.download_button(
                "📄 Exportar PDF", 
                generar_pdf("Reporte Inventario", st.session_state.df_hme), 
                f"Inventario_{datetime.now().strftime('%Y%m%d')}.pdf", 
                mime="application/pdf"
            )

    else:  # solo_vista
        st.warning("🔒 Modo Lectura. No tiene permisos para editar ni descargar información.")
        st.dataframe(acortar_columnas(st.session_state.df_hme), use_container_width=True, hide_index=True)

# -------------------------------------------------------------
# MÓDULO 3: HISTÓRICO DE ATENCIONES
# -------------------------------------------------------------
elif opcion == "🛠️ Histórico Atenciones":
    st.title("🛠️ Histórico de Atenciones")
    t1, t2 = st.tabs(["Atenciones 2026", "Atenciones 2025"])
    
    with t1:
        st.subheader("Año 2026")
        st.dataframe(acortar_columnas(st.session_state.df_2026), use_container_width=True, hide_index=True)
        if st.session_state.rol_actual in ["admin", "visor_exportador"]:
            c1, c2 = st.columns(2)
            with c1:
                st.download_button(
                    "📊 Descargar Excel (2026)", 
                    exportar_excel(st.session_state.df_2026), 
                    f"Atenciones_2026_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with c2:
                st.download_button(
                    "📄 Descargar PDF (2026)", 
                    generar_pdf("Atenciones 2026", st.session_state.df_2026), 
                    f"Atenciones_2026_{datetime.now().strftime('%Y%m%d')}.pdf", 
                    mime="application/pdf"
                )

    with t2:
        st.subheader("Año 2025")
        st.dataframe(acortar_columnas(st.session_state.df_2025), use_container_width=True, hide_index=True)
        if st.session_state.rol_actual in ["admin", "visor_exportador"]:
            c1, c2 = st.columns(2)
            with c1:
                st.download_button(
                    "📊 Descargar Excel (2025)", 
                    exportar_excel(st.session_state.df_2025), 
                    f"Atenciones_2025_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with c2:
                st.download_button(
                    "📄 Descargar PDF (2025)", 
                    generar_pdf("Atenciones 2025", st.session_state.df_2025), 
                    f"Atenciones_2025_{datetime.now().strftime('%Y%m%d')}.pdf", 
                    mime="application/pdf"
                )

# -------------------------------------------------------------
# MÓDULO 4: REGISTRAR NUEVA ATENCIÓN (SOLO ADMIN)
# -------------------------------------------------------------
elif opcion == "📝 Nueva Atención" and st.session_state.rol_actual == "admin":
    st.title("📝 Registrar Atención")
    with st.form("formulario_atencion"):
        col1, col2 = st.columns(2)
        with col1:
            tiendas_disponibles = st.session_state.df_hme['TIENDA'].unique()
            tienda_seleccionada = st.selectbox("Tienda:", tiendas_disponibles)
            fecha_atencion = st.date_input("Fecha:", datetime.now())
            tipo_atencion = st.selectbox("Tipo:", ["Preventivo", "Correctivo", "Garantía", "Instalación"])
            tecnico_responsable = st.text_input("Técnico:")

        with col2:
            diagnostico_tecnico = st.text_area("Diagnóstico / Problema:")
            solucion_aplicada = st.text_area("Solución Realizada:")
            costo_servicio = st.number_input("Costo ($ USD):", min_value=0.0, step=10.0)

        boton_enviar = st.form_submit_button("💾 Guardar Atención")

        if boton_enviar:
            filtro_tienda = st.session_state.df_hme[st.session_state.df_hme['TIENDA'] == tienda_seleccionada]
            ubicacion = filtro_tienda['UBICACIÓN'].values[0] if len(filtro_tienda) > 0 else ""

            nuevo_ticket = pd.DataFrame([{
                "FECHA": fecha_atencion.strftime("%Y-%m-%d"),
                "TIENDA": tienda_seleccionada,
                "UBICACIÓN": ubicacion,
                "TIPO DE ATENCION": tipo_atencion,
                "TECNICO": tecnico_responsable,
                "DIAGNOSTICO": diagnostico_tecnico,
                "SOLUCION": solucion_aplicada,
                "COSTO DE ATENCION": costo_servicio
            }])

            st.session_state.df_2026 = pd.concat([st.session_state.df_2026, nuevo_ticket], ignore_index=True)
            guardar_hoja_excel(st.session_state.df_2026, "Atenciones 2026")
            st.success("✅ ¡Atención registrada y guardada!")

# -------------------------------------------------------------
# MÓDULO 5: CATÁLOGO Y COTIZACIONES
# -------------------------------------------------------------
elif opcion == "💵 Cotizaciones":
    st.title("💵 Catálogo de Repuestos")
    st.dataframe(acortar_columnas(st.session_state.df_cotiz), use_container_width=True, hide_index=True)

# -------------------------------------------------------------
# MÓDULO 6: EXPORTAR DATOS (ADMIN Y VISOR_EXPORTADOR)
# -------------------------------------------------------------
elif opcion == "📥 Exportar Datos" and st.session_state.rol_actual in ["admin", "visor_exportador"]:
    st.title("📥 Centro de Exportación")
    modulo_exportar = st.selectbox("Seleccionar Tabla:", ["Inventario de Tiendas", "Atenciones 2026", "Atenciones 2025", "Catálogo de Repuestos"])

    if modulo_exportar == "Inventario de Tiendas":
        df_target = st.session_state.df_hme
    elif modulo_exportar == "Atenciones 2026":
        df_target = st.session_state.df_2026
    elif modulo_exportar == "Atenciones 2025":
        df_target = st.session_state.df_2025
    else:
        df_target = st.session_state.df_cotiz

    st.dataframe(acortar_columnas(df_target), use_container_width=True, hide_index=True)

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            "📊 Descargar Excel", 
            exportar_excel(df_target), 
            f"{modulo_exportar}.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col2:
        st.download_button(
            "📄 Descargar PDF", 
            generar_pdf(modulo_exportar, df_target), 
            f"{modulo_exportar}.pdf", 
            mime="application/pdf", 
            use_container_width=True
        )
