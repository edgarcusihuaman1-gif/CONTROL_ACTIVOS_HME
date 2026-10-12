import streamlit as st
import pandas as pd
import altair as alt
import io
import os
import json
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
USERS_FILE = "usuarios.json"

# --- GESTIÓN DE USUARIOS (JSON) ---
def cargar_usuarios():
    if not os.path.exists(USERS_FILE):
        usuarios_default = {
            "admin": {"clave": "hme2026", "rol": "Administrador", "nombre": "Edgar (Admin)"},
            "equipo": {"clave": "hme2025", "rol": "Visualizador", "nombre": "Colaborador HME"}
        }
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(usuarios_default, f, indent=4, ensure_ascii=False)
        return usuarios_default
    else:
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)

def guardar_usuarios(dict_usuarios):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(dict_usuarios, f, indent=4, ensure_ascii=False)

usuarios_db = cargar_usuarios()

def cargar_datos(file_path):
    if not os.path.exists(file_path):
        st.error(f"No se encuentra el archivo {file_path}. Asegúrate de que esté en la carpeta del proyecto.")
        return None, None, None, None
        
    xls = pd.ExcelFile(file_path)
    df_hme = pd.read_excel(xls, 'Drive HME')
    df_at_2025 = pd.read_excel(xls, 'Atenciones 2025')
    df_at_2026 = pd.read_excel(xls, 'Atenciones 2026')
    df_cotizaciones = pd.read_excel(xls, 'Cotizaciones')
    
    for df in [df_at_2025, df_at_2026]:
        if 'FECHA DE SOLICITUD' in df.columns:
            df['FECHA DE SOLICITUD'] = pd.to_datetime(df['FECHA DE SOLICITUD'], errors='coerce').dt.strftime('%Y-%m-%d')
            
    return df_hme, df_at_2025, df_at_2026, df_cotizaciones

if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.rol = None
    st.session_state.usuario_actual = ""

if 'df_hme' not in st.session_state:
    df_h, df_25, df_26, df_cot = cargar_datos(EXCEL_FILE)
    st.session_state.df_hme = df_h
    st.session_state.df_at_2025 = df_25
    st.session_state.df_at_2026 = df_26
    st.session_state.df_cotizaciones = df_cot

# --- LOGIN ---
if not st.session_state.autenticado:
    st.title("🔐 Acceso al Sistema - HME Drive Thru")
    st.markdown("---")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("### Por favor, ingresa tus credenciales")
        
        with st.form("form_login"):
            usuario_ingresado = st.text_input("Usuario")
            clave_ingresada = st.text_input("Contraseña", type="password")
            btn_login = st.form_submit_button("Ingresar al Sistema", use_container_width=True)
            
            if btn_login:
                usr = usuario_ingresado.strip().lower()
                pwd = clave_ingresada.strip()
                
                if usr in usuarios_db and usuarios_db[usr]["clave"] == pwd:
                    st.session_state.autenticado = True
                    st.session_state.rol = usuarios_db[usr]["rol"]
                    st.session_state.usuario_actual = usuarios_db[usr]["nombre"]
                    st.success(f"¡Bienvenido {st.session_state.usuario_actual}!")
                    st.rerun()
                else:
                    st.error("Usuario o contraseña incorrectos.")
                    
        st.markdown("---")
        st.info("💡 **Admin:** `admin` / `hme2026`\n\n🔹 **Equipo:** `equipo` / `hme2025`")
    st.stop()

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
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor('#1f77b4'), spaceAfter=15)
    
    elements.append(Paragraph(f"<b>Reporte: {titulo}</b>", title_style))
    elements.append(Spacer(1, 10))
    
    if df.empty:
        elements.append(Paragraph("Sin registros.", styles['Normal']))
    else:
        cols = list(df.columns[:6])
        data_t = [cols]
        for _, row in df.head(30).iterrows():
            data_t.append([str(row[c])[:25] for c in cols])
        t = Table(data_t, colWidths=[80, 80, 90, 90, 85, 85])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f77b4')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#d6d6d6')),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
        ]))
        elements.append(t)
        
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()

# --- APLICACIÓN ---
if 'df_hme' in st.session_state:
    st.title("📊 Sistema de Gestión y Control - HME Drive Thru")
    
    st.sidebar.header("👤 Sesión Activa")
    st.sidebar.markdown(f"**Usuario:** {st.session_state.usuario_actual}")
    st.sidebar.markdown(f"**Perfil:** `{st.session_state.rol}`")
    
    if st.session_state.rol == "Administrador":
        if st.sidebar.button("🔄 Sincronizar / Recargar Excel", use_container_width=True):
            df_h, df_25, df_26, df_cot = cargar_datos(EXCEL_FILE)
            st.session_state.df_hme = df_h
            st.session_state.df_at_2025 = df_25
            st.session_state.df_at_2026 = df_26
            st.session_state.df_cotizaciones = df_cot
            st.success("¡Datos sincronizados!")
            st.rerun()

    if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.session_state.rol = None
        st.rerun()

    st.sidebar.markdown("---")
    st.sidebar.header("🔍 Panel de Navegación")
    
    if st.session_state.rol == "Administrador":
        menu_opcion = st.sidebar.radio("Ir a la sección:", [
            "🏢 Estado Drive HME", "🛠️ Atenciones 2025", "🛠️ Atenciones 2026", 
            "💰 Cotizaciones", "➕ Registrar Nuevo Caso (2026)", "👤 Gestión de Usuarios (Admin)"
        ])
    else:
        menu_opcion = st.sidebar.radio("Ir a la sección:", [
            "🏢 Estado Drive HME", "🛠️ Atenciones 2025", "🛠️ Atenciones 2026", "💰 Cotizaciones"
        ])

    st.sidebar.markdown("---")

    if menu_opcion == "🏢 Estado Drive HME":
        c1, c2 = st.columns(2)
        with c1:
            marcas = ["Todas"] + list(st.session_state.df_hme["MARCA"].dropna().unique())
            filtro_marca = st.selectbox("Filtrar por Marca:", marcas)
        with c2:
            busqueda_tienda = st.text_input("Buscar Tienda / Código / Ubicación:", "")

        df_f = st.session_state.df_hme.copy()
        if filtro_marca != "Todas":
            df_f = df_f[df_f["MARCA"] == filtro_marca]
        if busqueda_tienda:
            df_f = df_f[
                df_f["TIENDA"].astype(str).str.contains(busqueda_tienda, case=False, na=False) |
                df_f["UBICACIÓN"].astype(str).str.contains(busqueda_tienda, case=False, na=False)
            ]

        col_title, col_btn1, col_btn2 = st.columns([3, 1.2, 1.2])
        with col_title:
            st.subheader("🏢 Estado General de Equipos en Tiendas HME")
        with col_btn1:
            st.download_button("📊 Excel", convertir_df_a_excel(df_f), "Estado_HME.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_btn2:
            st.download_button("📄 PDF", generar_pdf_resumen("Estado Drive HME", df_f), "Estado_HME.pdf", "application/pdf", use_container_width=True)
        
        st.markdown("---")
        total_t = len(df_f)
        total_op = df_f["HEADPHONE OPERATIVOS"].sum() if "HEADPHONE OPERATIVOS" in df_f.columns else 0
        total_av = df_f["HEADPHONE AVERIADOS"].sum() if "HEADPHONE AVERIADOS" in df_f.columns else 0
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Tiendas Monitoreadas", total_t)
        c2.metric("Headphones Operativos", int(total_op))
        c3.metric("Headphones Averiados", int(total_av))
        st.markdown("---")
        st.dataframe(df_f, use_container_width=True)

    elif menu_opcion == "🛠️ Atenciones 2025":
        df_25 = st.session_state.df_at_2025.copy()
        busq = st.text_input("Buscar en Atenciones 2025:", "")
        if busq:
            df_25 = df_25[df_25.astype(str).apply(lambda x: x.str.contains(busq, case=False)).any(axis=1)]
        st.subheader("🛠️ Atenciones - 2025")
        st.metric("Total", len(df_25))
        st.dataframe(df_25, use_container_width=True)

    elif menu_opcion == "🛠️ Atenciones 2026":
        df_26 = st.session_state.df_at_2026.copy()
        busq = st.text_input("Buscar en Atenciones 2026:", "")
        if busq:
            df_26 = df_26[df_26.astype(str).apply(lambda x: x.str.contains(busq, case=False)).any(axis=1)]
        st.subheader("🛠️ Atenciones - 2026")
        tot_c = df_26["COSTO DE ATENCION"].sum() if "COSTO DE ATENCION" in df_26.columns else 0
        c1, c2 = st.columns(2)
        c1.metric("Atenciones", len(df_26))
        c2.metric("Costo Total", f"S/. {tot_c:,.2f}")
        st.dataframe(df_26, use_container_width=True)

    elif menu_opcion == "💰 Cotizaciones":
        st.subheader("💰 Cotizaciones HME")
        st.dataframe(st.session_state.df_cotizaciones, use_container_width=True)

    elif menu_opcion == "➕ Registrar Nuevo Caso (2026)" and st.session_state.rol == "Administrador":
        st.subheader("➕ Registrar Nuevo Caso (2026)")
        with st.form("form_caso"):
            c1, c2, c3 = st.columns(3)
            with c1:
                marca = st.text_input("Marca", "DELOSI")
                tienda = st.text_input("Tienda", "KFC")
                ubicacion = st.text_input("Ubicación", "SAN ISIDRO")
            with c2:
                mes = st.selectbox("MES", ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"])
                fecha = st.date_input("FECHA")
                solicitante = st.text_input("SOLICITANTE", "EDGAR")
            with c3:
                horario = st.selectbox("DIURNO/NOCTURNO", ["HORARIO DE OFICINA", "NOCTURNO"])
                costo = st.number_input("COSTO DE ATENCION", min_value=0.0, value=50.0)
            
            solicitud = st.text_input("SOLICITUD", "SOPORTE")
            detalle = st.text_area("DETALLE", "SOPORTE TÉCNICO")
            
            if st.form_submit_button("💾 Guardar"):
                nuevo = pd.DataFrame([{
                    "ITEM": len(st.session_state.df_at_2026) + 1,
                    "MARCA": marca, "TIENDA": tienda, "UBICACIÓN\xa0": ubicacion,
                    "MES": mes, "FECHA DE SOLICITUD": pd.to_datetime(fecha).strftime('%Y-%m-%d'),
                    "SOLICITANTE": solicitante, "SOLICITUD": solicitud, "DETALLE": detalle,
                    "DIURNO/NOCTURNO": horario, "COSTO DE ATENCION": costo
                }])
                st.session_state.df_at_2026 = pd.concat([st.session_state.df_at_2026, nuevo], ignore_index=True)
                st.success("¡Registrado!")

    elif menu_opcion == "👤 Gestión de Usuarios (Admin)" and st.session_state.rol == "Administrador":
        st.subheader("👤 Gestión de Usuarios")
        c1, c2 = st.columns(2)
        with c1:
            with st.form("form_usr"):
                usr_n = st.text_input("Usuario").strip().lower()
                pwd_n = st.text_input("Contraseña", type="password").strip()
                nom_n = st.text_input("Nombre").strip()
                rol_n = st.selectbox("Rol", ["Visualizador", "Administrador"])
                if st.form_submit_button("💾 Guardar"):
                    if usr_n and pwd_n:
                        usuarios_db[usr_n] = {"clave": pwd_n, "rol": rol_n, "nombre": nom_n}
                        guardar_usuarios(usuarios_db)
                        st.success("¡Guardado!")
                        st.rerun()
        with c2:
            st.dataframe(pd.DataFrame([{"Usuario": k, **v} for k, v in usuarios_db.items()]), use_container_width=True)