import streamlit as st
from streamlit_calendar import calendar
import pandas as pd
from datetime import datetime, time, timedelta
import io
import os

# 1. CONFIGURACIÓN DE LA PÁGINA
st.set_page_config(layout="wide", page_title="Planificación VSP", page_icon="📅")

# --- VARIABLES GLOBALES ---
ARCHIVO_DB = "base_datos_vsp_sucre.xlsx"
CONTRASENA_ADMIN = "Vsp26"  

LISTA_RESPONSABLES = ["Seleccione...", "ADALGISA PATRON CONDE", "ANA GABRIELA DIAZ ANAYA", "ANA LUCIA MENDOZA TAMARA", "BALDIR PABA OSORIO", "LILIBETH DAZA CAMELO", "EDER JESUS PATERNINA RODRIGUEZ" , "ELIANA CECILIA MORALES MELENDEZ" , "ENITT DEL ROSARIO HERNANDEZ DORIAS" , "ESPERANZA DEL PILAR VARGAS VARGAS" , "HECTOR FABIO RENTERIA" , "ISAAC JACOB VELASQUEZ DOMINGUEZ" , "JAVIER MAURICIO CORREA PATERNOSTRO" , "KAREN MARGARITA ALDANA ARRIETA" , "KEVIN ALBERTO BARBARAN ALVAREZ" , "LEVY SUNILDA CAMPO LASSO" , "LOLI LUZ SIERRA DIAZ" , "LORENA PORTILLO CUENTAS" , "LUCIA CLARETH HERNANDEZ PEREZ" , "LUISA FERNANDA REYES DÍAZ" , "LUZMILA VILLAMIZAR MOLINA" , "MARIA CANDELARIA MEJIA LOPEZ" , "MARIA JOSE CANTILLO ROYERO" , "MARLON ESPITIA CERPA" , "MARTHA CECILIA MELENDEZ MARTINEZ" , "MERY DE JESUS NARVAEZ ASSIA" , "NICOLASA MARGARITA ARRIETA SERPA" , "NURYS CONCEPCIÓN HERRERA GUTIÉRREZ" , "VIRGINIA OLIVERO GARCIA" , "YARLENY ESTHER BERRIO ACOSTA" , "MARIA JOSE PEÑARANDA" , "BRENDER BARRIOS" , "ANA KARINA PEÑATES DE ARCE" , "DINO VERGARA PEREZ" , "JUAN CARLOS GARCIA VIVERO" , "MANUEL ORTEGA HERNANDEZ" , "MARIA CAMPO" , "VILMA MERCADO CUMPLIDO"]
LISTA_MUNICIPIOS = ["Seleccione...", "Buenavista", "Caimito", "Chalán", "Colosó", "Corozal", "Coveñas", "El Roble", "Galeras", "Guaranda", "La Unión", "Los Palmitos", "Majagual", "Morroa", "Ovejas", "Palmito", "Sampués", "San Benito Abad", "San Juan de Betulia", "San Marcos", "San Onofre", "San Pedro", "Sincé", "Sincelejo", "Sucre", "Tolú", "Toluviejo"]
LISTA_LUGARES = ["Seleccione...", "Sala Situacional", "Auditorio Panzigua", "Otro"]
LISTA_TIPOS_EVENTO = ["ASISTENCIA TECNICA", "BAC", "BAI", "CAPACITACION", "COMITÉ ESTADISTICAS VITALES", "COMITÉ SANIDAD PORTUARIA", "COVE" , "IEC" , "MESA DE TRABAJO" , "MONITOREO" , "REUNION" , "SAR" , "SEGUIMIENTO" , "UNIDAD DE ANALISIS" , "OTRO"]

# ==========================================
# FUNCIONES DE PERSISTENCIA SEGURA
# ==========================================
def inicializar_db():
    if not os.path.exists(ARCHIVO_DB):
        with pd.ExcelWriter(ARCHIVO_DB, engine='openpyxl') as writer:
            pd.DataFrame(columns=["Fecha", "Hora Inicio", "Hora Fin", "Responsable", "Tipo de Evento", "Municipio", "Lugar", "Vehículo", "Estado", "Observaciones"]).to_excel(writer, sheet_name='Eventos', index=False)
            pd.DataFrame(columns=["Semana_Inicio", "Integrantes", "Cargos"]).to_excel(writer, sheet_name='Disponibilidad', index=False)
            pd.DataFrame(columns=["Fecha_Acuerdo", "Compromiso", "Responsable", "Plazo", "Estado", "Respuesta_Avance"]).to_excel(writer, sheet_name='Compromisos', index=False)
    else:
        # Aseguramos de manera transparente que la columna de respuesta exista si el archivo ya venía de versiones previas
        try:
            excel_file = pd.ExcelFile(ARCHIVO_DB)
            if 'Compromisos' in excel_file.sheet_names:
                df_c_check = pd.read_excel(ARCHIVO_DB, sheet_name='Compromisos')
                if "Respuesta_Avance" not in df_c_check.columns:
                    df_c_check["Respuesta_Avance"] = ""
                    guardar_datos(df_c_check, 'Compromisos')
        except:
            pass

def cargar_datos(hoja):
    try:
        return pd.read_excel(ARCHIVO_DB, sheet_name=hoja).fillna("")
    except:
        return pd.DataFrame()

def guardar_datos(df, hoja):
    try:
        excel_file = pd.ExcelFile(ARCHIVO_DB)
        hojas = {h: excel_file.parse(h) for h in excel_file.sheet_names}
        hojas[hoja] = df 
        
        with pd.ExcelWriter(ARCHIVO_DB, engine='openpyxl') as writer:
            for nombre, datos in hojas.items():
                datos.to_excel(writer, sheet_name=nombre, index=False)
        return True
    except Exception as e:
        st.error(f"Error al guardar: {e}")
        return False

inicializar_db()

# ==========================================
# GESTIÓN DE NAVEGACIÓN (ESTADO DE LA SESIÓN)
# ==========================================
if "seccion_actual" not in st.session_state:
    st.session_state["seccion_actual"] = "🏠 Inicio / Resumen General"

# ==========================================
# BARRA LATERAL (SIDEBAR) - MENÚ DE BOTONES
# ==========================================
with st.sidebar:
    if os.path.exists("logo.png"): 
        st.image("logo.png", use_container_width=True)
    else: 
        st.headline("📋 VSP Sucre")
    
    st.markdown("### 🧭 Navegación")
    st.markdown("Seleccione el módulo al que desea acceder:")
    
    if st.button("🏠 Inicio / Resumen General", use_container_width=True):
        st.session_state["seccion_actual"] = "🏠 Inicio / Resumen General"
        st.rerun()
        
    if st.button("📝 Registrar Actividad", use_container_width=True):
        st.session_state["seccion_actual"] = "📝 Registrar Actividad"
        st.rerun()
        
    if st.button("🛡️ Disponibilidad Semanal", use_container_width=True):
        st.session_state["seccion_actual"] = "🛡️ Disponibilidad Semanal"
        st.rerun()
        
    if st.button("📝 Compromisos Técnicos", use_container_width=True):
        st.session_state["seccion_actual"] = "📝 Compromisos Técnicos"
        st.rerun()
        
    if st.button("🔍 Filtros y Dashboard", use_container_width=True):
        st.session_state["seccion_actual"] = "🔍 Filtros y Dashboard"
        st.rerun()
        
    if st.button("⚙️ Panel Maestro", use_container_width=True):
        st.session_state["seccion_actual"] = "⚙️ Panel Maestro"
        st.rerun()

    st.markdown("---")
    st.caption("Subprograma de Vigilancia en Salud Pública")

# ==========================================
# ENCABEZADO PRINCIPAL (FIJO)
# ==========================================
st.title("Sistema de Planificación VSP Sucre")
st.markdown("##### Gobernación de Sucre • Dirección de Salud Pública")
st.markdown(f"**Módulo actual:** `{st.session_state['seccion_actual']}`")
st.markdown("---")

# ==========================================
# PROCESAMIENTO Y CARGA DE DATOS PARA MÉTRICAS
# ==========================================
df_meta = cargar_datos('Eventos')
hoy = datetime.today().date()
veh_hoy = 0

if not df_meta.empty:
    df_meta["Fecha_DT"] = pd.to_datetime(df_meta["Fecha"], errors='coerce')
    df_ua = df_meta[(df_meta["Lugar"] == "UNIDAD DE ANALISIS") | (df_meta["Tipo de Evento"] == "UNIDAD DE ANALISIS")]
    ua_vig = len(df_ua[df_ua["Fecha_DT"].dt.date >= hoy])
    ua_ven = len(df_ua[df_ua["Fecha_DT"].dt.date < hoy])
    sala_mes = len(df_meta[(df_meta["Lugar"] == "Sala Situacional") & (df_meta["Fecha_DT"].dt.month == hoy.month)])
    veh_hoy = len(df_meta[(df_meta["Fecha"].astype(str) == hoy.strftime("%Y-%m-%d")) & (df_meta["Vehículo"] == "Sí")])

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("📌 Total Actividades", len(df_meta))
    m2.metric("📅 Actividades Hoy", len(df_meta[df_meta["Fecha"].astype(str) == hoy.strftime("%Y-%m-%d")]))
    m3.metric("🏢 Sala Situacional (Mes)", f"{sala_mes} Prog.")
    m4.metric("🧬 U. de Análisis", f"{len(df_ua)} Total", delta=f"🟢 {ua_vig} Vig | 🔴 {ua_ven} Ven", delta_color="normal")
    m5.metric("🚗 Necesitan Vehículo", veh_hoy)
    
    if veh_hoy >= 2 and st.session_state["seccion_actual"] == "🏠 Inicio / Resumen General":
        st.warning(f"⚠️ **Alerta Logística:** Hoy hay {veh_hoy} salidas que requieren vehículo. Coordinar prioridades.")


# ==========================================
# RENDERIZADO DINÁMICO DE PANTALLAS
# ==========================================

# --- PANTALLA 1: INICIO / RESUMEN GENERAL ---
if st.session_state["seccion_actual"] == "🏠 Inicio / Resumen General":
    
    # --- FILTRO GLOBAL DE LUGARES / ACTIVIDADES ---
    st.markdown("### 🔍 Filtrar Agenda Mensual")
    
    # Extraemos los lugares únicos de la base de datos para armar el filtro dinámicamente
    lista_lugares = ["Todos los lugares"]
    if not df_meta.empty:
        # Obtenemos valores únicos, eliminamos vacíos y ordenamos
        lugares_unicos = sorted([str(l) for l in df_meta["Lugar"].unique() if l and str(l).strip() != ""])
        lista_lugares.extend(lugares_unicos)
    
    lugar_seleccionado = st.selectbox(
        "Selecciona un espacio o actividad para filtrar el calendario:", 
        lista_lugares,
        key="filtro_lugar_home"
    )
    
    st.divider()

    col_izq, col_der = st.columns([2, 1])
    
    with col_izq:
        st.markdown("### 🗓️ Calendario Institucional Dinámico")
        if lugar_seleccionado != "Todos los lugares":
            st.caption(f"👁️ *Mostrando únicamente eventos programados en: **{lugar_seleccionado}***")
        else:
            st.caption("💡 *Haz clic en cualquier día o evento para desplegar el resumen detallado en la parte inferior.*")
        
        # --- FILTRADO DE DATOS PARA EL CALENDARIO ---
        if lugar_seleccionado != "Todos los lugares":
            df_eventos_cal = df_meta[df_meta["Lugar"] == lugar_seleccionado]
        else:
            df_eventos_cal = df_meta

        eventos_list = []
        for idx, r in df_eventos_cal.iterrows():
            color = "#36a2eb"  # Color azul por defecto
            if r["Lugar"] == "Sala Situacional": color = "#2ecc71"     # Verde
            elif r["Lugar"] == "Auditorio Panzigua": color = "#e67e22" # Naranja
            elif r["Lugar"] == "UNIDAD DE ANALISIS": color = "#9b59b6" # Morado
            
            v_emoji = " 🚗" if r["Vehículo"] == "Sí" else ""
            eventos_list.append({
                "title": f"[{r['Hora Inicio']}] {r['Tipo de Evento']} - {r['Responsable']}{v_emoji}",
                "start": f"{r['Fecha']}T{r['Hora Inicio']}:00",
                "end": f"{r['Fecha']}T{r['Hora Fin']}:00",
                "backgroundColor": color, 
                "borderColor": color
            })
        
        # Capturamos la interacción del usuario con el calendario
        interaccion_cal = calendar(
            events=eventos_list, 
            options={"locale": "es", "headerToolbar": {"right": "dayGridMonth,timeGridWeek"}}, 
            key="cal_vsp_interactivo"
        )

    with col_der:
        st.markdown("### 🛡️ Disponibilidad Esta Semana")
        hoy_dt = datetime.today()
        lunes = hoy_dt - timedelta(days=hoy_dt.weekday())
        df_d = cargar_datos('Disponibilidad')
        reg_sem = df_d[df_d["Semana_Inicio"] == lunes.strftime("%Y-%m-%d")]
        
        if not reg_sem.empty:
            integrantes = str(reg_sem.iloc[0]["Integrantes"]).split(";")
            cargos = str(reg_sem.iloc[0]["Cargos"]).split(";")
            for i, nombre in enumerate(integrantes):
                st.info(f"👤 **{nombre}**  \n*{cargos[i]}*")
        else:
            st.warning("⚠️ Sin equipo de turno asignado para esta semana.")
            
        st.markdown("---")
        st.markdown("### 📝 Compromisos Recientes Pendientes")
        df_cv = cargar_datos('Compromisos')
        if not df_cv.empty:
            df_pendientes = df_cv[df_cv["Estado"].str.contains("PENDIENTE", na=False)].tail(5)
            if not df_pendientes.empty:
                st.dataframe(df_pendientes[["Compromiso", "Responsable", "Plazo"]], use_container_width=True, hide_index=True)
            else:
                st.success("🟢 ¡Todos los compromisos están al día!")
        else:
            st.info("No hay compromisos en la base de datos.")

    # =========================================================
    # VISOR DINÁMICO POR DÍA (DETALLE INFERIOR AL HACER CLICK)
    # =========================================================
    st.markdown("---")
    st.subheader("🔍 Agenda y Eventos del Día Seleccionado")
    
    # Por defecto, si no ha hecho clic, el sistema asume la fecha de hoy
    fecha_seleccionada = hoy.strftime("%Y-%m-%d")
    origen_click = "Hoy"

    # Validamos si el usuario interactuó con el calendario
    if interaccion_cal:
        if "eventClick" in interaccion_cal:
            fecha_cruda = interaccion_cal["eventClick"]["event"]["start"]
            fecha_seleccionada = fecha_cruda.split("T")[0]
            origen_click = "Evento Seleccionado"
        elif "dateClick" in interaccion_cal:
            fecha_cruda = interaccion_cal["dateClick"]["date"]
            fecha_seleccionada = fecha_cruda.split("T")[0]
            origen_click = "Día Seleccionado"

    try:
        fecha_f = datetime.strptime(fecha_seleccionada, "%Y-%m-%d").strftime("%d/%m/%Y")
    except:
        fecha_f = fecha_seleccionada

    st.markdown(f"📅 Visualizando agenda para el: **{fecha_f}** *(Filtro por: {origen_click})*")

    # Para el visor del día, respetamos TAMBIÉN el filtro de lugar que seleccionó el usuario arriba
    if lugar_seleccionado != "Todos los lugares":
        df_filtrado_dia = df_meta[(df_meta["Fecha"].astype(str) == fecha_seleccionada) & (df_meta["Lugar"] == lugar_seleccionado)]
    else:
        df_filtrado_dia = df_meta[df_meta["Fecha"].astype(str) == fecha_seleccionada]

    if not df_filtrado_dia.empty:
        df_filtrado_dia = df_filtrado_dia.sort_values(by="Hora Inicio")
        
        for _, fila in df_filtrado_dia.iterrows():
            transporte = "🚗 Requiere Vehículo Institucional" if fila["Vehículo"] == "Sí" else "🚶 Sin requerimiento de vehículo"
            
            with st.container():
                st.markdown(f"""
                <div style="padding:15px; border-radius:10px; background-color:rgba(30, 41, 59, 0.4); margin-bottom:12px; border-left: 5px solid #36a2eb;">
                    <h4 style="margin:0px; color:#38bdf8;">⏰ {fila['Hora Inicio']} - {fila['Hora Fin']} | {fila['Tipo de Evento']}</h4>
                    <p style="margin:4px 0px;">👤 <b>Responsable:</b> {fila['Responsable']}</p>
                    <p style="margin:4px 0px;">📍 <b>Lugar:</b> {fila['Lugar']} ({fila['Municipio']})</p>
                    <p style="margin:4px 0px; color:#cbd5e1;"><b>Observaciones:</b> {fila['Observaciones'] if fila['Observaciones'] else 'Ninguna'}</p>
                    <p style="margin:4px 0px; font-size:0.9em; color:#a7f3d0;"><b>Logística:</b> {transporte}</p>
                </div>
                """, unsafe_allow_html=True)
    else:
        st.info("🟢 No hay actividades programadas en el sistema para los criterios seleccionados en esta fecha.")

# --- PANTALLA 2: REGISTRAR ACTIVIDAD ---
elif st.session_state["seccion_actual"] == "📝 Registrar Actividad":
    st.markdown("### 📝 Formulario de Registro de Eventos")
    with st.form("form_reg", clear_on_submit=True):
        f_fecha = st.date_input("Fecha", value=datetime.today())
        c_h1, c_h2 = st.columns(2)
        with c_h1: f_hi = st.time_input("Hora Inicio", value=time(8, 0))
        with c_h2: f_hf = st.time_input("Hora Fin", value=time(10, 0))
        
        f_resp = st.selectbox("Responsable", LISTA_RESPONSABLES)
        f_tipo = st.selectbox("Tipo de Evento", LISTA_TIPOS_EVENTO)
        f_mun = st.selectbox("Municipio", LISTA_MUNICIPIOS)
        f_lugar = st.selectbox("Lugar", LISTA_LUGARES)
        f_veh = st.toggle("¿Requiere Vehículo?")
        f_obs = st.text_area("Observaciones")
        
        btn_guardar = st.form_submit_button("💾 Agendar y Guardar Actividad")

    if btn_guardar:
        cruce = False
        if f_lugar in ["Sala Situacional", "Auditorio Panzigua"]:
            eventos_cal = cargar_datos('Eventos')
            for _, e in eventos_cal.iterrows():
                if str(e["Fecha"]) == f_fecha.strftime("%Y-%m-%d") and e["Lugar"] == f_lugar:
                    if f_hi.strftime("%H:%M") < str(e["Hora Fin"]) and f_hf.strftime("%H:%M") > str(e["Hora Inicio"]):
                        cruce = True; break
        
        if cruce: 
            st.error(f"🚨 El espacio '{f_lugar}' ya está reservado en ese horario.")
        elif "Seleccione..." in [f_resp, f_mun, f_lugar]: 
            st.error("❌ Faltan campos obligatorios.")
        else:
            nuevo = pd.DataFrame([{"Fecha": f_fecha.strftime("%Y-%m-%d"), "Hora Inicio": f_hi.strftime("%H:%M"), "Hora Fin": f_hf.strftime("%H:%M"), "Responsable": f_resp, "Tipo de Evento": f_tipo, "Municipio": f_mun, "Lugar": f_lugar, "Vehículo": "Sí" if f_veh else "No", "Estado": "Programado", "Observaciones": f_obs}])
            df_act = cargar_datos('Eventos')
            guardar_datos(pd.concat([df_act, nuevo], ignore_index=True), 'Eventos')
            st.success("🎉 ¡Actividad registrada exitosamente!")
            st.session_state["seccion_actual"] = "🏠 Inicio / Resumen General"
            st.rerun()


# --- PANTALLA 3: DISPONIBILIDAD SEMANAL ---
elif st.session_state["seccion_actual"] == "🛡️ Disponibilidad Semanal":
    st.markdown("### 🛡️ Equipo de Disponibilidad por Semana Epidemiológica (S.E.)")
    
    # --- FUNCIONES AUXILIARES DE FECHAS Y S.E. ---
    def obtener_semana_epidemiologica(fecha):
        """Retorna el número de Semana Epidemiológica según el estándar."""
        return fecha.isocalendar()[1]

    def formatear_fecha_vsp(fecha_dt):
        """Formatea una fecha a 'LUNES 18 DE MAYO DE 2026' en mayúsculas."""
        meses = {
            1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
            7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
        }
        dias = {
            0: "LUNES", 1: "MARTES", 2: "MIÉRCOLES", 3: "JUEVES", 4: "VIERNES", 5: "SÁBADO", 6: "DOMINGO"
        }
        dia_nombre = dias[fecha_dt.weekday()]
        mes_nombre = meses[fecha_dt.month]
        return f"{dia_nombre} {fecha_dt.day} DE {mes_nombre} DE {fecha_dt.year}"

    def generar_semanas_del_mes(anio, mes):
        """Genera una lista de tuplas con las semanas que inician o caen en un mes específico."""
        semanas = []
        fecha_iter = datetime(anio, mes, 1)
        if mes == 12:
            ultimo_dia = datetime(anio, 12, 31)
        else:
            ultimo_dia = datetime(anio, mes + 1, 1) - timedelta(days=1)
            
        lunes_inicial = fecha_iter - timedelta(days=fecha_iter.weekday())
        
        while lunes_inicial <= ultimo_dia:
            domingo_final = lunes_inicial + timedelta(days=6)
            semana_epi = obtener_semana_epidemiologica(lunes_inicial)
            
            if (lunes_inicial, domingo_final, semana_epi) not in semanas:
                semanas.append((lunes_inicial, domingo_final, semana_epi))
            lunes_inicial += timedelta(days=7)
            
        return semanas

    # --- FECHAS DE REFERENCIA EN TIEMPO REAL ---
    hoy_dt = datetime.today()
    lunes_actual = hoy_dt - timedelta(days=hoy_dt.weekday())
    domingo_actual = lunes_actual + timedelta(days=6)
    semana_actual_epi = obtener_semana_epidemiologica(hoy_dt)

    # --- CARGA Y LIMPIEZA ABSOLUTA DE DATOS HISTÓRICOS ---
    df_raw = cargar_datos('Disponibilidad')
    
    df_d = df_raw.copy()
    for col in ["Semana_Inicio", "Integrantes", "Cargos", "Laboratorio_Responsable", "Laboratorio_Cargo"]:
        if col not in df_d.columns:
            df_d[col] = ""
        df_d[col] = df_d[col].astype(str).fillna("").replace("nan", "").str.strip()

    # Contenedor de visualización de la S.E. Actual en Curso
    st.markdown(f"""
    <div style="padding:15px; border-radius:10px; background-color:rgba(14, 116, 144, 0.2); border-left: 5px solid #06b6d4; margin-bottom:20px;">
        <h4 style="margin:0px; color:#22d3ee;">📆 SEMANA EPIDEMIOLÓGICA ACTUAL: S.E. {semana_actual_epi}</h4>
        <p style="margin:5px 0px 0px 0px; font-weight:bold; color:#cbd5e1;">
            Desde el {formatear_fecha_vsp(lunes_actual)} <br>Hasta el {formatear_fecha_vsp(domingo_actual)}
        </p>
    </div>
    """, unsafe_allow_html=True)

    # Buscar si existe registro cargado para la semana en curso en la vista pública
    fecha_hoy_str = lunes_actual.strftime("%Y-%m-%d")
    reg_sem = df_d[df_d["Semana_Inicio"] == fecha_hoy_str]
    
    if not reg_sem.empty:
        st.markdown("#### 👥 Equipo de Turno Asignado (Sala de Análisis de Riesgo)")
        integrantes = [i for i in str(reg_sem.iloc[0]["Integrantes"]).split(";") if i.strip()]
        cargos = [c for c in str(reg_sem.iloc[0]["Cargos"]).split(";") if c.strip()]
        
        if integrantes:
            cols = st.columns(len(integrantes))
            for i, nombre in enumerate(integrantes):
                with cols[i]:
                    st.info(f"👤 **{nombre}**\n\n💼 *{cargos[i] if i < len(cargos) else 'Asignado'}*")
        else:
            st.warning("⚠️ No hay integrantes de vigilancia registrados para esta semana.")
            
        lab_resp = reg_sem.iloc[0]["Laboratorio_Responsable"]
        lab_cargo = reg_sem.iloc[0]["Laboratorio_Cargo"]
        
        if lab_resp:
            st.markdown("---")
            st.markdown("#### 🔬 Turno Especial - Laboratorio de Salud Pública")
            st.success(f"🧫 **Responsable:** {lab_resp}  \n💼 **Función:** {lab_cargo if lab_cargo else 'Personal de Turno'}")
    else:
        st.warning("⚠️ No se ha asignado un equipo de guardia institucional para la actual Semana Epidemiológica.")

    # =========================================================
    # FILTRO OPTIMIZADO: CONSULTA DE DISPONIBILIDADES TODO 2026
    # =========================================================
    st.markdown("---")
    st.markdown("### 🔍 Consultar Programación de Próximas Semanas")
    st.caption("Utilice este filtro para verificar con antelación qué funcionarios estarán disponibles en cualquier semana restante del año 2026.")

    opciones_futuras = []
    mapeo_futuro = {}
    
    # Recorremos desde el mes actual hasta diciembre (mes 12) del año 2026
    for mes_futuro in range(hoy_dt.month, 13):
        semanas_calculadas = generar_semanas_del_mes(2026, mes_futuro)
        for lunes_f, domingo_f, num_sem_f in semanas_calculadas:
            # Filtramos para que solo aparezcan de la semana actual en adelante
            if lunes_f >= lunes_actual:
                txt_filtro = f"S.E. {num_sem_f} (Del {lunes_f.strftime('%d/%m/%Y')} al {domingo_f.strftime('%d/%m/%Y')})"
                if txt_filtro not in opciones_futuras:
                    opciones_futuras.append(txt_filtro)
                    mapeo_futuro[txt_filtro] = (lunes_f, num_sem_f)

    semana_consulta = st.selectbox("Seleccione una semana para ver los disponibles:", opciones_futuras, key="sb_consulta_futura")

    if semana_consulta:
        lunes_c, n_sem_c = mapeo_futuro[semana_consulta]
        fecha_c_str = lunes_c.strftime("%Y-%m-%d")
        
        # Consultar el Excel cargado
        registro_consulta = df_d[df_d["Semana_Inicio"] == fecha_c_str]
        
        if not registro_consulta.empty:
            st.markdown(f"#### 📅 Asignaciones registradas para la **S.E. {n_sem_c}**:")
            
            # Separar datos de Vigilancia
            int_c = [i for i in str(registro_consulta.iloc[0]["Integrantes"]).split(";") if i.strip()]
            carg_c = [c for c in str(registro_consulta.iloc[0]["Cargos"]).split(";") if c.strip()]
            
            st.markdown("**🛡️ Personal de Vigilancia (Sala de Análisis de Riesgo):**")
            if int_c:
                cols_c = st.columns(len(int_c))
                for idx_c, nom_c in enumerate(int_c):
                    with cols_c[idx_c]:
                        st.warning(f"👤 **{nom_c}**\n\n💼 *{carg_c[idx_c] if idx_c < len(carg_c) else 'Disponible'}*")
            else:
                st.info("No se detallaron nombres de integrantes para esta semana.")
                
            # Separar datos de Laboratorio
            lab_resp_c = registro_consulta.iloc[0]["Laboratorio_Responsable"]
            lab_cargo_c = registro_consulta.iloc[0]["Laboratorio_Cargo"]
            
            if lab_resp_c:
                st.markdown("**🔬 Apoyo de Laboratorio de Salud Pública:**")
                st.code(f"🔬 RESPONSABLE: {lab_resp_c}\n💼 FUNCIÓN/CARGO: {lab_cargo_c if lab_cargo_c else 'Asignado'}", language="text")
        else:
            st.error(f"📋 Aún no se han cargado datos de disponibilidad para la Semana Epidemiológica {n_sem_c}.")

    st.divider()
    
    # =========================================================
    # ZONA ADMINISTRATIVA: PANEL DE GESTIÓN Y PLANIFICACIÓN
    # =========================================================
    st.subheader("🔐 Panel de Configuración y Gestión de Disponibilidad")
    clave_dispo = st.text_input("Ingrese la Llave de Seguridad para habilitar el gestor de turnos:", type="password", key="pwd_dispo")
    
    if clave_dispo == CONTRASENA_ADMIN:
        st.success("🔓 Autenticación correcta: Modo Administrador Habilitado.")
        
        st.markdown("#### 📅 Selección de Cronograma Epidemiológico")
        
        meses_nombres = {
            1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
            7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
        }
        
        opciones_meses = []
        mes_ref = hoy_dt.month
        anio_ref = hoy_dt.year
        
        # Mantenemos la generación de meses en el panel de control del administrador para agendar adelante
        for m in range(6):
            mes_calculado = mes_ref + m
            anio_calculado = anio_ref
            if mes_calculado > 12:
                mes_calculado -= 12
                anio_calculado += 1
            opciones_meses.append(f"{meses_nombres[mes_calculado].upper()} {anio_calculado}")
        
        mes_seleccionado_texto = st.selectbox("Seleccione el mes a programar:", opciones_meses, key="sb_mes_dispo")
        
        nombre_mes_sel = mes_seleccionado_texto.split(" ")[0]
        anio_sel = int(mes_seleccionado_texto.split(" ")[1])
        mes_sel = [k for k, v in meses_nombres.items() if v.upper() == nombre_mes_sel][0]
        
        lista_semanas_mes = generar_semanas_del_mes(anio_sel, mes_sel)
        
        opciones_semanas_select = []
        mapeo_fechas_semanas = {}
        
        for lunes_i, domingo_f, num_semana in lista_semanas_mes:
            texto_semana = f"S.E. {num_semana} (Del {lunes_i.strftime('%d/%m')} al {domingo_f.strftime('%d/%m')})"
            opciones_semanas_select.append(texto_semana)
            mapeo_fechas_semanas[texto_semana] = (lunes_i, domingo_f, num_semana)
        
        semana_seleccionada_texto = st.selectbox("Seleccione la Semana Epidemiológica (S.E.) a configurar o modificar:", opciones_semanas_select, key="sb_semana_dispo")
        
        if semana_seleccionada_texto:
            lunes_elegido, domingo_elegido, n_semana_elegida = mapeo_fechas_semanas[semana_seleccionada_texto]
            fecha_clave_str = lunes_elegido.strftime("%Y-%m-%d")
            
            st.markdown("**Rango Oficial Validado:**")
            st.code(f"DESDE: {formatear_fecha_vsp(lunes_elegido)}\nHASTA: {formatear_fecha_vsp(domingo_elegido)}", language="text")
            
            registro_existente = df_d[df_d["Semana_Inicio"] == fecha_clave_str]
            
            int_previos = []
            carg_previos = []
            lab_resp_previo = ""
            lab_cargo_previo = ""
            
            if not registro_existente.empty:
                st.info("ℹ️ Esta S.E. ya contiene un registro. Al presionar guardar, se actualizarán los datos existentes sin duplicar.")
                int_previos = [i for i in str(registro_existente.iloc[0]["Integrantes"]).split(";") if i.strip()]
                carg_previos = [c for c in str(registro_existente.iloc[0]["Cargos"]).split(";") if c.strip()]
                lab_resp_previo = registro_existente.iloc[0]["Laboratorio_Responsable"]
                lab_cargo_previo = registro_existente.iloc[0]["Laboratorio_Cargo"]
            
            st.markdown("---")
            st.markdown(f"#### 📝 Registro de Turnos para la S.E. {n_semana_elegida}")
            
            num_integrantes = st.slider("Asignar cantidad de funcionarios para Vigilancia (Sala de Análisis):", 1, 9, len(int_previos) if int_previos else 2, key="slider_num_int")
            
            nuevos_integrantes = []
            nuevos_cargos = []
            
            for idx in range(num_integrantes):
                st.markdown(f"**🔹 Funcionario de Vigilancia {idx + 1}**")
                c1, c2 = st.columns(2)
                
                idx_def_resp = 0
                if int_previos and idx < len(int_previos):
                    if int_previos[idx] in LISTA_RESPONSABLES:
                        idx_def_resp = LISTA_RESPONSABLES.index(int_previos[idx])
                
                val_cargo_def = ""
                if carg_previos and idx < len(carg_previos):
                    val_cargo_def = carg_previos[idx]
                    
                with c1:
                    nuevos_integrantes.append(st.selectbox(f"Nombre del Responsable {idx+1}", LISTA_RESPONSABLES, index=idx_def_resp, key=f"se_resp_{idx}_{n_semana_elegida}"))
                with c2:
                    nuevos_cargos.append(st.text_input(f"Cargo / Función {idx+1}", value=val_cargo_def, placeholder="Ej: Epidemiólogo", key=f"se_cargo_{idx}_{n_semana_elegida}"))
            
            st.markdown("---")
            st.markdown("#### 🔬 Asignación del Laboratorio de Salud Pública")
            st.caption("Registre manualmente el personal asignado por el laboratorio para el apoyo analítico de la semana.")
            
            c_lab1, c_lab2 = st.columns(2)
            with c_lab1:
                lab_responsable_input = st.text_input("Nombre completo del Responsable (Laboratorio):", value=lab_resp_previo, placeholder="Ej: Dra. María Pérez", key="input_lab_responsable")
            with c_lab2:
                lab_cargo_input = st.text_input("Cargo / Componente (Laboratorio):", value=lab_cargo_previo, placeholder="Ej: Profesional Universitario", key="input_lab_cargo")
            
            st.markdown("---")
            
            if st.button("💾 GUARDAR TODOS LOS CAMBIOS DE LA SEMANA", type="primary", use_container_width=True, key="btn_guardar_dispo_def"):
                if any(x == "Seleccione..." for x in nuevos_integrantes) or any(c.strip() == "" for c in nuevos_cargos):
                    st.error("❌ Todos los campos de nombres y cargos activos para la Sala de Análisis de Riesgos deben estar completos.")
                else:
                    datos_nueva_fila = {
                        "Semana_Inicio": str(fecha_clave_str),
                        "Integrantes": ";".join(nuevos_integrantes),
                        "Cargos": ";".join(nuevos_cargos),
                        "Laboratorio_Responsable": str(lab_responsable_input.strip()),
                        "Laboratorio_Cargo": str(lab_cargo_input.strip())
                    }
                    
                    df_limpio = df_d[df_d["Semana_Inicio"] != fecha_clave_str].copy()
                    df_insertar = pd.DataFrame([datos_nueva_fila])
                    df_final_guardar = pd.concat([df_limpio, df_insertar], ignore_index=True)
                    df_final_guardar = df_final_guardar.loc[:, ~df_final_guardar.columns.str.contains('^Unnamed')]
                    
                    guardar_datos(df_final_guardar, 'Disponibilidad')
                    
                    st.success(f"🎉 ¡Éxito! Turno de Vigilancia y Laboratorio para la S.E. {n_semana_elegida} consolidado en el Excel.")
                    st.rerun()
                    
    elif clave_dispo != "":
        st.error("🔑 Llave incorrecta. Modificación bloqueada.")
# --- PANTALLA 4: COMPROMISOS TÉCNICOS (CON EDICIÓN ABIERTA Y ELIMINACIÓN CON LLAVE) ---
elif st.session_state["seccion_actual"] == "📝 Compromisos Técnicos":
    st.markdown("### 📝 Gestión Avanzada de Compromisos y Actas")
    
    CARPETA_SOPORTES = "soportes_compromisos"
    if not os.path.exists(CARPETA_SOPORTES):
        os.makedirs(CARPETA_SOPORTES)

    df_cv = cargar_datos('Compromisos')
    if "Respuesta_Avance" not in df_cv.columns:
        df_cv["Respuesta_Avance"] = ""
    if "Ruta_Soporte" not in df_cv.columns:
        df_cv["Ruta_Soporte"] = ""

    # --- 1. VISUALIZACIÓN DE LA MATRIZ PÚBLICA ---
    st.subheader("📋 Matriz de Seguimiento General")
    if not df_cv.empty:
        df_publico = df_cv[["Fecha_Acuerdo", "Responsable", "Compromiso", "Plazo", "Estado"]].copy()
        st.dataframe(df_publico, use_container_width=True, hide_index=True)
    else:
        st.info("No hay compromisos registrados en la base de datos.")

    st.divider()

    # --- 2. CONTROL DE ACCESO ÚNICO PARA EL ADMINISTRADOR ---
    st.subheader("🔐 Panel de Control del Administrador")
    clave_maestra = st.text_input("Ingrese la Llave de Seguridad para desbloquear funciones de Administrador:", type="password", key="pwd_maestra_compromisos")
    
    if clave_maestra == CONTRASENA_ADMIN:
        st.success("🔓 Autenticación Exitosa: Modo Administrador Activo")
        
        # --- A. VER RESPUESTAS Y DESCARGAR ACTAS ---
        st.markdown("#### 👁️ Respuestas Detalladas y Soportes")
        if not df_cv.empty:
            st.dataframe(df_cv[["Fecha_Acuerdo", "Responsable", "Compromiso", "Estado", "Respuesta_Avance", "Ruta_Soporte"]], use_container_width=True, hide_index=True)
            
            df_con_soporte = df_cv[df_cv["Ruta_Soporte"] != ""]
            if not df_con_soporte.empty:
                st.markdown("##### 📁 Descarga de Documentos Adjuntos")
                for idx, fila in df_con_soporte.iterrows():
                    nombre_archivo = fila["Ruta_Soporte"]
                    ruta_completa = os.path.join(CARPETA_SOPORTES, nombre_archivo)
                    
                    if os.path.exists(ruta_completa):
                        with open(ruta_completa, "rb") as f:
                            bytes_archivo = f.read()
                        
                        col_info, col_btn = st.columns([3, 1])
                        col_info.write(f"📄 **ID {idx}** ({fila['Responsable']}): {nombre_archivo}")
                        col_btn.download_button(label="⬇️ Descargar Soporte", data=bytes_archivo, file_name=nombre_archivo, key=f"dl_{idx}")
            else:
                st.info("No hay archivos adjuntos en el sistema.")
        
        st.divider()
        
        # --- B. ELIMINACIÓN DE REGISTROS (Segura usando Índices Reales) ---
        st.markdown("#### 🚨 Zona de Eliminación Definitiva")
        if not df_cv.empty:
            opciones_eliminar = ["Seleccione..."]
            for idx, fila in df_cv.iterrows():
                opciones_eliminar.append(f"{idx} - [{fila['Fecha_Acuerdo']}] {fila['Responsable']}: {str(fila['Compromiso'])[:50]}...")
            
            seleccion_eliminar = st.selectbox("Seleccione qué compromiso desea borrar permanentemente:", opciones_eliminar, key="sb_eliminar_directo")
            
            if seleccion_eliminar != "Seleccione...":
                idx_eliminar = int(seleccion_eliminar.split(" - ")[0])
                
                if st.button("❌ ELIMINAR REGISTRO SELECCIONADO", type="primary", use_container_width=True):
                    archivo_a_borrar = df_cv.loc[idx_eliminar, "Ruta_Soporte"]
                    if archivo_a_borrar:
                        ruta_archivo_borrar = os.path.join(CARPETA_SOPORTES, archivo_a_borrar)
                        if os.path.exists(ruta_archivo_borrar):
                            os.remove(ruta_archivo_borrar)
                    
                    df_final_c = df_cv.drop(idx_eliminar) # Eliminación directa por etiqueta de índice
                    guardar_datos(df_final_c, 'Compromisos')
                    st.success("💥 Registro borrado de manera irreversible.")
                    st.rerun()
        else:
            st.info("No hay compromisos para eliminar.")
            
    elif clave_maestra != "":
        st.error("🔑 Llave incorrecta. Las funciones de visualización de actas y eliminación están bloqueadas.")

    st.divider()

    # --- 3. ACCIONES GENERALES: CREAR O EDITAR COMPROMISOS ---
    col_crear, col_responder = st.columns(2)
    
    with col_crear:
        st.markdown("#### 📌 Asignar Nuevo Compromiso")
        with st.form("form_nuevo_compromiso", clear_on_submit=True):
            c_acuerdo = st.text_area("Descripción del compromiso:")
            c_resp = st.selectbox("Responsable Asignado:", LISTA_RESPONSABLES, key="nuevo_c_resp")
            c_plazo = st.date_input("Fecha Límite:")
            
            if st.form_submit_button("💾 Registrar Compromiso"):
                if c_acuerdo.strip() == "" or c_resp == "Seleccione...":
                    st.error("❌ Por favor completa la descripción y selecciona un responsable.")
                else:
                    nuevo_c = pd.DataFrame([{
                        "Fecha_Acuerdo": datetime.now().strftime("%Y-%m-%d"), 
                        "Compromiso": c_acuerdo, 
                        "Responsable": c_resp, 
                        "Plazo": c_plazo.strftime("%Y-%m-%d"), 
                        "Estado": "🔴 PENDIENTE",
                        "Respuesta_Avance": "",
                        "Ruta_Soporte": ""
                    }])
                    guardar_datos(pd.concat([df_cv, nuevo_c], ignore_index=True), 'Compromisos')
                    st.success("🎉 Compromiso asignado con éxito.")
                    st.rerun()

    with col_responder:
        st.markdown("#### 🔄 Editar Compromiso y Subir Respuesta / Acta")
        if not df_cv.empty:
            opciones_compromisos = []
            for idx, fila in df_cv.iterrows():
                opciones_compromisos.append(f"{idx} - [{fila['Responsable']}] {str(fila['Compromiso'])[:40]}...")
            
            # El selector se queda afuera para refrescar dinámicamente los campos inferiores
            seleccionado = st.selectbox("Seleccione el compromiso que desea editar o responder:", opciones_compromisos, key="sb_editar_compromiso")
            idx_seleccionado = int(seleccionado.split(" - ")[0])
            
            # USO DE .loc PARA SEGURIDAD DEL ÍNDICE REAL
            fila_actual = df_cv.loc[idx_seleccionado]
            
            # Quitamos el componente 'with st.form' para la edición directa para evitar congelamiento de datos de Streamlit
            st.markdown(f"**Modificando Registro ID:** `{idx_seleccionado}`")
            edit_compromiso = st.text_area("Editar descripción del compromiso:", value=fila_actual['Compromiso'], key=f"txt_{idx_seleccionado}")
            
            idx_resp_def = LISTA_RESPONSABLES.index(fila_actual['Responsable']) if fila_actual['Responsable'] in LISTA_RESPONSABLES else 0
            edit_resp = st.selectbox("Reasignar Responsable:", LISTA_RESPONSABLES, index=idx_resp_def, key=f"resp_{idx_seleccionado}")
            
            try:
                fecha_def = datetime.strptime(str(fila_actual['Plazo']), "%Y-%m-%d").date()
            except:
                fecha_def = datetime.today().date()
            edit_plazo = st.date_input("Modificar Fecha Límite:", value=fecha_def, key=f"fecha_{idx_seleccionado}")
            
            st.markdown("---")
            nuevo_estado = st.selectbox("Actualizar Estado:", ["🔴 PENDIENTE", "🟢 CUMPLIDO"], index=0 if fila_actual['Estado'] == "🔴 PENDIENTE" else 1, key=f"est_{idx_seleccionado}")
            nueva_respuesta = st.text_area("Escribir Respuesta / Avance realizado:", value=fila_actual['Respuesta_Avance'], key=f"resp_av_{idx_seleccionado}")
            archivo_subido = st.file_uploader("Adjuntar Acta o Soporte (PDF, EXCEL):", type=["pdf", "xlsx", "xls"], key=f"file_{idx_seleccionado}")
            
            if st.button("✅ Guardar Cambios y Archivo", use_container_width=True, key=f"btn_save_{idx_seleccionado}"):
                nombre_archivo_final = fila_actual['Ruta_Soporte']
                if archivo_subido is not None:
                    nombre_archivo_final = f"compromiso_{idx_seleccionado}_{archivo_subido.name}"
                    ruta_guardado = os.path.join(CARPETA_SOPORTES, nombre_archivo_final)
                    with open(ruta_guardado, "wb") as f:
                        f.write(archivo_subido.getbuffer())
                
                # Guardado explícito y exacto en la celda correspondiente
                df_cv.at[idx_seleccionado, "Compromiso"] = edit_compromiso
                df_cv.at[idx_seleccionado, "Responsable"] = edit_resp
                df_cv.at[idx_seleccionado, "Plazo"] = edit_plazo.strftime("%Y-%m-%d")
                df_cv.at[idx_seleccionado, "Estado"] = nuevo_estado
                df_cv.at[idx_seleccionado, "Respuesta_Avance"] = nueva_respuesta
                df_cv.at[idx_seleccionado, "Ruta_Soporte"] = nombre_archivo_final
                
                guardar_datos(df_cv, 'Compromisos')
                st.success("🎉 Compromiso modificado correctamente.")
                st.rerun()
        else:
            st.info("No hay compromisos disponibles para editar.")
# --- PANTALLA 5: FILTROS Y DASHBOARD ---
elif st.session_state["seccion_actual"] == "🔍 Filtros y Dashboard":
    st.subheader("🔍 Consultas y Dashboard Gerencial")
    df_busc = cargar_datos('Eventos')
    
    if not df_busc.empty:
        col_f1, col_f2, col_f3 = st.columns([2, 1, 1])
        with col_f1:
            txt = st.text_input("Buscador por texto (Responsable, Lugar, etc.)", key="search_txt")
        with col_f2:
            f_mun = st.selectbox("Filtrar por Municipio", ["Todos"] + sorted(LISTA_MUNICIPIOS[1:]), key="filter_mun")
        with col_f3:
            f_tipo = st.selectbox("Filtrar por Tipo", ["Todos"] + sorted(LISTA_TIPOS_EVENTO), key="filter_tipo")

        df_f = df_busc.copy()
        if f_mun != "Todos": df_f = df_f[df_f["Municipio"] == f_mun]
        if f_tipo != "Todos": df_f = df_f[df_f["Tipo de Evento"] == f_tipo]
        if txt: df_f = df_f[df_f.astype(str).apply(lambda x: x.str.contains(txt, case=False, na=False)).any(axis=1)]

        st.markdown(f"**Registros encontrados:** {len(df_f)}")
        st.dataframe(df_f, use_container_width=True, hide_index=True)
        
        if not df_f.empty:
            st.divider()
            c_g1, c_g2 = st.columns(2)
            with c_g1:
                st.markdown("##### 📍 Actividades por Municipio")
                data_mun = df_f["Municipio"].value_counts()
                if not data_mun.empty: st.bar_chart(data_mun, color="#38bdf8")
            with c_g2:
                st.markdown("##### 📋 Volumen por Tipo de Evento")
                data_tipo = df_f["Tipo de Evento"].value_counts()
                if not data_tipo.empty: st.bar_chart(data_tipo, color="#deff9a")
    else:
        st.info("👋 Aún no hay actividades registradas.")


# --- PANTALLA 6: PANEL MAESTRO ---
elif st.session_state["seccion_actual"] == "⚙️ Panel Maestro":
    pwd = st.text_input("Llave de Seguridad", type="password")
    if pwd == CONTRASENA_ADMIN:
        st.success("🔓 Acceso Administrativo")
        df_admin = cargar_datos('Eventos')
        if not df_admin.empty:
            df_admin['ID_Fila'] = range(len(df_admin))
            df_admin['Detalle'] = df_admin['Fecha'].astype(str) + " - " + df_admin['Responsable']
            seleccion = st.selectbox("Seleccione registro para eliminar:", ["Seleccione..."] + df_admin['Detalle'].tolist())
            if seleccion != "Seleccione...":
                if st.button("🚨 ELIMINAR PERMANENTEMENTE"):
                    idx = df_admin[df_admin['Detalle'] == seleccion]['ID_Fila'].values[0]
                    df_final = df_admin.drop(df_admin.index[idx]).drop(columns=['ID_Fila', 'Detalle'])
                    guardar_datos(df_final, 'Eventos')
                    st.success("Registro eliminado correctamente.")
                    st.rerun()