import streamlit as st
from streamlit_calendar import calendar
import pandas as pd
from datetime import datetime, time
import io
import os

# Configuración de la página (Ancho completo)
st.set_page_config(layout="wide", page_title="Calendario de Eventos", page_icon="📅")

st.title("📅 Sistema de Planificación Actividades - VSP")
st.markdown("---")

# Nombre del archivo de persistencia local
ARCHIVO_DB = "base_datos_eventos.xlsx"
CONTRASENA_ADMIN = "Vsp26"  # <-- CAMBIA TU CONTRASEÑA AQUÍ SI LO DESEAS

# ==========================================
# FUNCIONES DE ALMACENAMIENTO (PERSISTENCIA)
# ==========================================
def cargar_datos_desde_excel():
    """Lee el archivo Excel local y lo convierte al formato que necesita el calendario"""
    if not os.path.exists(ARCHIVO_DB):
        return []
        
    try:
        df = pd.read_excel(ARCHIVO_DB)
        df = df.fillna("")
        
        eventos_cargados = []
        for index, row in df.iterrows():
            # DETERMINACIÓN DE COLORES PRIORITARIOS POR LUGAR O TIPO
            if row["Lugar"] == "Sala Situacional":
                color_evento = "#2ecc71"  # Verde Esmeralda
            elif row["Lugar"] == "Auditorio Panzigua":
                color_evento = "#e67e22"  # Naranja Institutional
            elif row["Lugar"] == "UNIDAD DE ANALISIS" or row["Tipo de Evento"] == "UNIDAD DE ANALISIS":
                color_evento = "#9b59b6"  # Púrpura Distintivo para Unidad de Análisis
            else:
                colores = {
                    "ASISTENCIA TECNICA": "#4bc0c0", "BAC": "#36a2eb", "BAI": "#ff6384",
                    "CAPACITACION": "#36a2eb", "COMITÉ ESTADISTICAS VITALES": "#ffcd56",
                    "COMITÉ SANIDAD PORTUARIA": "#ffcd56", "COVE": "#ff9f40", "IEC": "#9966ff",
                    "MESA DE TRABAJO": "#4bc0c0", "MONITOREO": "#c9cbcf", "REUNION": "#ff9f40",
                    "SAR": "#ff6384", "SEGUIMIENTO": "#c9cbcf", "OTRO": "#c9cbcf"
                }
                color_evento = colores.get(row["Tipo de Evento"], "#36a2eb")

            fecha_str = str(row["Fecha"]).split(" ")[0]
            start_iso = f"{fecha_str}T{row['Hora Inicio']}:00"
            end_iso = f"{fecha_str}T{row['Hora Fin']}:00"
            
            str_tiempo = f"{row['Hora Inicio']}-{row['Hora Fin']}"
            titulo = f"[{str_tiempo}] {row['Tipo de Evento']} en {row['Lugar']} - {row['Municipio']} ({row['Responsable']})"

            eventos_cargados.append({
                "id": str(index),  # Asignamos el índice de la fila como ID único del evento
                "title": titulo,
                "start": start_iso,
                "end": end_iso,
                "backgroundColor": color_evento,
                "borderColor": color_evento,
                "extendedProps": {
                    "responsable": str(row["Responsable"]),
                    "tipo": str(row["Tipo de Evento"]),
                    "modalidad": str(row["Modalidad"]),
                    "municipio": str(row["Municipio"]),
                    "zona": str(row["Zona"]),
                    "lugar": str(row["Lugar"]),
                    "hora_i": str(row["Hora Inicio"]),
                    "hora_f": str(row["Hora Fin"]),
                    "vehiculo": str(row["Vehículo"]),
                    "publico": str(row["Público Objetivo"]),
                    "estado": str(row["Estado"]),
                    "observaciones": str(row.get("Observaciones", ""))
                }
            })
        return eventos_cargados
    except Exception as e:
        st.error(f"Error al leer la base de datos local: {e}")
        return []

def guardar_evento_en_excel(nuevo_registro):
    """Agrega un registro nuevo al archivo Excel permanente"""
    if os.path.exists(ARCHIVO_DB):
        df_actual = pd.read_excel(ARCHIVO_DB)
        df_nuevo = pd.DataFrame([nuevo_registro])
        df_final = pd.concat([df_actual, df_nuevo], ignore_index=True)
    else:
        df_final = pd.DataFrame([nuevo_registro])
    df_final.to_excel(ARCHIVO_DB, index=False)

def eliminar_evento_por_indice(indice_eliminar):
    """Elimina una fila específica del archivo Excel usando su número de fila"""
    if os.path.exists(ARCHIVO_DB):
        df = pd.read_excel(ARCHIVO_DB)
        if 0 <= indice_eliminar < len(df):
            df = df.drop(df.index[indice_eliminar])
            df.to_excel(ARCHIVO_DB, index=False)
            return True
    return False

# Cargar eventos históricos directamente al estado de la sesión
if "eventos_lista" not in st.session_state:
    st.session_state.eventos_lista = cargar_datos_desde_excel()

# ==========================================
# CONFIGURACIÓN DE LISTAS DESPLEGABLES
# ==========================================
LISTA_RESPONSABLES = ["Seleccione...", "ADALGISA PATRON CONDE", "ANA GABRIELA DIAZ ANAYA", "ANA LUCIA MENDOZA TAMARA", "BALDIR PABA OSORIO", "DIANA PAOLA PALENCIA SÁNCHEZ", "EDER JESUS PATERNINA RODRIGUEZ" , "ELIANA CECILIA MORALES MELENDEZ" , "ENITT DEL ROSARIO HERNANDEZ DORIAS" , "ESPERANZA DEL PILAR VARGAS VARGAS" , "HECTOR FABIO RENTERIA" , "ISAAC JACOB VELASQUEZ DOMINGUEZ" , "JAVIER MAURICIO CORREA PATERNOSTRO" , "KAREN MARGARITA ALDANA ARRIETA" , "KEVIN ALBERTO BARBARAN ALVAREZ" , "LEVY SUNILDA CAMPO LASSO" , "LOLI LUZ SIERRA DIAZ" , "LORENA PORTILLO CUENTAS" , "LUCIA CLARETH HERNANDEZ PEREZ" , "LUISA FERNANDA REYES DÍAZ" , "LUZMILA VILLAMIZAR MOLINA" , "MARIA CANDELARIA MEJIA LOPEZ" , "MARIA JOSE CANTILLO ROYERO" , "MARLON ESPITIA CERPA" , "MARTHA CECILIA MELENDEZ MARTINEZ" , "MERY DE JESUS NARVAEZ ASSIA" , "NICOLASA MARGARITA ARRIETA SERPA" , "NURYS CONCEPCIÓN HERRERA GUTIÉRREZ" , "VIRGINIA OLIVERO GARCIA" , "YARLENY ESTHER BERRIO ACOSTA" , "MARIA JOSE PEÑARANDA" , "BRENDER BARRIOS" , "ANA KARINA PEÑATES DE ARCE" , "DINO VERGARA PEREZ" , "JUAN CARLOS GARCIA VIVERO" , "MANUEL ORTEGA HERNANDEZ" , "MARIA CAMPO" , "VILMA MERCADO CUMPLIDO"]
LISTA_MUNICIPIOS = ["Seleccione...", "Buenavista", "Caimito", "Chalán", "Colosó", "Corozal", "Coveñas", "El Roble", "Galeras", "Guaranda", "La Unión", "Los Palmitos", "Majagual", "Morroa", "Ovejas", "Palmito", "Sampués", "San Benito Abad", "San Juan de Betulia", "San Marcos", "San Onofre", "San Pedro", "Sincé", "Sincelejo", "Sucre", "Tolú", "Toluviejo"]
LISTA_LUGARES = ["Seleccione...", "Sala Situacional", "Auditorio Panzigua", "UNIDAD DE ANALISIS", "Otro"]
LISTA_TIPOS_EVENTO = ["ASISTENCIA TECNICA", "BAC", "BAI", "CAPACITACION", "COMITÉ ESTADISTICAS VITALES", "COMITÉ SANIDAD PORTUARIA", "COVE" , "IEC" , "MESA DE TRABAJO" , "MONITOREO" , "REUNION" , "SAR" , "SEGUIMIENTO" , "UNIDAD DE ANALISIS" , "OTRO"]

# ==========================================
# COLUMNA IZQUIERDA: FORMULARIO DE INGRESO
# ==========================================
col_form, col_cal = st.columns([1, 2])

with col_form:
    # --- SISTEMA DE AUTENTICACIÓN ---
    with st.expander("🔑 Acceso Restringido (Gestión de Agenda)"):
        password_input = st.text_input("Ingrese la Llave Maestra:", type="password")
        es_admin = (password_input == CONTRASENA_ADMIN)
        if es_admin:
            st.success("🔓 Modo Administrador Activo. Puedes eliminar eventos.")
            
    st.subheader("📝 Registrar Nuevo Evento")
    
    with st.form("formulario_evento", clear_on_submit=True):
        fecha = st.date_input("Fecha del Evento", value=datetime.today())
        
        col_h1, col_h2 = st.columns(2)
        with col_h1:
            hora_inicio = st.time_input("Hora de Inicio", value=time(8, 0))
        with col_h2:
            hora_final = st.time_input("Hora Final", value=time(10, 0))
            
        responsable = st.selectbox("Responsable / Coordinador", LISTA_RESPONSABLES)
        tipo_evento = st.selectbox("Tipo de Evento", LISTA_TIPOS_EVENTO)
        modalidad = st.radio("Modalidad", ["Presencial", "Virtual", "Salida Campo"], horizontal=True)
        municipio = st.selectbox("Municipio (Sucre)", LISTA_MUNICIPIOS)
        zona = st.radio("Zona", ["Urbana", "Rural"], horizontal=True)
        
        lugar_seleccionado = st.selectbox("Lugar del Evento", LISTA_LUGARES)
        lugar_especifico_otro = st.text_input("Si seleccionó 'Otro', especifique cuál:", placeholder="Ej. Aula Hospital")
        
        requiere_vehiculo = st.toggle("¿Requiere Vehículo?")
        publico_objetivo = st.text_input("Público Objetivo / Dirigido a:", placeholder="Ej. Personal de salud")
        estado_evento = st.selectbox("Estado Inicial", ["Programado", "Confirmado", "Pendiente por Confirmar"])
        observaciones = st.text_area("Observaciones adicionales")
        
        guardar = st.form_submit_button("💾 Guardar Evento")

    if guardar:
        lugar_final = lugar_seleccionado
        if lugar_seleccionado == "Otro":
            lugar_final = lugar_especifico_otro

        h_inicio_str = hora_inicio.strftime('%H:%M')
        h_final_str = hora_final.strftime('%H:%M')
        fecha_actual_str = fecha.strftime("%Y-%m-%d")
        fecha_legible = fecha.strftime("%d/%m/%Y")

        cruce_detectado = False
        nombre_cruce = ""
        
        if lugar_final in ["Sala Situacional", "Auditorio Panzigua", "UNIDAD DE ANALISIS"]:
            for ev in st.session_state.eventos_lista:
                p = ev["extendedProps"]
                if ev["start"].split("T")[0] == fecha_actual_str and p["lugar"] == lugar_final:
                    if h_inicio_str < p["hora_f"] and h_final_str > p["hora_i"]:
                        cruce_detectado = True
                        nombre_cruce = f"{ev['title']}"
                        break

        if hora_final <= hora_inicio:
            st.error("❌ La hora final no puede ser menor o igual a la hora de inicio.")
        elif responsable == "Seleccione..." or municipio == "Seleccione..." or lugar_seleccionado == "Seleccione...":
            st.error("❌ Por favor, seleccione opciones válidas.")
        elif lugar_seleccionado == "Otro" and not lugar_especifico_otro:
            st.error("❌ Escriba el nombre del lugar alternativo.")
        elif cruce_detectado:
            st.error(f"""
            🚨 **¡CHOQUE DE HORARIO DETECTADO!**  
            La fecha seleccionada (**{fecha_legible}**) en el horario de **{h_inicio_str} a {h_final_str}** ya se encuentra programada para ese espacio.  
            
            * **Lugar afectado:** {lugar_final}  
            * **Evento ocupante:** {nombre_cruce}  
            
            ⚠️ **Por favor, cambie la hora o la fecha para poder registrar el evento.**
            """)
        else:
            nuevo_registro_excel = {
                "Fecha": fecha_actual_str,
                "Hora Inicio": h_inicio_str,
                "Hora Fin": h_final_str,
                "Responsable": responsable,
                "Tipo de Evento": tipo_evento,
                "Modalidad": modalidad,
                "Municipio": municipio,
                "Zona": zona,
                "Lugar": lugar_final,
                "Vehículo": "Sí" if requiere_vehiculo else "No",
                "Público Objetivo": publico_objetivo,
                "Estado": estado_evento,
                "Observaciones": observaciones
            }
            
            guardar_evento_en_excel(nuevo_registro_excel)
            st.session_state.eventos_lista = cargar_datos_desde_excel()
            st.success("¡Evento validado y guardado permanentemente!")
            st.rerun()

# ==========================================
# COLUMNA DERECHA: CALENDARIO DINÁMICO
# ==========================================
with col_cal:
    st.subheader("📆 Vista de Actividades")
    
    calendar_options = {
        "headerToolbar": {"left": "prev,next today", "center": "title", "right": "dayGridMonth,timeGridWeek,timeGridDay"},
        "initialView": "dayGridMonth",
        "locale": "es",
        "selectable": True,
        "navLinks": True,
        "slotMinTime": "06:00:00", "slotMaxTime": "20:00:00",
        "allDaySlot": False
    }
    
    state = calendar(events=st.session_state.eventos_lista, options=calendar_options, key="calendar")
    
    if state.get("eventClick"):
        st.markdown("---")
        st.subheader("🔍 Detalles del Evento Seleccionado")
        evento_click = state["eventClick"]["event"]
        props = evento_click.get("extendedProps", {})
        id_evento = evento_click.get("id") # Captura la fila real en el Excel

        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**📌 Título:** {evento_click.get('title')}")
            st.markdown(f"**⏰ Horario:** {props.get('hora_i')} a {props.get('hora_f')}")
            st.markdown(f"**👤 Responsable:** {props.get('responsable')}")
            st.markdown(f"**🏷️ Tipo:** {props.get('tipo')}")
            st.markdown(f"**📍 Municipio:** {props.get('municipio')} ({props.get('zona')})")
        with col2:
            st.markdown(f"**🏢 Lugar:** {props.get('lugar')}")
            st.markdown(f"**💻 Modalidad:** {props.get('modalidad')}")
            st.markdown(f"**👥 Dirigido a:** {props.get('publico')}")
            st.markdown(f"**🚗 Requiere Vehículo:** {props.get('vehiculo')}")
            st.markdown(f"**🟢 Estado:** {props.get('estado')}")
        st.markdown(f"**📝 Observaciones:** {props.get('observaciones')}")
        
        # --- BOTÓN DE ELIMINACIÓN CONDICIONAL ---
        if es_admin and id_evento is not None:
            st.markdown(" ")
            if st.button("🗑️ ELIMINAR ESTE EVENTO PERMANENTEMENTE", type="primary"):
                if eliminar_evento_por_indice(int(id_evento)):
                    st.success("¡Evento eliminado de la base de datos con éxito!")
                    st.session_state.eventos_lista = cargar_datos_desde_excel()
                    st.rerun()

# ==========================================
# SECCIÓN INFERIOR: BUSCADOR INTELIGENTE Y BASE DE DATOS
# ==========================================
if os.path.exists(ARCHIVO_DB):
    st.markdown("---")
    st.subheader("🔍 Buscador y Consolidado General de Eventos")
    
    df_visualizacion = pd.read_excel(ARCHIVO_DB)
    df_visualizacion = df_visualizacion.fillna("")
    
    # Lista de tipos para el nuevo filtro desplegable
    LISTA_TIPOS_FILTRO = ["Todos"] + LISTA_TIPOS_EVENTO
    
    # Componentes visuales del buscador modificado
    col_b1, col_b2 = st.columns([2, 1])
    with col_b1:
        texto_busqueda = st.text_input("⌨️ Buscar por texto (Responsable, Lugar, Municipio, Público...):", placeholder="Ej. Elvira, Sala Situacional, Sincelejo...")
    with col_b2:
        filtro_tipo = st.selectbox("🏷️ Filtrar por Tipo de Evento:", LISTA_TIPOS_FILTRO)

    # Aplicar filtros dinámicos en tiempo real
    df_filtrado = df_visualizacion.copy()
    
    if filtro_tipo != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Tipo de Evento"] == filtro_tipo]
        
    if texto_busqueda:
        # Busca la palabra en todas las columnas de texto simultáneamente
        mascara = df_filtrado.astype(str).apply(lambda x: x.str.contains(texto_busqueda, case=False)).any(axis=1)
        df_filtrado = df_filtrado[mascara]
        
    # Mostrar el resultado filtrado
    st.markdown(f"*Se encontraron {len(df_filtrado)} eventos registrados.*")
    st.dataframe(df_filtrado, use_container_width=True)
    
    # Exportación inteligente
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_filtrado.to_excel(writer, index=False, sheet_name='Cronograma_Filtrado')
    buffer.seek(0)
    
    st.download_button(
        label="📥 Descargar Vista Actual en EXCEL (.xlsx)",
        data=buffer,
        file_name="cronograma_reporte.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )