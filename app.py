import streamlit as st
import pandas as pd
from datetime import datetime

# Configuración inicial de la página
st.set_page_config(page_title="Control de Kilometraje", layout="wide")

st.title("🚛 Control de Kilometraje y Tiempos de Ruta")

# --- ENCABEZADO CORPORATIVO (LOGO Y EMPRESA) ---
col_logo, col_info = st.columns([1, 4])

with col_logo:
    try:
        st.image(r"C:\Users\CarlosArenas\Desktop\fletes nuevo.jpg", width=120)
    except:
        st.markdown("### 🚛 [ LOGO ]")

with col_info:
    st.markdown("### Triturados El Roble & Fletes AVE")
    st.caption("📧 carlos.arenas@triturados.com | 📞 Tel: 8119886131")
    st.markdown("*Porque tu satisfacción es nuestro destino* | *Tu carga en regla, tu destino asegurado*")

st.markdown("---")

# --- PIN DE ADMINISTRADOR ---
PIN_ADMIN = "1234"

# --- PIN DE ADMINISTRADOR ---
PIN_ADMIN = "1234"

# 1. Control de Pestaña Activa mediante session_state
if "pestaña_activa" not in st.session_state:
    st.session_state.pestaña_activa = "🟢 Etapa 1: Salida Patio Base"

# 2. Inicializar listas y base de datos de clientes
if "lista_unidades" not in st.session_state:
    st.session_state.lista_unidades = ["C-24", "C-25", "C-26", "T-01", "T-02"]

if "lista_operadores" not in st.session_state:
    st.session_state.lista_operadores = [
        "Octavio Rodrigo Serrano Saavedra",
        "Juan Carlos Arenas",
        "Pedro Martinez",
        "Jose Luis Rodriguez"
    ]

if "lista_clientes" not in st.session_state:
    st.session_state.lista_clientes = [
        "16 - Néstor Guerrero Segura",
        "Triturados El Roble",
        "Cliente Genérico"
    ]

if "db_clientes" not in st.session_state:
    st.session_state.db_clientes = {
        "16 - Néstor Guerrero Segura": {
            "contacto": "Néstor Guerrero",
            "telefono": "8126807881",
            "latitud": 25.844412,
            "longitud": -100.395043
        },
        "Triturados El Roble": {
            "contacto": "Carlos Arenas",
            "telefono": "8119886131",
            "latitud": 25.795000,
            "longitud": -100.300000
        },
        "Cliente Genérico": {
            "contacto": "Atención a Clientes",
            "telefono": "8000000000",
            "latitud": 25.686614,
            "longitud": -100.316113
        }
    }

# Variables de registro de odómetro y tiempos
if "km_salida_patio" not in st.session_state:
    st.session_state.km_salida_patio = 145826.0
if "hora_salida_patio" not in st.session_state:
    st.session_state.hora_salida_patio = ""

if "km_llegada_pedrera" not in st.session_state:
    st.session_state.km_llegada_pedrera = 145826.0
if "hora_llegada_pedrera" not in st.session_state:
    st.session_state.hora_llegada_pedrera = ""

if "km_salida_pedrera" not in st.session_state:
    st.session_state.km_salida_pedrera = 145826.0
if "hora_salida_pedrera" not in st.session_state:
    st.session_state.hora_salida_pedrera = ""

if "km_llegada_cliente" not in st.session_state:
    st.session_state.km_llegada_cliente = 145826.0
if "hora_llegada_cliente" not in st.session_state:
    st.session_state.hora_llegada_cliente = ""

if "km_salida_cliente" not in st.session_state:
    st.session_state.km_salida_cliente = 145826.0
if "hora_salida_cliente" not in st.session_state:
    st.session_state.hora_salida_cliente = ""

if "km_final_viaje" not in st.session_state:
    st.session_state.km_final_viaje = 145826.0
if "hora_cierre_viaje" not in st.session_state:
    st.session_state.hora_cierre_viaje = ""

if "cli_contacto" not in st.session_state:
    st.session_state.cli_contacto = ""
if "cli_telefono" not in st.session_state:
    st.session_state.cli_telefono = ""
if "cli_latitud" not in st.session_state:
    st.session_state.cli_latitud = 25.844412
if "cli_longitud" not in st.session_state:
    st.session_state.cli_longitud = -100.395043

def actualizar_campos_cliente():
    c_sel = st.session_state.get("sel_cliente_e2")
    if c_sel in st.session_state.db_clientes:
        info = st.session_state.db_clientes[c_sel]
        st.session_state.cli_contacto = info.get("contacto", "")
        st.session_state.cli_telefono = str(info.get("telefono", ""))
        st.session_state.cli_latitud = float(info.get("latitud", 25.844412))
        val_lon = float(info.get("longitud", -100.395043))
        st.session_state.cli_longitud = abs(val_lon) * -1.0

# --- BARRA LATERAL: CONTROL DE ACCESO ADMIN ---
st.sidebar.title("🔒 Acceso Administrador")
pin_input = st.sidebar.text_input("Ingresa PIN de Administrador:", type="password")

es_admin = (pin_input == PIN_ADMIN)

if es_admin:
    st.sidebar.success("🔑 Modo Administrador Activo")
else:
    if pin_input != "":
        st.sidebar.error("❌ PIN Incorrecto")
    else:
        st.sidebar.info("Modo Usuario (Sólo lectura de catálogos)")

# --- NAVEGACIÓN CONTROLADA ---
opciones_pestañas = [
    "🟢 Etapa 1: Salida Patio Base",
    "🟡 Etapa Intermedia: Pedrera (Llegada/Salida)",
    "🟠 Etapa 2: Llegada/Salida Cliente",
    "🔴 Etapa 3: Cierre de Ruta / Regreso"
]

st.radio(
    "Selecciona la Etapa:",
    opciones_pestañas,
    key="pestaña_activa",
    horizontal=True
)

st.markdown("---")

# ==========================================
# --- ETAPA 1: SALIDA DE PATIO BASE ---
# ==========================================
if st.session_state.pestaña_activa == "🟢 Etapa 1: Salida Patio Base":
    st.header("Etapa 1: Salida de Patio Base")

    if es_admin:
        with st.expander("⚙️ Administración de Catálogos (Alta, Carga Excel y Eliminación)"):
            col_m1, col_m2, col_m3 = st.columns(3)
            
            with col_m1:
                st.subheader("1. Alta Manual")
                nuevo_tipo = st.selectbox("¿Qué deseas agregar?", ["Unidad", "Operador", "Cliente"], key="alta_tipo")
                nuevo_nombre = st.text_input("Escribe el nombre o código:", key="alta_nombre")
                nuevo_contacto = ""
                nuevo_tel = ""
                if nuevo_tipo == "Cliente":
                    nuevo_contacto = st.text_input("Contacto del Cliente (Opcional):", key="alta_contacto")
                    nuevo_tel = st.text_input("Teléfono (Opcional):", key="alta_tel")
                
                if st.button("➕ Agregar al Catálogo"):
                    if nuevo_nombre.strip():
                        val = nuevo_nombre.strip()
                        if nuevo_tipo == "Unidad" and val not in st.session_state.lista_unidades:
                            st.session_state.lista_unidades.append(val)
                            st.success(f"Unidad '{val}' agregada.")
                        elif nuevo_tipo == "Operador" and val not in st.session_state.lista_operadores:
                            st.session_state.lista_operadores.append(val)
                            st.success(f"Operador '{val}' agregado.")
                        elif nuevo_tipo == "Cliente" and val not in st.session_state.lista_clientes:
                            st.session_state.lista_clientes.append(val)
                            st.session_state.db_clientes[val] = {
                                "contacto": nuevo_contacto.strip(),
                                "telefono": nuevo_tel.strip(),
                                "latitud": 25.844412,
                                "longitud": -100.395043
                            }
                            st.success(f"Cliente '{val}' agregado.")
                        st.rerun()
                    else:
                        st.warning("Escribe un valor válido.")

            with col_m2:
                st.subheader("2. Carga Masiva (Excel)")
                archivo_excel = st.file_uploader("Sube tu archivo .xlsx o .xls", type=["xlsx", "xls"])
                
                if archivo_excel is not None:
                    try:
                        df = pd.read_excel(archivo_excel)
                        st.dataframe(df.head(3))
                        
                        if st.button("📥 Importar desde Excel"):
                            cols_map = {str(col).strip().lower(): col for col in df.columns}
                            
                            col_u = next((c for norm, c in cols_map.items() if "unidad" in norm), None)
                            if col_u:
                                vals = df[col_u].dropna().astype(str).str.strip().tolist()
                                vals = [v for v in vals if v and v.lower() != "none"]
                                st.session_state.lista_unidades = list(set(st.session_state.lista_unidades + vals))

                            col_o = next((c for norm, c in cols_map.items() if "operador" in norm), None)
                            if col_o:
                                vals = df[col_o].dropna().astype(str).str.strip().tolist()
                                vals = [v for v in vals if v and v.lower() != "none"]
                                st.session_state.lista_operadores = list(set(st.session_state.lista_operadores + vals))

                            col_num = next((c for norm, c in cols_map.items() if "num_cliente" in norm or "num" in norm), None)
                            col_nom = next((c for norm, c in cols_map.items() if "nombre" in norm or "cliente" in norm), None)
                            col_cont = next((c for norm, c in cols_map.items() if "contacto" in norm), None)
                            col_tel = next((c for norm, c in cols_map.items() if "tel" in norm or "telefono" in norm), None)
                            col_lat = next((c for norm, c in cols_map.items() if "lat" in norm or "latitud" in norm), None)
                            col_lon = next((c for norm, c in cols_map.items() if "lon" in norm or "longitud" in norm), None)
                            
                            if col_nom or col_num:
                                for idx, row in df.iterrows():
                                    num_val = str(row[col_num]).split('.')[0].strip() if col_num and pd.notna(row[col_num]) else ""
                                    nom_val = str(row[col_nom]).strip() if col_nom and pd.notna(row[col_nom]) else ""
                                    
                                    if num_val and nom_val:
                                        cliente_key = f"{num_val} - {nom_val}"
                                    elif num_val:
                                        cliente_key = num_val
                                    else:
                                        cliente_key = nom_val
                                        
                                    if cliente_key and cliente_key.lower() != "none":
                                        if cliente_key not in st.session_state.lista_clientes:
                                            st.session_state.lista_clientes.append(cliente_key)
                                        
                                        if cliente_key not in st.session_state.db_clientes:
                                            st.session_state.db_clientes[cliente_key] = {
                                                "contacto": "",
                                                "telefono": "",
                                                "latitud": 25.844412,
                                                "longitud": -100.395043
                                            }
                                        
                                        if col_cont and pd.notna(row[col_cont]):
                                            st.session_state.db_clientes[cliente_key]["contacto"] = str(row[col_cont]).strip()
                                        if col_tel and pd.notna(row[col_tel]):
                                            tel_raw = str(row[col_tel]).split('.')[0].strip()
                                            st.session_state.db_clientes[cliente_key]["telefono"] = tel_raw
                                            
                                        if col_lat and pd.notna(row[col_lat]):
                                            try:
                                                st.session_state.db_clientes[cliente_key]["latitud"] = float(row[col_lat])
                                            except: pass
                                            
                                        if col_lon and pd.notna(row[col_lon]):
                                            try:
                                                val_lon = abs(float(row[col_lon])) * -1.0
                                                st.session_state.db_clientes[cliente_key]["longitud"] = val_lon
                                            except: pass

                            st.success("✅ Clientes e información importados correctamente.")
                            st.rerun()
                    except Exception as e:
                        st.error(f"Error al procesar el archivo: {e}")

            with col_m3:
                st.subheader("3. Eliminar Registros")
                tipo_eliminar = st.selectbox("¿Qué deseas eliminar?", ["Unidad", "Operador", "Cliente"], key="del_tipo")
                
                if tipo_eliminar == "Unidad":
                    item_eliminar = st.selectbox("Selecciona Unidad a eliminar:", st.session_state.lista_unidades, key="del_u")
                elif tipo_eliminar == "Operador":
                    item_eliminar = st.selectbox("Selecciona Operador a eliminar:", st.session_state.lista_operadores, key="del_o")
                else:
                    item_eliminar = st.selectbox("Selecciona Cliente a eliminar:", st.session_state.lista_clientes, key="del_c")
                
                if st.button("🗑️ Eliminar del Catálogo"):
                    if item_eliminar:
                        if tipo_eliminar == "Unidad":
                            st.session_state.lista_unidades.remove(item_eliminar)
                        elif tipo_eliminar == "Operador":
                            st.session_state.lista_operadores.remove(item_eliminar)
                        elif tipo_eliminar == "Cliente":
                            st.session_state.lista_clientes.remove(item_eliminar)
                            if item_eliminar in st.session_state.db_clientes:
                                del st.session_state.db_clientes[item_eliminar]
                            
                        st.success(f"❌ {tipo_eliminar} '{item_eliminar}' eliminado.")
                        st.rerun()

        st.markdown("---")

    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        unidad_seleccionada = st.selectbox("Selecciona la Unidad:", st.session_state.lista_unidades, key="sel_unidad")
    with col_sel2:
        operador_seleccionado = st.selectbox("Selecciona el Operador:", st.session_state.lista_operadores, key="sel_operador")
        
    st.session_state.unidad_activa = unidad_seleccionada
    st.session_state.operador_activo = operador_seleccionado
    
    st.markdown(
        f"**Unidad seleccionada:** `{unidad_seleccionada}` | "
        f"**Operador:** `{operador_seleccionado}` | "
        f"**Kilómetros de patio:** `{st.session_state.km_salida_patio}`"
    )
    st.markdown("---")
    
    col_e1_1, col_e1_2 = st.columns(2)
    with col_e1_1:
        st.subheader("Captura de Salida de Patio")
        km_patio_input = st.number_input("KM Salida de Patio Base:", value=float(st.session_state.km_salida_patio), step=1.0, key="km_patio_in")
        
        if st.button("Registrar Salida de Patio Base", key="btn_e1"):
            st.session_state.km_salida_patio = km_patio_input
            st.session_state.hora_salida_patio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.session_state.pestaña_activa = "🟡 Etapa Intermedia: Pedrera (Llegada/Salida)"
            st.rerun()

    with col_e1_2:
        st.subheader("Confirmación de Salida")
        if st.session_state.hora_salida_patio:
            st.info(f"📍 **KM Salida Patio:** {st.session_state.km_salida_patio} KM")
            st.info(f"⏱️ **Fecha y Hora de Salida:** {st.session_state.hora_salida_patio}")
        else:
            st.caption("Aún no se ha registrado la fecha/hora de salida.")

# ==========================================
# --- ETAPA INTERMEDIA: PEDRERA ---
# ==========================================
elif st.session_state.pestaña_activa == "🟡 Etapa Intermedia: Pedrera (Llegada/Salida)":
    st.header("Etapa Intermedia: Registro en Pedrera")
    
    unidad = st.session_state.get("unidad_activa", "C-24")
    operador = st.session_state.get("operador_activo", "Octavio Rodrigo Serrano Saavedra")
    st.markdown(f"**Unidad Activa:** `{unidad}` | **Operador:** `{operador}`")
    st.markdown("---")
    
    col_p1, col_p2 = st.columns(2)
    
    with col_p1:
        st.subheader("1. Llegada a Pedrera")
        km_llegada_p_in = st.number_input("KM Llegada a Pedrera:", value=float(st.session_state.km_llegada_pedrera), step=1.0, key="km_llegada_p_in")
        if st.button("Registrar Llegada a Pedrera", key="btn_llegada_p"):
            st.session_state.km_llegada_pedrera = km_llegada_p_in
            st.session_state.hora_llegada_pedrera = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.success("✅ Llegada a Pedrera registrada.")
            
        if st.session_state.hora_llegada_pedrera:
            st.caption(f"⏱️ **Llegada Registrada:** {st.session_state.hora_llegada_pedrera}")

    with col_p2:
        st.subheader("2. Salida de Pedrera")
        km_salida_p_in = st.number_input("KM Salida de Pedrera:", value=float(st.session_state.km_salida_pedrera), step=1.0, key="km_salida_p_in")
        if st.button("Registrar Salida de Pedrera", key="btn_salida_p"):
            st.session_state.km_salida_pedrera = km_salida_p_in
            st.session_state.hora_salida_pedrera = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.session_state.pestaña_activa = "🟠 Etapa 2: Llegada/Salida Cliente"
            st.rerun()
            
        if st.session_state.hora_salida_pedrera:
            st.caption(f"⏱️ **Salida Registrada:** {st.session_state.hora_salida_pedrera}")

# ==========================================
# --- ETAPA 2: LLEGADA Y SALIDA CON CLIENTE ---
# ==========================================
elif st.session_state.pestaña_activa == "🟠 Etapa 2: Llegada/Salida Cliente":
    st.header("Etapa 2: Registro de Llegada y Salida con el Cliente")
    
    unidad = st.session_state.get("unidad_activa", "C-24")
    operador = st.session_state.get("operador_activo", "Octavio Rodrigo Serrano Saavedra")
    st.markdown(f"**Unidad Activa:** `{unidad}` | **Operador:** `{operador}`")
    st.markdown("---")
    
    col_c1, col_c2 = st.columns(2)
    
    with col_c1:
        st.subheader("Cliente Asignado:")
        
        if not st.session_state.cli_contacto and st.session_state.lista_clientes:
            actualizar_campos_cliente()
            
        cliente_seleccionado = st.selectbox(
            "Escribe o selecciona el cliente:", 
            st.session_state.lista_clientes, 
            key="sel_cliente_e2",
            on_change=actualizar_campos_cliente
        )
        
        st.subheader("Datos de Entrega / Ubicación:")
        
        contacto = st.text_input("Contacto:", value=st.session_state.cli_contacto, key="txt_contacto_fixed")
        telefono = st.text_input("Teléfono:", value=st.session_state.cli_telefono, key="txt_tel_fixed")
        
        col_geo1, col_geo2 = st.columns(2)
        with col_geo1:
            latitud = st.number_input("Latitud:", value=float(st.session_state.cli_latitud), format="%.6f", key="geo_lat_fixed")
        with col_geo2:
            longitud = st.number_input("Longitud:", value=float(st.session_state.cli_longitud), format="%.6f", key="geo_lon_fixed")
            
        url_maps = f"https://www.google.com/maps?q={latitud:.6f},{longitud:.6f}"
        st.markdown(f"📍 **Ubicación en Google Maps:** [Abrir mapa en Google Maps]({url_maps})")

    with col_c2:
        if st.session_state.hora_salida_pedrera:
            st.info(f"⏱️ **Hora Salida Pedrera:** {st.session_state.hora_salida_pedrera}")
        else:
            st.warning("⚠️ Pendiente de registrar salida en Pedrera")
            
        st.markdown("---")
        
        st.subheader("1. Registro de Llegada al Cliente")
        km_llegada_cli_in = st.number_input(
            "KM Llegada con Cliente:", 
            value=float(st.session_state.km_salida_pedrera), 
            step=1.0,
            key="km_llegada_cli_input"
        )
        
        if st.button("Registrar Llegada con Cliente", key="btn_llegada_cliente"):
            st.session_state.km_llegada_cliente = km_llegada_cli_in
            st.session_state.hora_llegada_cliente = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if st.session_state.hora_llegada_cliente:
            st.success(f"✅ **Entrada registrada a las** {st.session_state.hora_llegada_cliente} | **KM:** {st.session_state.km_llegada_cliente}")
            st.caption(f"👤 **Contacto:** {contacto} | 📞 **Tel:** {telefono}")
        
        st.markdown("---")
        
        st.subheader("2. Registro de Salida con el Cliente")
        km_salida_cli_in = st.number_input(
            "KM Salida con Cliente:", 
            value=float(st.session_state.km_llegada_cliente), 
            step=1.0,
            key="km_salida_cli_input"
        )
        
        if st.button("Registrar Salida con Cliente", key="btn_salida_cliente"):
            st.session_state.km_salida_cliente = km_salida_cli_in
            st.session_state.hora_salida_cliente = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.session_state.pestaña_activa = "🔴 Etapa 3: Cierre de Ruta / Regreso"
            st.rerun()

        if st.session_state.hora_salida_cliente:
            st.info(f"🚀 **Salida registrada a las** {st.session_state.hora_salida_cliente} | **KM:** {st.session_state.km_salida_cliente}")

# ==========================================
# --- ETAPA 3: CIERRE DE RUTA / REGRESO ---
# ==========================================
elif st.session_state.pestaña_activa == "🔴 Etapa 3: Cierre de Ruta / Regreso":
    st.header("Etapa 3: Cierre de Ruta / Regreso a Patio Base")
    
    unidad = st.session_state.get("unidad_activa", "C-24")
    operador = st.session_state.get("operador_activo", "Octavio Rodrigo Serrano Saavedra")
    st.markdown(f"**Unidad Activa:** `{unidad}` | **Operador:** `{operador}`")
    st.markdown("---")
    
    col_e3_1, col_e3_2 = st.columns(2)
    
    with col_e3_1:
        st.subheader("Captura de Cierre y Regreso")
        
        st.text_input(
            "KM Salida con Cliente (Origen anterior):",
            value=f"{st.session_state.km_salida_cliente:.1f} KM",
            disabled=True,
            key="km_salida_cli_e3_disabled"
        )
        
        km_final_in = st.number_input(
            "KM Final del Viaje (Llegada Patio Base):",
            value=float(st.session_state.km_salida_cliente),
            step=1.0,
            key="km_final_viaje_input"
        )
        
        if st.button("🏁 Registrar Cierre de Ruta", key="btn_cierre_ruta"):
            if km_final_in < st.session_state.km_salida_cliente:
                st.error("⚠️ El KM Final no puede ser menor al KM de Salida del Cliente.")
            else:
                st.session_state.km_final_viaje = km_final_in
                st.session_state.hora_cierre_viaje = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                st.success("✅ Cierre de ruta registrado exitosamente.")

    with col_e3_2:
        st.subheader("Resumen de Cierre de Ruta")
        
        if st.session_state.hora_cierre_viaje:
            distancia_total = st.session_state.km_final_viaje - st.session_state.km_salida_patio
            
            st.success(f"🏁 **Cierre de Ruta Registrado:** {st.session_state.hora_cierre_viaje}")
            st.info(f"📍 **KM Final:** {st.session_state.km_final_viaje} KM")
            st.metric("📏 Distancia Total Recorrida", f"{distancia_total:.1f} KM")
            
            if st.session_state.hora_salida_cliente:
                st.caption(f"⏱️ **Salida del cliente:** {st.session_state.hora_salida_cliente}")
        else:
            st.caption("Aún no se ha registrado el cierre final de la ruta.")