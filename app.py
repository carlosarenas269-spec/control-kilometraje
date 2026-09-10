import os
import streamlit as st
import pandas as pd
from datetime import datetime

# ==========================================
# CONFIGURACIÓN Y PERSISTENCIA DE DATOS
# ==========================================
ARCHIVO_REGISTROS = "registros_rutas.xlsx"
ARCHIVO_CATALOGOS = "catalogos_guardados.xlsx"

def guardar_en_excel(data_dict, archivo_excel=ARCHIVO_REGISTROS):
    """Guarda los datos de la ruta en un archivo Excel de forma segura."""
    try:
        if os.path.exists(archivo_excel):
            df_existente = pd.read_excel(archivo_excel)
            df_nuevo = pd.DataFrame([data_dict])
            df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
        else:
            df_final = pd.DataFrame([data_dict])

        df_final.to_excel(archivo_excel, index=False)
    except Exception as e:
        st.error(f"Error al guardar el registro en Excel: {e}")

def guardar_catalogos_en_disco():
    """Guarda las listas y la base de datos de clientes actualizadas en un Excel permanente."""
    try:
        with pd.ExcelWriter(ARCHIVO_CATALOGOS, engine='openpyxl') as writer:
            pd.DataFrame({"Unidades": st.session_state.lista_unidades}).to_excel(writer, sheet_name="Unidades", index=False)
            pd.DataFrame({"Operadores": st.session_state.lista_operadores}).to_excel(writer, sheet_name="Operadores", index=False)
            
            clientes_data = []
            for cli, info in st.session_state.db_clientes.items():
                clientes_data.append({
                    "Cliente": cli,
                    "Contacto": info.get("contacto", ""),
                    "Telefono": info.get("telefono", ""),
                    "Latitud": info.get("latitud", 25.844412),
                    "Longitud": info.get("longitud", -100.395043)
                })
            pd.DataFrame(clientes_data).to_excel(writer, sheet_name="Clientes", index=False)
    except Exception as e:
        print(f"Error al guardar catálogos en disco: {e}")

def cargar_catalogos_de_disco():
    """Carga los catálogos desde el disco si el archivo existe."""
    if os.path.exists(ARCHIVO_CATALOGOS):
        try:
            xl = pd.ExcelFile(ARCHIVO_CATALOGOS)
            if "Unidades" in xl.sheet_names:
                df_u = pd.read_excel(ARCHIVO_CATALOGOS, sheet_name="Unidades")
                st.session_state.lista_unidades = df_u["Unidades"].dropna().astype(str).tolist()
            
            if "Operadores" in xl.sheet_names:
                df_o = pd.read_excel(ARCHIVO_CATALOGOS, sheet_name="Operadores")
                st.session_state.lista_operadores = df_o["Operadores"].dropna().astype(str).tolist()
                
            if "Clientes" in xl.sheet_names:
                df_c = pd.read_excel(ARCHIVO_CATALOGOS, sheet_name="Clientes")
                nueva_db = {}
                lista_c = []
                for _, row in df_c.iterrows():
                    cli_nombre = str(row["Cliente"])
                    lista_c.append(cli_nombre)
                    nueva_db[cli_nombre] = {
                        "contacto": str(row["Contacto"]) if pd.notna(row["Contacto"]) else "",
                        "telefono": str(row["Telefono"]) if pd.notna(row["Telefono"]) else "",
                        "latitud": float(row["Latitud"]) if pd.notna(row["Latitud"]) else 25.844412,
                        "longitud": float(row["Longitud"]) if pd.notna(row["Longitud"]) else -100.395043
                    }
                if lista_c:
                    st.session_state.lista_clientes = lista_c
                    st.session_state.db_clientes = nueva_db
        except Exception as e:
            print(f"Error al leer catálogos de disco: {e}")

# ==========================================
# CONFIGURACIÓN INICIAL DE LA APP
# ==========================================
st.set_page_config(page_title="Fletes AVE Control", layout="wide")
st.title("🚛 Fletes AVE Control")

PIN_ADMIN = "1234"

opciones_pestañas = [
    "🟢 Etapa 1: Salida Patio Base",
    "🟡 Etapa Intermedia: Pedrera (Llegada/Salida)",
    "🟠 Etapa 2: Llegada/Salida Cliente",
    "🔴 Etapa 3: Cierre de Ruta / Regreso"
]

if "etapa_idx" not in st.session_state:
    st.session_state.etapa_idx = 0

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

if "catalogos_cargados" not in st.session_state:
    cargar_catalogos_de_disco()
    st.session_state.catalogos_cargados = True

# Variables de odómetro y tiempos
for var, val in [
    ("km_salida_patio", 145826.0), ("hora_salida_patio", ""),
    ("km_llegada_pedrera", 145826.0), ("hora_llegada_pedrera", ""),
    ("km_salida_pedrera", 145826.0), ("hora_salida_pedrera", ""),
    ("km_llegada_cliente", 145826.0), ("hora_llegada_cliente", ""),
    ("km_salida_cliente", 145826.0), ("hora_salida_cliente", ""),
    ("km_final_viaje", 145826.0), ("hora_cierre_viaje", ""),
    ("cli_contacto", ""), ("cli_telefono", ""),
    ("cli_latitud", 25.844412), ("cli_longitud", -100.395043)
]:
    if var not in st.session_state:
        st.session_state[var] = val

def actualizar_campos_cliente():
    c_sel = st.session_state.get("sel_cliente_e2")
    if c_sel in st.session_state.db_clientes:
        info = st.session_state.db_clientes[c_sel]
        st.session_state.cli_contacto = info.get("contacto", "")
        st.session_state.cli_telefono = str(info.get("telefono", ""))
        st.session_state.cli_latitud = float(info.get("latitud", 25.844412))
        val_lon = float(info.get("longitud", -100.395043))
        st.session_state.cli_longitud = abs(val_lon) * -1.0

# ==========================================
# BARRA LATERAL: LOGO Y CONTROL DE ACCESO ADMIN
# ==========================================
try:
    st.sidebar.image(r"C:\Users\CarlosArenas\Desktop\logo_fletes_ave.jpg", use_container_width=True)
except Exception:
    try:
        st.sidebar.image("logo_fletes_ave.jpg", use_container_width=True)
    except Exception:
        st.sidebar.warning("⚠️ No se pudo cargar 'logo_fletes_ave.jpg'. Verifica la ruta.")

st.sidebar.title("🔒 Acceso Administrador")
pin_input = st.sidebar.text_input("Ingresa PIN de Administrador:", type="password")
es_admin = (pin_input == PIN_ADMIN)

if es_admin:
    st.sidebar.success("🔑 Modo Administrador Activo")
else:
    if pin_input != "":
        st.sidebar.error("❌ PIN Incorrecto")
    else:
        st.sidebar.info("Modo Usuario (Sólo lectura)")

# ==========================================
# NAVEGACIÓN POR PESTAÑAS (USANDO ÍNDICE)
# ==========================================
pestana_activa = st.radio(
    "Selecciona la Etapa:", 
    opciones_pestañas, 
    index=st.session_state.etapa_idx,
    horizontal=True
)

st.session_state.etapa_idx = opciones_pestañas.index(pestana_activa)

st.markdown("---")

# ==========================================
# ETAPA 1: SALIDA DE PATIO BASE
# ==========================================
if st.session_state.etapa_idx == 0:
    st.header("Etapa 1: Salida de Patio Base")

    if es_admin:
        with st.expander("⚙️ Administración de Catálogos"):
            col_m1, col_m2, col_m3 = st.columns(3)
            
            with col_m1:
                st.subheader("1. Alta Manual")
                nuevo_tipo = st.selectbox("¿Qué deseas agregar?", ["Unidad", "Operador", "Cliente"], key="alta_tipo")
                nuevo_nombre = st.text_input("Escribe el nombre o código:", key="alta_nombre")
                nuevo_contacto = ""
                nuevo_tel = ""
                if nuevo_tipo == "Cliente":
                    nuevo_contacto = st.text_input("Contacto (Opcional):", key="alta_contacto")
                    nuevo_tel = st.text_input("Teléfono (Opcional):", key="alta_tel")
                
                if st.button("➕ Agregar al Catálogo"):
                    if nuevo_nombre.strip():
                        val = nuevo_nombre.strip()
                        if nuevo_tipo == "Unidad" and val not in st.session_state.lista_unidades:
                            st.session_state.lista_unidades.append(val)
                        elif nuevo_tipo == "Operador" and val not in st.session_state.lista_operadores:
                            st.session_state.lista_operadores.append(val)
                        elif nuevo_tipo == "Cliente" and val not in st.session_state.lista_clientes:
                            st.session_state.lista_clientes.append(val)
                            st.session_state.db_clientes[val] = {
                                "contacto": nuevo_contacto.strip(),
                                "telefono": nuevo_tel.strip(),
                                "latitud": 25.844412,
                                "longitud": -100.395043
                            }
                        guardar_catalogos_en_disco()
                        st.success(f"'{val}' agregado con éxito.")
                        st.rerun()
                    else:
                        st.warning("Escribe un valor válido.")

            with col_m2:
                st.subheader("2. Carga Masiva (Excel)")
                archivo_excel = st.file_uploader("Sube tu archivo .xlsx", type=["xlsx", "xls"])
                if archivo_excel is not None:
                    try:
                        df = pd.read_excel(archivo_excel)
                        if st.button("📥 Importar desde Excel"):
                            cols_map = {str(col).strip().lower(): col for col in df.columns}
                            col_u = next((c for norm, c in cols_map.items() if "unidad" in norm), None)
                            if col_u:
                                vals = df[col_u].dropna().astype(str).str.strip().tolist()
                                st.session_state.lista_unidades = list(set(st.session_state.lista_unidades + vals))

                            col_o = next((c for norm, c in cols_map.items() if "operador" in norm), None)
                            if col_o:
                                vals = df[col_o].dropna().astype(str).str.strip().tolist()
                                st.session_state.lista_operadores = list(set(st.session_state.lista_operadores + vals))

                            guardar_catalogos_en_disco()
                            st.success("✅ Importado correctamente.")
                            st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

            with col_m3:
                st.subheader("3. Eliminar Registros")
                tipo_eliminar = st.selectbox("¿Qué deseas eliminar?", ["Unidad", "Operador", "Cliente"], key="del_tipo")
                item_eliminar = st.selectbox("Selecciona:", st.session_state.lista_unidades if tipo_eliminar=="Unidad" else (st.session_state.lista_operadores if tipo_eliminar=="Operador" else st.session_state.lista_clientes), key="del_item")
                
                if st.button("🗑️ Eliminar"):
                    if item_eliminar:
                        if tipo_eliminar == "Unidad": st.session_state.lista_unidades.remove(item_eliminar)
                        elif tipo_eliminar == "Operador": st.session_state.lista_operadores.remove(item_eliminar)
                        elif tipo_eliminar == "Cliente":
                            st.session_state.lista_clientes.remove(item_eliminar)
                            if item_eliminar in st.session_state.db_clientes: del st.session_state.db_clientes[item_eliminar]
                        guardar_catalogos_en_disco()
                        st.success("Eliminado correctamente.")
                        st.rerun()

        st.markdown("---")

    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        unidad_seleccionada = st.selectbox("Selecciona la Unidad:", st.session_state.lista_unidades, key="sel_unidad")
    with col_sel2:
        operador_seleccionado = st.selectbox("Selecciona el Operador:", st.session_state.lista_operadores, key="sel_operador")
        
    st.session_state.unidad_activa = unidad_seleccionada
    st.session_state.operador_activo = operador_seleccionado
    
    col_e1_1, col_e1_2 = st.columns(2)
    with col_e1_1:
        st.subheader("Captura de Salida de Patio")
        km_patio_input = st.number_input("KM Salida de Patio Base:", value=float(st.session_state.km_salida_patio), step=1.0, key="km_patio_in")
        
        if st.button("Registrar Salida de Patio Base", key="btn_e1"):
            st.session_state.km_salida_patio = km_patio_input
            st.session_state.hora_salida_patio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.session_state.km_llegada_pedrera = km_patio_input
            st.session_state.etapa_idx = 1
            st.rerun()

    with col_e1_2:
        st.subheader("Confirmación")
        if st.session_state.hora_salida_patio:
            st.info(f"📍 **KM:** {st.session_state.km_salida_patio} | **Hora:** {st.session_state.hora_salida_patio}")

# ==========================================
# ETAPA INTERMEDIA: PEDRERA
# ==========================================
elif st.session_state.etapa_idx == 1:
    st.header("Etapa Intermedia: Registro en Pedrera")
    st.markdown(f"**Unidad:** `{st.session_state.get('unidad_activa', 'C-24')}` | **Operador:** `{st.session_state.get('operador_activo', '')}`")
    st.markdown("---")

    col_p1, col_p2 = st.columns(2)
    with col_p1:
        st.subheader("1. Llegada a Pedrera")
        km_llegada_p_in = st.number_input("KM Llegada Pedrera:", value=float(st.session_state.km_llegada_pedrera), step=1.0, key="km_llegada_p_in")
        if st.button("Registrar Llegada a Pedrera", key="btn_llegada_p"):
            st.session_state.km_llegada_pedrera = km_llegada_p_in
            st.session_state.hora_llegada_pedrera = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.session_state.km_salida_pedrera = km_llegada_p_in
            st.success("✅ Llegada registrada.")
            st.rerun()
        if st.session_state.hora_llegada_pedrera:
            st.caption(f"⏱️ {st.session_state.hora_llegada_pedrera}")

    with col_p2:
        st.subheader("2. Salida de Pedrera")
        km_salida_p_in = st.number_input("KM Salida Pedrera:", value=float(st.session_state.km_salida_pedrera), step=1.0, key="km_salida_p_in")
        if st.button("Registrar Salida de Pedrera", key="btn_salida_p"):
            st.session_state.km_salida_pedrera = km_salida_p_in
            st.session_state.hora_salida_pedrera = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.session_state.km_llegada_cliente = km_salida_p_in
            st.session_state.etapa_idx = 2
            st.rerun()
        if st.session_state.hora_salida_pedrera:
            st.caption(f"⏱️ {st.session_state.hora_salida_pedrera}")

# ==========================================
# ETAPA 2: LLEGADA Y SALIDA CON CLIENTE
# ==========================================
elif st.session_state.etapa_idx == 2:
    st.header("Etapa 2: Registro con el Cliente")
    st.markdown(f"**Unidad:** `{st.session_state.get('unidad_activa', 'C-24')}` | **Operador:** `{st.session_state.get('operador_activo', '')}`")
    st.markdown("---")
    
    col_c1, col_c2 = st.columns(2)
    
    with col_c1:
        if not st.session_state.cli_contacto and st.session_state.lista_clientes:
            actualizar_campos_cliente()
            
        st.selectbox("Selecciona el cliente:", st.session_state.lista_clientes, key="sel_cliente_e2", on_change=actualizar_campos_cliente)
        
        info_actual = st.session_state.db_clientes.get(st.session_state.get("sel_cliente_e2"), {})
        cliente_nombre_seleccionado = st.session_state.get("sel_cliente_e2", "")
        contacto = info_actual.get("contacto", "")
        telefono = info_actual.get("telefono", "")
        latitud = info_actual.get("latitud", 25.844412)
        longitud = info_actual.get("longitud", -100.395043)
        
        st.markdown(f"**🏢 Cliente Seleccionado:** `{cliente_nombre_seleccionado}`")
        
        st.text_input("Contacto:", value=contacto, disabled=True)
        st.text_input("Teléfono:", value=telefono, disabled=True)
        
        if telefono and telefono.isdigit():
            col_btn_tel, col_btn_wa = st.columns(2)
            with col_btn_tel:
                st.markdown(
                    f'<a href="tel:{telefono}" target="_self">'
                    f'<button style="width:100%; background-color:#2e7d32; color:white; border:none; padding:8px; border-radius:5px; font-weight:bold; cursor:pointer;">📞 Llamar</button>'
                    f'</a>',
                    unsafe_allow_html=True
                )
            with col_btn_wa:
                tel_clean = telefono.strip()
                st.markdown(
                    f'<a href="https://wa.me/{tel_clean}" target="_blank">'
                    f'<button style="width:100%; background-color:#25D366; color:white; border:none; padding:8px; border-radius:5px; font-weight:bold; cursor:pointer;">💬 WhatsApp</button>'
                    f'</a>',
                    unsafe_allow_html=True
                )
        
        col_geo1, col_geo2 = st.columns(2)
        with col_geo1: st.number_input("Latitud:", value=float(latitud), format="%.6f", disabled=True)
        with col_geo2: st.number_input("Longitud:", value=float(longitud), format="%.6f", disabled=True)
            
        st.markdown(f"📍 [Abrir en Google Maps](https://www.google.com/maps?q={latitud:.6f},{longitud:.6f})")

    with col_c2:
        st.subheader("1. Registro de Llegada")
        km_llegada_cli_in = st.number_input("KM Llegada con Cliente:", value=float(st.session_state.km_llegada_cliente), step=1.0, key="km_llegada_cli_input")
        
        if st.button("Registrar Llegada con Cliente", key="btn_llegada_cliente"):
            st.session_state.km_llegada_cliente = km_llegada_cli_in
            st.session_state.hora_llegada_cliente = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.session_state.km_salida_cliente = km_llegada_cli_in
            st.rerun()

        if st.session_state.hora_llegada_cliente:
            st.success(f"✅ Llegada: {st.session_state.hora_llegada_cliente} | KM: {st.session_state.km_llegada_cliente}")
        
        st.markdown("---")
        
        st.subheader("2. Registro de Salida")
        km_salida_cli_in = st.number_input("KM Salida con Cliente:", value=float(st.session_state.km_salida_cliente), step=1.0, key="km_salida_cli_input")
        
        if st.button("Registrar Salida con Cliente", key="btn_salida_cliente"):
            st.session_state.km_salida_cliente = km_salida_cli_in
            st.session_state.hora_salida_cliente = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.session_state.km_final_viaje = km_salida_cli_in
            st.session_state.etapa_idx = 3
            st.rerun()

        if st.session_state.hora_salida_cliente:
            st.info(f"🚀 Salida: {st.session_state.hora_salida_cliente} | KM: {st.session_state.km_salida_cliente}")

# ==========================================
# ETAPA 3: CIERRE DE RUTA / REGRESO
# ==========================================
elif st.session_state.etapa_idx == 3:
    st.header("Etapa 3: Cierre de Ruta / Regreso a Patio Base")
    st.markdown(f"**Unidad:** `{st.session_state.get('unidad_activa', 'C-24')}` | **Operador:** `{st.session_state.get('operador_activo', '')}`")
    st.markdown("---")
    
    col_e3_1, col_e3_2 = st.columns(2)
    
    with col_e3_1:
        st.subheader("Captura de Cierre")
        km_final_in = st.number_input("KM Final del Viaje (Llegada Patio):", value=float(st.session_state.get("km_final_viaje", 0.0)), step=1.0, key="km_final_viaje_input")

        if st.button("Registrar Cierre de Ruta", key="btn_cierre_ruta"):
            if km_final_in < st.session_state.get("km_salida_cliente", 0.0):
                st.error("⚠️ El KM Final no puede ser menor al KM de Salida del Cliente.")
            else:
                st.session_state.km_final_viaje = km_final_in
                st.session_state.hora_cierre_viaje = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                tiempo_total_str = "N/A"
                if st.session_state.get("hora_salida_patio"):
                    try:
                        dt_salida = datetime.strptime(st.session_state.get("hora_salida_patio"), "%Y-%m-%d %H:%M:%S")
                        dt_cierre = datetime.strptime(st.session_state.get("hora_cierre_viaje"), "%Y-%m-%d %H:%M:%S")
                        diff = dt_cierre - dt_salida
                        horas = int(diff.total_seconds() // 3600)
                        minutos = int((diff.total_seconds() % 3600) // 60)
                        tiempo_total_str = f"{horas} hrs {minutos} mins"
                    except Exception:
                        pass

                guardar_en_excel({
                    "Fecha/Hora Cierre": st.session_state.get("hora_cierre_viaje", datetime.now()),
                    "Operador": st.session_state.get('operador_activo', ''),
                    "Unidad": st.session_state.get('unidad_activa', ''),
                    "Km Salida Patio": st.session_state.get("km_salida_patio", 0),
                    "Hora Salida Patio": st.session_state.get("hora_salida_patio", ""),
                    "Km Llegada Pedrera": st.session_state.get("km_llegada_pedrera", 0),
                    "Hora Llegada Pedrera": st.session_state.get("hora_llegada_pedrera", ""),
                    "Km Salida Pedrera": st.session_state.get("km_salida_pedrera", 0),
                    "Hora Salida Pedrera": st.session_state.get("hora_salida_pedrera", ""),
                    "Km Llegada Cliente": st.session_state.get("km_llegada_cliente", 0),
                    "Hora Llegada Cliente": st.session_state.get("hora_llegada_cliente", ""),
                    "Km Salida Cliente": st.session_state.get("km_salida_cliente", 0),
                    "Hora Salida Cliente": st.session_state.get("hora_salida_cliente", ""),
                    "Km Final Viaje": km_final_in,
                    "Fecha/Hora Llegada Patio": st.session_state.get("hora_cierre_viaje", ""),
                    "Distancia Total (KM)": km_final_in - st.session_state.get("km_salida_patio", 0),
                    "Tiempo Total Viaje": tiempo_total_str,
                    "Estado": "Ruta Finalizada / Cierre Exitoso",
                })
                st.success("✅ Cierre de ruta registrado exitosamente y guardado en Excel.")
                st.rerun()

        if st.session_state.get("hora_cierre_viaje"):
            st.markdown("---")
            if st.button("🔄 Iniciar Nueva Ruta (Volver a Etapa 1)", type="primary", key="btn_nueva_ruta"):
                ultimo_km = st.session_state.get("km_final_viaje", 145826.0)
                
                st.session_state.km_salida_patio = ultimo_km
                st.session_state.hora_salida_patio = ""
                st.session_state.km_llegada_pedrera = ultimo_km
                st.session_state.hora_llegada_pedrera = ""
                st.session_state.km_salida_pedrera = ultimo_km
                st.session_state.hora_salida_pedrera = ""
                st.session_state.km_llegada_cliente = ultimo_km
                st.session_state.hora_llegada_cliente = ""
                st.session_state.km_salida_cliente = ultimo_km
                st.session_state.hora_salida_cliente = ""
                st.session_state.km_final_viaje = ultimo_km
                st.session_state.hora_cierre_viaje = ""
                
                st.session_state.etapa_idx = 0
                st.rerun()

    with col_e3_2:
        st.subheader("Resumen de Cierre")
        if st.session_state.hora_cierre_viaje:
            distancia_total = st.session_state.km_final_viaje - st.session_state.km_salida_patio
            st.success(f"🏁 **Cierre:** {st.session_state.hora_cierre_viaje}")
            st.metric("📏 Distancia Total Recorrida", f"{distancia_total:.1f} KM")
            
            if st.session_state.get("hora_salida_patio"):
                try:
                    dt_salida = datetime.strptime(st.session_state.get("hora_salida_patio"), "%Y-%m-%d %H:%M:%S")
                    dt_cierre = datetime.strptime(st.session_state.get("hora_cierre_viaje"), "%Y-%m-%d %H:%M:%S")
                    diff = dt_cierre - dt_salida
                    horas = int(diff.total_seconds() // 3600)
                    minutos = int((diff.total_seconds() % 3600) // 60)
                    st.metric("⏱️ Tiempo Total del Viaje", f"{horas} hrs {minutos} mins")
                except Exception:
                    pass
