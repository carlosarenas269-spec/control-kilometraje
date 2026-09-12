import os
import streamlit as st
import pandas as pd
from datetime import datetime

# ==========================================
# CONFIGURACIÓN Y PERSISTENCIA DE DATOS
# ==========================================
ARCHIVO_REGISTROS = "registros_rutas.xlsx"
ARCHIVO_CATALOGOS = "catalogos_guardados.xlsx"

def asegurar_longitud_negativa(val):
    """Garantiza que la longitud sea negativa para que caiga en América/México y no en China."""
    try:
        fval = float(val)
        return -abs(fval)
    except Exception:
        return -100.293652

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
            for num_cli, sucursales in st.session_state.db_clientes.items():
                for suc_nombre, info in sucursales.items():
                    clientes_data.append({
                        "num_cliente": num_cli,
                        "nombre": suc_nombre,
                        "contacto": info.get("contacto", ""),
                        "telefono": info.get("telefono", ""),
                        "latitud": info.get("latitud", 25.844412),
                        "longitud": asegurar_longitud_negativa(info.get("longitud", -100.293652))
                    })
            pd.DataFrame(clientes_data).to_excel(writer, sheet_name="Clientes", index=False)
    except Exception as e:
        print(f"Error al guardar catálogos en disco: {e}")

def cargar_catalogos_de_disco():
    """Carga los catálogos desde el disco si el archivo existe, soportando num_cliente duplicado."""
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
                    num_cli = str(row["num_cliente"]) if "num_cliente" in row and pd.notna(row["num_cliente"]) else "S/N"
                    suc_nombre = str(row["nombre"]) if "nombre" in row and pd.notna(row["nombre"]) else "Principal"
                    
                    if num_cli not in lista_c:
                        lista_c.append(num_cli)
                    if num_cli not in nueva_db:
                        nueva_db[num_cli] = {}
                        
                    nueva_db[num_cli][suc_nombre] = {
                        "contacto": str(row["contacto"]) if "contacto" in row and pd.notna(row["contacto"]) else "",
                        "telefono": str(row["telefono"]) if "telefono" in row and pd.notna(row["telefono"]) else "",
                        "latitud": float(row["latitud"]) if "latitud" in row and pd.notna(row["latitud"]) else 25.844412,
                        "longitud": asegurar_longitud_negativa(row["longitud"] if "longitud" in row and pd.notna(row["longitud"]) else -100.293652)
                    }
                st.session_state.lista_clientes = lista_c
                st.session_state.db_clientes = nueva_db
        except Exception as e:
            print(f"Error al leer catálogos de disco: {e}")

# ==========================================
# CONFIGURACIÓN INICIAL DE LA APP
# ==========================================
st.set_page_config(page_title="Fletes AVE Control", layout="wide")

col_titulo, col_logo = st.columns([4, 1])

with col_titulo:
    st.title("🚛 Control de Fletes AVE")

with col_logo:
    try:
        st.image(r"C:\Users\CarlosArenas\Desktop\logo_fletes_ave.jpg", width=120)
    except Exception:
        try:
            st.image("logo_fletes_ave.jpg", width=120)
        except Exception:
            pass

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
    st.session_state.lista_clientes = ["3185", "1198", "1555"]

if "db_clientes" not in st.session_state:
    st.session_state.db_clientes = {
        "3185": {
            "HOLCIM MEXICO OPERACIONES (Escobedo)": {
                "contacto": "GUADALUPE",
                "telefono": "1",
                "latitud": 25.847237,
                "longitud": -100.293652
            },
            "HOLCIM MEXICO OPERACIONES (Litos)": {
                "contacto": "ING. IVAN",
                "telefono": "1",
                "latitud": 25.898032,
                "longitud": -100.200360
            }
        },
        "1198": {
            "ORGANIZACIÓN Y SERVICIO PARA LA CONSTRUCCION, S.A. DE C.V.": {
                "contacto": "",
                "telefono": "2381915291",
                "latitud": 25.829361,
                "longitud": -100.230889
            }
        },
        "1555": {
            "GRUPO PERFIMEXA SA DE CV": {
                "contacto": "",
                "telefono": "8115783462",
                "latitud": 26.004936,
                "longitud": -100.296356
            }
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
    ("cli_latitud", 25.844412), ("cli_longitud", -100.395043),
    ("ruta_guardada", False) # Bandera para controlar la visualización del botón de nueva ruta
]:
    if var not in st.session_state:
        st.session_state[var] = val

def actualizar_campos_cliente():
    c_sel = st.session_state.get("sel_cliente_e2")
    s_sel = st.session_state.get("sel_sucursal_e2")
    if c_sel in st.session_state.db_clientes and s_sel in st.session_state.db_clientes[c_sel]:
        info = st.session_state.db_clientes[c_sel][s_sel]
        st.session_state.cli_contacto = info.get("contacto", "")
        st.session_state.cli_telefono = str(info.get("telefono", ""))
        st.session_state.cli_latitud = float(info.get("latitud", 25.844412))
        st.session_state.cli_longitud = asegurar_longitud_negativa(info.get("longitud", -100.395043))

# ==========================================
# BARRA LATERAL: CONTROL DE ACCESO ADMIN
# ==========================================
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
# NAVEGACIÓN POR PESTAÑAS
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
        with st.expander("⚙️ Administración de Catálogos y Reportes"):
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            
            with col_m1:
                st.subheader("1. Alta Manual")
                nuevo_tipo = st.selectbox("¿Qué deseas agregar?", ["Unidad", "Operador", "Cliente / Dirección"], key="alta_tipo")
                
                if nuevo_tipo == "Cliente / Dirección":
                    num_cli_input = st.text_input("Número de Cliente (ej. 3185):", key="alta_num_cli")
                    nombre_dir_input = st.text_input("Nombre / Sucursal (ej. HOLCIM (Escobedo)):", key="alta_nom_dir")
                    contacto_input = st.text_input("Contacto:", key="alta_contacto")
                    tel_input = st.text_input("Teléfono:", key="alta_tel")
                    lat_input = st.number_input("Latitud:", value=25.847237, format="%.6f", key="alta_lat")
                    lon_input = st.number_input("Longitud (ej. -100.293652):", value=-100.293652, format="%.6f", key="alta_lon")
                else:
                    nuevo_nombre = st.text_input("Escribe el nombre o código:", key="alta_nombre")
                
                if st.button("➕ Agregar al Catálogo"):
                    if nuevo_tipo == "Cliente / Dirección":
                        if num_cli_input.strip() and nombre_dir_input.strip():
                            nc = num_cli_input.strip()
                            nd = nombre_dir_input.strip()
                            if nc not in st.session_state.lista_clientes:
                                st.session_state.lista_clientes.append(nc)
                            if nc not in st.session_state.db_clientes:
                                st.session_state.db_clientes[nc] = {}
                            st.session_state.db_clientes[nc][nd] = {
                                "contacto": contacto_input.strip(),
                                "telefono": tel_input.strip(),
                                "latitud": lat_input,
                                "longitud": asegurar_longitud_negativa(lon_input)
                            }
                            guardar_catalogos_en_disco()
                            st.success(f"Dirección agregada al cliente {nc} con éxito.")
                            st.rerun()
                        else:
                            st.warning("Completa el número de cliente y la sucursal.")
                    else:
                        if nuevo_nombre.strip():
                            val = nuevo_nombre.strip()
                            if nuevo_tipo == "Unidad" and val not in st.session_state.lista_unidades:
                                st.session_state.lista_unidades.append(val)
                            elif nuevo_tipo == "Operador" and val not in st.session_state.lista_operadores:
                                st.session_state.lista_operadores.append(val)
                            guardar_catalogos_en_disco()
                            st.success(f"'{val}' agregado con éxito.")
                            st.rerun()

            with col_m2:
                st.subheader("2. Carga Masiva (Excel)")
                archivo_excel = st.file_uploader("Sube tu archivo .xlsx (columnas num_cliente, nombre, contacto, telefono, latitud, longitud)", type=["xlsx", "xls"])
                if archivo_excel is not None:
                    try:
                        df = pd.read_excel(archivo_excel)
                        if st.button("📥 Importar Clientes desde Excel"):
                            if "num_cliente" in df.columns and "nombre" in df.columns:
                                nueva_db = {}
                                lista_c = []
                                for _, row in df.iterrows():
                                    nc = str(row["num_cliente"])
                                    nd = str(row["nombre"])
                                    if nc not in lista_c:
                                        lista_c.append(nc)
                                    if nc not in nueva_db:
                                        nueva_db[nc] = {}
                                    
                                    lon_val = row["longitud"] if "longitud" in df.columns and pd.notna(row["longitud"]) else -100.395043
                                    
                                    nueva_db[nc][nd] = {
                                        "contacto": str(row["contacto"]) if "contacto" in df.columns and pd.notna(row["contacto"]) else "",
                                        "telefono": str(row["telefono"]) if "telefono" in df.columns and pd.notna(row["telefono"]) else "",
                                        "latitud": float(row["latitud"]) if "latitud" in df.columns and pd.notna(row["latitud"]) else 25.844412,
                                        "longitud": asegurar_longitud_negativa(lon_val)
                                    }
                                st.session_state.lista_clientes = lista_c
                                st.session_state.db_clientes = nueva_db
                                guardar_catalogos_en_disco()
                                st.success("✅ Base de clientes importada correctamente con longitudes corregidas.")
                                st.rerun()
                            else:
                                st.error("El archivo Excel debe contener las columnas 'num_cliente' y 'nombre'.")
                    except Exception as e:
                        st.error(f"Error: {e}")

            with col_m3:
                st.subheader("3. Eliminar Registros")
                tipo_eliminar = st.selectbox("¿Qué deseas eliminar?", ["Unidad", "Operador", "Cliente completo"], key="del_tipo")
                item_eliminar = st.selectbox("Selecciona:", st.session_state.lista_unidades if tipo_eliminar=="Unidad" else (st.session_state.lista_operadores if tipo_eliminar=="Operador" else st.session_state.lista_clientes), key="del_item")
                
                if st.button("🗑️ Eliminar"):
                    if item_eliminar:
                        if tipo_eliminar == "Unidad": st.session_state.lista_unidades.remove(item_eliminar)
                        elif tipo_eliminar == "Operador": st.session_state.lista_operadores.remove(item_eliminar)
                        elif tipo_eliminar == "Cliente completo":
                            st.session_state.lista_clientes.remove(item_eliminar)
                            if item_eliminar in st.session_state.db_clientes: del st.session_state.db_clientes[item_eliminar]
                        guardar_catalogos_en_disco()
                        st.success("Eliminado correctamente.")
                        st.rerun()

            with col_m4:
                st.subheader("4. Descargar Reportes")
                if os.path.exists(ARCHIVO_REGISTROS):
                    with open(ARCHIVO_REGISTROS, "rb") as file:
                        st.download_button(
                            label="📥 Descargar Excel de Viajes",
                            data=file,
                            file_name="registros_rutas.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.info("Aún no hay registros guardados.")

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
            st.session_state.ruta_guardada = False # Reiniciar bandera de ruta guardada al iniciar una nueva
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
        # Selección del Número de Cliente
        cliente_sel = st.selectbox("Selecciona el Número de Cliente:", st.session_state.lista_clientes, key="sel_cliente_e2")
        
        # Obtener las direcciones (nombres) asociadas a ese número de cliente
        sucursales_dict = st.session_state.db_clientes.get(cliente_sel, {"Dirección Principal": {"contacto": "", "telefono": "", "latitud": 25.844412, "longitud": -100.395043}})
        lista_sucursales = list(sucursales_dict.keys())
        
        # Selección de la dirección específica
        sucursal_sel = st.selectbox("Selecciona la Dirección / Planta:", lista_sucursales, key="sel_sucursal_e2")
        
        # Actualizar campos automáticos
        actualizar_campos_cliente()
        
        info_actual = sucursales_dict.get(sucursal_sel, {})
        contacto = info_actual.get("contacto", "")
        telefono = info_actual.get("telefono", "")
        latitud = float(info_actual.get("latitud", 25.844412))
        longitud = asegurar_longitud_negativa(info_actual.get("longitud", -100.395043))
        
        st.markdown(f"**🏢 Cliente N°:** `{cliente_sel}`\n\n**📍 Planta/Destino:** `{sucursal_sel}`")
        
        st.text_input("Contacto:", value=contacto, disabled=True)
        st.text_input("Teléfono:", value=telefono, disabled=True)
        
        if telefono and telefono.isdigit() and len(telefono) > 1:
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
        with col_geo1: st.number_input("Latitud:", value=latitud, format="%.6f", disabled=True)
        with col_geo2: st.number_input("Longitud:", value=longitud, format="%.6f", disabled=True)
            
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
        km_final_in = st.number_input("KM Final del Viaje (Patio Llegada):", value=float(st.session_state.get("km_final_viaje", 0.0)), step=1.0, key="km_final_viaje_input")

        # ----------------------------------------------------
        # 1. CÁLCULO EN TIEMPO REAL: Kilómetros Totales del Viaje
        # ----------------------------------------------------
        km_salida_base = st.session_state.get("km_salida_patio", 0.0)
        km_recorridos = km_final_in - km_salida_base
        if km_recorridos < 0:
            km_recorridos = 0.0
            
        st.info(f"📊 **Kilómetros Totales del Viaje:** `{km_recorridos:,.1f} km` (Desde Salida Patio: {km_salida_base:,.1f})")

        # ----------------------------------------------------
        # 2. CÁLCULO EN TIEMPO REAL: Tiempo Total de Ruta
        # ----------------------------------------------------
        tiempo_total_str = "N/A"
        if st.session_state.get("hora_salida_patio"):
            try:
                dt_salida = datetime.strptime(st.session_state.get("hora_salida_patio"), "%Y-%m-%d %H:%M:%S")
                dt_actual = datetime.now()
                diff = dt_actual - dt_salida
                horas = int(diff.total_seconds() // 3600)
                minutos = int((diff.total_seconds() % 3600) // 60)
                tiempo_total_str = f"{horas} hrs {minutos} mins"
            except Exception:
                pass
                
        st.success(f"⏱️ **Tiempo Total Transcurrido:** `{tiempo_total_str}`")
        # ----------------------------------------------------

        if st.button("Registrar Cierre de Ruta", key="btn_cierre_ruta"):
            if km_final_in < st.session_state.get("km_salida_cliente", 0.0):
                st.error("⚠️ El KM Final no puede ser menor al KM de Salida del Cliente.")
            else:
                st.session_state.km_final_viaje = km_final_in
                st.session_state.hora_cierre_viaje = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                # Recalcular tiempo exacto al momento de guardar
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
                    "Num Cliente": st.session_state.get("sel_cliente_e2", ""),
                    "Planta / Destino": st.session_state.get("sel_sucursal_e2", ""),
                    "Km Salida Patio": km_salida_base,
                    "Hora Salida Patio": st.session_state.get("hora_salida_patio", ""),
                    "Km Llegada Pedrera": st.session_state.get("km_llegada_pedrera", 0),
                    "Hora Llegada Pedrera": st.session_state.get("hora_llegada_pedrera", ""),
                    "Km Salida Pedrera": st.session_state.get("km_salida_pedrera", ""),
                    "Hora Salida Pedrera": st.session_state.get("hora_salida_pedrera", ""),
                    "Km Llegada Cliente": st.session_state.get("km_llegada_cliente", 0),
                    "Hora Llegada Cliente": st.session_state.get("hora_llegada_cliente", ""),
                    "Km Salida Cliente": st.session_state.get("km_salida_cliente", 0),
                    "Hora Salida Cliente": st.session_state.get("hora_salida_cliente", ""),
                    "Km Final Patio": st.session_state.get("km_final_viaje", 0),
                    "Km Totales Ruta": km_recorridos,
                    "Hora Cierre Viaje": st.session_state.get("hora_cierre_viaje", ""),
                    "Tiempo Total": tiempo_total_str
                })
                
                st.session_state.ruta_guardada = True
                st.success("🎉 ¡Ruta guardada y cerrada con éxito en el archivo Excel!")
                st.balloons()
                st.rerun()

    with col_e3_2:
        st.subheader("Estado del Cierre")
        if st.session_state.get("ruta_guardada", False):
            st.success("✅ Ruta finalizada correctamente.")
            
            if st.button("🔄 Iniciar Nuevo Registro de Ruta", type="primary"):
                st.session_state.hora_salida_patio = ""
                st.session_state.hora_llegada_pedrera = ""
                st.session_state.hora_salida_pedrera = ""
                st.session_state.hora_llegada_cliente = ""
                st.session_state.hora_salida_cliente = ""
                st.session_state.hora_cierre_viaje = ""
                st.session_state.ruta_guardada = False
                st.session_state.etapa_idx = 0
                st.rerun()
        else:
            st.info("Complete el odómetro final y haga clic en 'Registrar Cierre de Ruta' para finalizar.")