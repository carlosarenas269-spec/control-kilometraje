import os
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone

# ==========================================
# CONFIGURACIÓN Y PERSISTENCIA DE DATOS
# ==========================================
ARCHIVO_REGISTROS = "registros_rutas.xlsx"
ARCHIVO_CATALOGOS = "catalogos_guardados.xlsx"
ARCHIVO_INCIDENCIAS = "registros_incidencias.xlsx"
ARCHIVO_BORRADOR = "borrador_viaje_activo.xlsx"

def obtener_hora_mexico():
    """Obtiene la hora actual ajustada estrictamente a la zona horaria de México (CST / UTC-6)."""
    tz_mexico = timezone(timedelta(hours=-6))
    return datetime.now(tz_mexico).strftime("%Y-%m-%d %H:%M:%S")

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

def guardar_incidencia_en_excel(data_dict):
    """Guarda las incidencias reportadas en un archivo Excel independiente."""
    try:
        if os.path.exists(ARCHIVO_INCIDENCIAS):
            df_existente = pd.read_excel(ARCHIVO_INCIDENCIAS)
            df_nuevo = pd.DataFrame([data_dict])
            df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
        else:
            df_final = pd.DataFrame([data_dict])

        df_final.to_excel(ARCHIVO_INCIDENCIAS, index=False)
    except Exception as e:
        st.error(f"Error al guardar la incidencia en Excel: {e}")

def guardar_borrador_en_disco():
    """Guarda el estado actual del viaje en un archivo temporal para evitar pérdidas por recargas."""
    try:
        datos_borrador = {
            "etapa_idx": st.session_state.get("etapa_idx", 0),
            "unidad_activa": st.session_state.get("unidad_activa", "C-24"),
            "operador_activo": st.session_state.get("operador_activo", ""),
            "km_salida_patio": st.session_state.get("km_salida_patio", 145826.0),
            "hora_salida_patio": st.session_state.get("hora_salida_patio", ""),
            "km_llegada_pedrera": st.session_state.get("km_llegada_pedrera", 145826.0),
            "hora_llegada_pedrera": st.session_state.get("hora_llegada_pedrera", ""),
            "km_salida_pedrera": st.session_state.get("km_salida_pedrera", 145826.0),
            "hora_salida_pedrera": st.session_state.get("hora_salida_pedrera", ""),
            "km_llegada_cliente": st.session_state.get("km_llegada_cliente", 145826.0),
            "hora_llegada_cliente": st.session_state.get("hora_llegada_cliente", ""),
            "km_salida_cliente": st.session_state.get("km_salida_cliente", 145826.0),
            "hora_salida_cliente": st.session_state.get("hora_salida_cliente", ""),
            "km_final_viaje": st.session_state.get("km_final_viaje", 145826.0),
            "hora_cierre_viaje": st.session_state.get("hora_cierre_viaje", ""),
            "ruta_guardada": st.session_state.get("ruta_guardada", False)
        }
        pd.DataFrame([datos_borrador]).to_excel(ARCHIVO_BORRADOR, index=False)
    except Exception as e:
        print(f"Error al guardar borrador: {e}")

def cargar_borrador_de_disco():
    """Restaura el avance del viaje si el archivo temporal existe."""
    if os.path.exists(ARCHIVO_BORRADOR):
        try:
            df_b = pd.read_excel(ARCHIVO_BORRADOR)
            if not df_b.empty:
                row = df_b.iloc[0]
                st.session_state.etapa_idx = int(row.get("etapa_idx", 0))
                st.session_state.unidad_activa = str(row.get("unidad_activa", "C-24"))
                st.session_state.operador_activo = str(row.get("operador_activo", ""))
                st.session_state.km_salida_patio = float(row.get("km_salida_patio", 145826.0))
                st.session_state.hora_salida_patio = str(row.get("hora_salida_patio", "")) if pd.notna(row.get("hora_salida_patio")) else ""
                st.session_state.km_llegada_pedrera = float(row.get("km_llegada_pedrera", 145826.0))
                st.session_state.hora_llegada_pedrera = str(row.get("hora_llegada_pedrera", "")) if pd.notna(row.get("hora_llegada_pedrera")) else ""
                st.session_state.km_salida_pedrera = float(row.get("km_salida_pedrera", 145826.0))
                st.session_state.hora_salida_pedrera = str(row.get("hora_salida_pedrera", "")) if pd.notna(row.get("hora_salida_pedrera")) else ""
                st.session_state.km_llegada_cliente = float(row.get("km_llegada_cliente", 145826.0))
                st.session_state.hora_llegada_cliente = str(row.get("hora_llegada_cliente", "")) if pd.notna(row.get("hora_llegada_cliente")) else ""
                st.session_state.km_salida_cliente = float(row.get("km_salida_cliente", 145826.0))
                st.session_state.hora_salida_cliente = str(row.get("hora_salida_cliente", "")) if pd.notna(row.get("hora_salida_cliente")) else ""
                st.session_state.km_final_viaje = float(row.get("km_final_viaje", 145826.0))
                st.session_state.hora_cierre_viaje = str(row.get("hora_cierre_viaje", "")) if pd.notna(row.get("hora_cierre_viaje")) else ""
                st.session_state.ruta_guardada = bool(row.get("ruta_guardada", False))
        except Exception as e:
            print(f"Error al leer borrador: {e}")

def guardar_catalogos_en_disco():
    """Guarda exactamente el estado actual de la sesión en el archivo Excel permanente."""
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
    """Carga los catálogos desde el disco si el archivo existe."""
    if os.path.exists(ARCHIVO_CATALOGOS):
        try:
            xl = pd.ExcelFile(ARCHIVO_CATALOGOS)
            
            if "Unidades" in xl.sheet_names:
                df_u = pd.read_excel(ARCHIVO_CATALOGOS, sheet_name="Unidades")
                col_u = "Unidades" if "Unidades" in df_u.columns else df_u.columns[0]
                unidades_disco = df_u[col_u].dropna().astype(str).tolist()
                if unidades_disco:
                    st.session_state.lista_unidades = unidades_disco
            
            if "Operadores" in xl.sheet_names:
                df_o = pd.read_excel(ARCHIVO_CATALOGOS, sheet_name="Operadores")
                col_o = "Operadores" if "Operadores" in df_o.columns else df_o.columns[0]
                operadores_disco = df_o[col_o].dropna().astype(str).tolist()
                if operadores_disco:
                    st.session_state.lista_operadores = operadores_disco
                
            if "Clientes" in xl.sheet_names:
                df_c = pd.read_excel(ARCHIVO_CATALOGOS, sheet_name="Clientes")
                if not df_c.empty:
                    nueva_db = {}
                    lista_c = []
                    for _, row in df_c.iterrows():
                        num_cli = str(row.get("num_cliente", row.get("ID_Cliente", row.iloc[0] if len(row)>0 else "S/N")))
                        suc_nombre = str(row.get("nombre", row.get("Nombre_Cliente", row.iloc[1] if len(row)>1 else "Principal")))
                        
                        if num_cli and num_cli.lower() != 'nan':
                            if num_cli not in lista_c:
                                lista_c.append(num_cli)
                            if num_cli not in nueva_db:
                                nueva_db[num_cli] = {}
                                
                            try:
                                lat_val = float(row.get("latitud", 25.844412))
                                if pd.isna(lat_val): lat_val = 25.844412
                            except:
                                lat_val = 25.844412

                            try:
                                lon_val = float(row.get("longitud", -100.395043))
                                if pd.isna(lon_val): lon_val = -100.395043
                            except:
                                lon_val = -100.395043

                            nueva_db[num_cli][suc_nombre] = {
                                "contacto": str(row.get("contacto", "")) if pd.notna(row.get("contacto", "")) else "",
                                "telefono": str(row.get("telefono", "")) if pd.notna(row.get("telefono", "")) else "",
                                "latitud": lat_val,
                                "longitud": asegurar_longitud_negativa(lon_val)
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

# Valores predeterminados base
UNIDADES_DEFAULT = ["C-24", "C-25", "C-26", "T-01", "T-02"]
OPERADORES_DEFAULT = [
    "Octavio Rodrigo Serrano Saavedra",
    "Juan Carlos Arenas",
    "Pedro Martinez",
    "Jose Luis Rodriguez"
]
CLIENTES_DEFAULT = ["3185", "1198", "1555"]
DB_CLIENTES_DEFAULT = {
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

if "lista_unidades" not in st.session_state:
    st.session_state.lista_unidades = list(UNIDADES_DEFAULT)

if "lista_operadores" not in st.session_state:
    st.session_state.lista_operadores = list(OPERADORES_DEFAULT)

if "lista_clientes" not in st.session_state:
    st.session_state.lista_clientes = list(CLIENTES_DEFAULT)

if "db_clientes" not in st.session_state:
    st.session_state.db_clientes = dict(DB_CLIENTES_DEFAULT)

# Cargar desde disco al iniciar la app
if "catalogos_cargados" not in st.session_state:
    if os.path.exists(ARCHIVO_CATALOGOS):
        cargar_catalogos_de_disco()
    else:
        guardar_catalogos_en_disco()
    cargar_borrador_de_disco()
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
    ("ruta_guardada", False),
    ("unidad_activa", "C-24"),
    ("operador_activo", "")
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
# BARRA LATERAL: CONTROL DE ACCESO ADMIN & INCIDENCIAS
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

st.sidebar.markdown("---")

st.sidebar.subheader("⚠️ Reporte de Incidencias")
with st.sidebar.expander("Registrar Incidencia / Parada"):
    tipo_incidencia = st.selectbox(
        "Tipo de Incidencia:", 
        ["Accidente", "Retraso", "Falla mecánica", "Ponchadura", "Hora de comida", "Otro"]
    )
    comentario_incidencia = st.text_area("Detalles / Comentarios:")
    
    if st.button("🚨 Guardar Incidencia"):
        ahora_inc = obtener_hora_mexico()
        data_inc = {
            "Fecha/Hora": ahora_inc,
            "Unidad": st.session_state.get("unidad_activa", "N/D"),
            "Operador": st.session_state.get("operador_activo", "N/D"),
            "Tipo Incidencia": tipo_incidencia,
            "Detalles": comentario_incidencia
        }
        guardar_incidencia_en_excel(data_inc)
        st.sidebar.success("✅ Incidencia registrada correctamente.")

if es_admin:
    if os.path.exists(ARCHIVO_INCIDENCIAS):
        with open(ARCHIVO_INCIDENCIAS, "rb") as file:
            st.sidebar.download_button(
                label="📥 Descargar Excel de Incidencias",
                data=file,
                file_name="registros_incidencias.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

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
                    # Uso de text_input para evitar ceros automáticos o alteraciones al pegar
                    lat_str = st.text_input("Latitud (ej. 25.765219):", value="25.847237", key="alta_lat")
                    lon_str = st.text_input("Longitud (ej. -100.293652):", value="-100.293652", key="alta_lon")
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
                            
                            # Conversión segura desde texto a float
                            try:
                                lat_val = float(lat_str.strip())
                            except ValueError:
                                lat_val = 25.847237
                                
                            try:
                                lon_val = float(lon_str.strip())
                            except ValueError:
                                lon_val = -100.293652

                            st.session_state.db_clientes[nc][nd] = {
                                "contacto": contacto_input.strip(),
                                "telefono": tel_input.strip(),
                                "latitud": lat_val,
                                "longitud": asegurar_longitud_negativa(lon_val)
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
                st.subheader("2. Carga Masiva Inteligente")
                tipo_carga = st.selectbox("Selecciona qué contiene tu Excel:", ["Clientes con Sucursales", "Lista de Operadores", "Lista de Unidades"], key="tipo_carga_excel")
                archivo_excel = st.file_uploader("Sube tu archivo .xlsx o .xls", type=["xlsx", "xls"], key="uploader_masivo")
                
                if archivo_excel is not None:
                    try:
                        df_preview = pd.read_excel(archivo_excel)
                        if not df_preview.empty:
                            if tipo_carga == "Clientes con Sucursales":
                                st.markdown("##### 🗺️ Mapeo Automático de Columnas:")
                                columnas_disponibles = list(df_preview.columns)
                                
                                def buscar_indice(palabras_clave):
                                    for idx, col in enumerate(columnas_disponibles):
                                        c_low = str(col).lower()
                                        if any(p in c_low for p in palabras_clave):
                                            return idx
                                    return 0

                                idx_cli = buscar_indice(["cliente", "num", "id", "no_"])
                                idx_dir = buscar_indice(["nombre", "sucursal", "direccion", "planta", "destino"])
                                idx_cont = buscar_indice(["contacto"])
                                idx_tel = buscar_indice(["telefono", "tel", "cel"])
                                idx_lat = buscar_indice(["lat"])
                                idx_lon = buscar_indice(["lon", "long"])

                                c_map1, c_map2 = st.columns(2)
                                with c_map1:
                                    col_cli = st.selectbox("Columna de Cliente / ID:", columnas_disponibles, index=idx_cli)
                                    col_dir = st.selectbox("Columna de Sucursal / Nombre:", columnas_disponibles, index=idx_dir)
                                    col_cont = st.selectbox("Columna de Contacto:", [None] + columnas_disponibles, index=(idx_cont + 1) if idx_cont < len(columnas_disponibles) else 0)
                                with c_map2:
                                    col_tel = st.selectbox("Columna de Teléfono:", [None] + columnas_disponibles, index=(idx_tel + 1) if idx_tel < len(columnas_disponibles) else 0)
                                    col_lat = st.selectbox("Columna de LATITUD:", columnas_disponibles, index=idx_lat)
                                    col_lon = st.selectbox("Columna de LONGITUD:", columnas_disponibles, index=idx_lon)
                                
                                if st.button("📥 Procesar e Importar Clientes con Coordenadas"):
                                    nueva_db = {}
                                    lista_c = []
                                    for _, row in df_preview.iterrows():
                                        nc = str(row[col_cli]).strip() if pd.notna(row[col_cli]) else ""
                                        nd = str(row[col_dir]).strip() if pd.notna(row[col_dir]) else "Principal"
                                        contacto_val = str(row[col_cont]).strip() if col_cont and pd.notna(row[col_cont]) else ""
                                        tel_val = str(row[col_tel]).strip() if col_tel and pd.notna(row[col_tel]) else ""
                                        
                                        try: lat_val = float(row[col_lat]) if pd.notna(row[col_lat]) else 25.844412
                                        except: lat_val = 25.844412
                                        
                                        try: lon_val = float(row[col_lon]) if pd.notna(row[col_lon]) else -100.395043
                                        except: lon_val = -100.395043

                                        if nc and nc.lower() != 'nan':
                                            if nc not in lista_c:
                                                lista_c.append(nc)
                                            if nc not in nueva_db:
                                                nueva_db[nc] = {}
                                            
                                            nueva_db[nc][nd] = {
                                                "contacto": contacto_val,
                                                "telefono": tel_val,
                                                "latitud": lat_val,
                                                "longitud": asegurar_longitud_negativa(lon_val)
                                            }

                                    if lista_c:
                                        st.session_state.lista_clientes = lista_c
                                        st.session_state.db_clientes = nueva_db
                                        guardar_catalogos_en_disco()
                                        st.success("✅ ¡Base de clientes y coordenadas importadas con éxito!")
                                        st.rerun()
                                    else:
                                        st.warning("No se encontraron registros válidos.")
                            
                            elif tipo_carga == "Lista de Operadores":
                                col_op = st.selectbox("Columna de Operadores:", list(df_preview.columns))
                                if st.button("📥 Importar Operadores"):
                                    nuevos_ops = []
                                    for _, row in df_preview.iterrows():
                                        v = str(row[col_op]).strip()
                                        if v and v.lower() != 'nan': nuevos_ops.append(v)
                                    if nuevos_ops:
                                        st.session_state.lista_operadores = nuevos_ops
                                        guardar_catalogos_en_disco()
                                        st.success("✅ Operadores actualizados.")
                                        st.rerun()

                            elif tipo_carga == "Lista de Unidades":
                                col_uni = st.selectbox("Columna de Unidades:", list(df_preview.columns))
                                if st.button("📥 Importar Unidades"):
                                    nuevas_uni = []
                                    for _, row in df_preview.iterrows():
                                        v = str(row[col_uni]).strip()
                                        if v and v.lower() != 'nan': nuevas_uni.append(v)
                                    if nuevas_uni:
                                        st.session_state.lista_unidades = nuevas_uni
                                        guardar_catalogos_en_disco()
                                        st.success("✅ Unidades actualizadas.")
                                        st.rerun()
                        else:
                            st.warning("El archivo está vacío.")
                    except Exception as e:
                        st.error(f"Error al leer el archivo: {e}")

            with col_m3:
                st.subheader("3. Eliminar Catálogos")
                tipo_eliminar = st.selectbox("¿Qué deseas eliminar?", ["Unidad", "Operador", "Cliente completo"], key="del_tipo")
                item_eliminar = st.selectbox("Selecciona:", st.session_state.lista_unidades if tipo_eliminar=="Unidad" else (st.session_state.lista_operadores if tipo_eliminar=="Operador" else st.session_state.lista_clientes), key="del_item")
                
                if st.button("🗑️ Eliminar de Catálogo"):
                    if item_eliminar:
                        if tipo_eliminar == "Unidad":
                            if item_eliminar in st.session_state.lista_unidades:
                                st.session_state.lista_unidades.remove(item_eliminar)
                        elif tipo_eliminar == "Operador":
                            if item_eliminar in st.session_state.lista_operadores:
                                st.session_state.lista_operadores.remove(item_eliminar)
                        elif tipo_eliminar == "Cliente completo":
                            if item_eliminar in st.session_state.lista_clientes:
                                st.session_state.lista_clientes.remove(item_eliminar)
                            if item_eliminar in st.session_state.db_clientes:
                                del st.session_state.db_clientes[item_eliminar]
                        
                        guardar_catalogos_en_disco()
                        st.success(f"'{item_eliminar}' eliminado y guardado correctamente.")
                        st.rerun()
                
                st.markdown("---")
                if st.button("🔄 Forzar Limpieza y Reinicio de Base de Datos"):
                    if os.path.exists(ARCHIVO_CATALOGOS):
                        try:
                            os.remove(ARCHIVO_CATALOGOS)
                        except:
                            pass
                    st.session_state.lista_unidades = list(UNIDADES_DEFAULT)
                    st.session_state.lista_operadores = list(OPERADORES_DEFAULT)
                    st.session_state.lista_clientes = list(CLIENTES_DEFAULT)
                    st.session_state.db_clientes = dict(DB_CLIENTES_DEFAULT)
                    guardar_catalogos_en_disco()
                    st.success("✅ Base de datos reiniciada con éxito.")
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
        st.subheader("📊 Historial de Viajes Registrados (Vista Administrador)")
        if os.path.exists(ARCHIVO_REGISTROS):
            try:
                df_historial = pd.read_excel(ARCHIVO_REGISTROS)
                if not df_historial.empty:
                    st.dataframe(df_historial, use_container_width=True)
                    
                    col_dl1, col_dl2 = st.columns(2)
                    with col_dl1:
                        with open(ARCHIVO_REGISTROS, "rb") as f_excel:
                            st.download_button(
                                label="📥 Descargar Historial en Excel (.xlsx)",
                                data=f_excel,
                                file_name="registros_rutas.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="btn_dl_excel_main"
                            )
                    with col_dl2:
                        csv_bytes = df_historial.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            label="📥 Descargar Historial en CSV (.csv)",
                            data=csv_bytes,
                            file_name="registros_rutas.csv",
                            mime="text/csv",
                            key="btn_dl_csv_main"
                        )
                    
                    if st.button("🗑️ Borrar Todo el Historial de Viajes", type="primary"):
                        os.remove(ARCHIVO_REGISTROS)
                        st.success("Historial eliminado correctamente.")
                        st.rerun()
            except Exception as e:
                st.error(f"Error al leer el historial: {e}")
        else:
            st.info("No hay historial de viajes registrado todavía.")

    st.markdown("---")
    st.subheader("🚚 Datos Generales del Viaje")

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        unidad_sel = st.selectbox(
            "Selecciona la Unidad Activa:", 
            st.session_state.lista_unidades, 
            index=st.session_state.lista_unidades.index(st.session_state.unidad_activa) if st.session_state.unidad_activa in st.session_state.lista_unidades else 0
        )
        st.session_state.unidad_activa = unidad_sel

    with col_u2:
        operador_sel = st.selectbox(
            "Selecciona el Operador:", 
            st.session_state.lista_operadores, 
            index=st.session_state.lista_operadores.index(st.session_state.operador_activo) if st.session_state.operador_activo in st.session_state.lista_operadores else 0
        )
        st.session_state.operador_activo = operador_sel

    st.markdown("---")
    st.subheader("1️⃣ Odómetro y Hora: Salida de Patio Base")

    col_s1, col_s2 = st.columns(2)
    with col_s1:
        st.session_state.km_salida_patio = st.number_input("Kilometraje Salida Patio:", value=float(st.session_state.km_salida_patio), step=1.0, format="%.1f")
    with col_s2:
        if not st.session_state.hora_salida_patio:
            if st.button("⏱️ Registrar Hora Salida Patio"):
                st.session_state.hora_salida_patio = obtener_hora_mexico()
                guardar_borrador_en_disco()
                st.success(f"Registrado: {st.session_state.hora_salida_patio}")
                st.rerun()
        st.write(f"**Hora Salida Patio:** {st.session_state.hora_salida_patio if st.session_state.hora_salida_patio else 'Pendiente'}")

    if st.session_state.hora_salida_patio:
        if st.button("➡️ Avanzar a Etapa Intermedia (Pedrera)"):
            st.session_state.etapa_idx = 1
            guardar_borrador_en_disco()
            st.rerun()

# ==========================================
# ETAPA INTERMEDIA: PEDRERA
# ==========================================
elif st.session_state.etapa_idx == 1:
    st.header("Etapa Intermedia: Pedrera (Llegada y Salida)")

    col_p1, col_p2 = st.columns(2)
    with col_p1:
        st.subheader("📥 Llegada a Pedrera")
        st.session_state.km_llegada_pedrera = st.number_input("KM Llegada Pedrera:", value=float(st.session_state.km_llegada_pedrera), step=1.0, format="%.1f")
        if not st.session_state.hora_llegada_pedrera:
            if st.button("⏱️ Registrar Hora Llegada Pedrera"):
                st.session_state.hora_llegada_pedrera = obtener_hora_mexico()
                guardar_borrador_en_disco()
                st.success(f"Registrado: {st.session_state.hora_llegada_pedrera}")
                st.rerun()
        st.write(f"**Hora:** {st.session_state.hora_llegada_pedrera if st.session_state.hora_llegada_pedrera else 'Pendiente'}")

    with col_p2:
        st.subheader("📤 Salida de Pedrera")
        st.session_state.km_salida_pedrera = st.number_input("KM Salida Pedrera:", value=float(st.session_state.km_salida_pedrera), step=1.0, format="%.1f")
        if not st.session_state.hora_salida_pedrera:
            if st.button("⏱️ Registrar Hora Salida Pedrera"):
                st.session_state.hora_salida_pedrera = obtener_hora_mexico()
                guardar_borrador_en_disco()
                st.success(f"Registrado: {st.session_state.hora_salida_pedrera}")
                st.rerun()
        st.write(f"**Hora:** {st.session_state.hora_salida_pedrera if st.session_state.hora_salida_pedrera else 'Pendiente'}")

    col_nav1, col_nav2 = st.columns(2)
    with col_nav1:
        if st.button("⬅️ Volver a Etapa 1"):
            st.session_state.etapa_idx = 0
            guardar_borrador_en_disco()
            st.rerun()
    with col_nav2:
        if st.session_state.hora_salida_pedrera:
            if st.button("➡️ Avanzar a Etapa Cliente"):
                st.session_state.etapa_idx = 2
                guardar_borrador_en_disco()
                st.rerun()

# ==========================================
# ETAPA 2: LLEGADA / SALIDA CLIENTE
# ==========================================
elif st.session_state.etapa_idx == 2:
    st.header("Etapa 2: Llegada y Salida con el Cliente")

    col_cli1, col_cli2 = st.columns(2)
    with col_cli1:
        cliente_sel = st.selectbox("Selecciona Número de Cliente:", st.session_state.lista_clientes, key="sel_cliente_e2", on_change=actualizar_campos_cliente)
    
    with col_cli2:
        sucursales_disponibles = list(st.session_state.db_clientes.get(cliente_sel, {}).keys())
        sucursal_sel = st.selectbox("Selecciona Sucursal / Destino:", sucursales_disponibles, key="sel_sucursal_e2", on_change=actualizar_campos_cliente)

    # Actualizar datos de contacto y coordenadas actuales si cambian
    if cliente_sel in st.session_state.db_clientes and sucursal_sel in st.session_state.db_clientes[cliente_sel]:
        info_actual = st.session_state.db_clientes[cliente_sel][sucursal_sel]
        st.session_state.cli_contacto = info_actual.get("contacto", "")
        st.session_state.cli_telefono = str(info_actual.get("telefono", ""))
        st.session_state.cli_latitud = float(info_actual.get("latitud", 25.844412))
        st.session_state.cli_longitud = asegurar_longitud_negativa(info_actual.get("longitud", -100.395043))

    st.info(f"👤 **Contacto:** {st.session_state.cli_contacto if st.session_state.cli_contacto else 'N/D'} | 📞 **Teléfono:** {st.session_state.cli_telefono if st.session_state.cli_telefono else 'N/D'}")
    st.write(f"📍 **Coordenadas de Destino:** Lat: `{st.session_state.cli_latitud}` | Lon: `{st.session_state.cli_longitud}`")

    st.markdown("---")
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        st.subheader("📥 Llegada con Cliente")
        st.session_state.km_llegada_cliente = st.number_input("KM Llegada Cliente:", value=float(st.session_state.km_llegada_cliente), step=1.0, format="%.1f")
        if not st.session_state.hora_llegada_cliente:
            if st.button("⏱️ Registrar Hora Llegada Cliente"):
                st.session_state.hora_llegada_cliente = obtener_hora_mexico()
                guardar_borrador_en_disco()
                st.success(f"Registrado: {st.session_state.hora_llegada_cliente}")
                st.rerun()
        st.write(f"**Hora:** {st.session_state.hora_llegada_cliente if st.session_state.hora_llegada_cliente else 'Pendiente'}")

    with col_c2:
        st.subheader("📤 Salida con Cliente")
        st.session_state.km_salida_cliente = st.number_input("KM Salida Cliente:", value=float(st.session_state.km_salida_cliente), step=1.0, format="%.1f")
        if not st.session_state.hora_salida_cliente:
            if st.button("⏱️ Registrar Hora Salida Cliente"):
                st.session_state.hora_salida_cliente = obtener_hora_mexico()
                guardar_borrador_en_disco()
                st.success(f"Registrado: {st.session_state.hora_salida_cliente}")
                st.rerun()
        st.write(f"**Hora:** {st.session_state.hora_salida_cliente if st.session_state.hora_salida_cliente else 'Pendiente'}")

    col_nav3, col_nav4 = st.columns(2)
    with col_nav3:
        if st.button("⬅️ Volver a Etapa Intermedia"):
            st.session_state.etapa_idx = 1
            guardar_borrador_en_disco()
            st.rerun()
    with col_nav4:
        if st.session_state.hora_salida_cliente:
            if st.button("➡️ Avanzar a Cierre de Ruta"):
                st.session_state.etapa_idx = 3
                guardar_borrador_en_disco()
                st.rerun()

# ==========================================
# ETAPA 3: CIERRE DE RUTA / REGRESO
# ==========================================
elif st.session_state.etapa_idx == 3:
    st.header("Etapa 3: Cierre de Ruta / Regreso a Base")

    st.session_state.km_final_viaje = st.number_input("Kilometraje Final del Viaje (Regreso a Base):", value=float(st.session_state.km_final_viaje), step=1.0, format="%.1f")
    
    if not st.session_state.hora_cierre_viaje:
        if st.button("⏱️ Registrar Hora Cierre de Viaje"):
            st.session_state.hora_cierre_viaje = obtener_hora_mexico()
            guardar_borrador_en_disco()
            st.success(f"Registrado: {st.session_state.hora_cierre_viaje}")
            st.rerun()
    
    st.write(f"**Hora Cierre de Viaje:** {st.session_state.hora_cierre_viaje if st.session_state.hora_cierre_viaje else 'Pendiente'}")

    st.markdown("---")
    if not st.session_state.ruta_guardada:
        if st.button("💾 Guardar Viaje Completo en Historial", type="primary"):
            if st.session_state.hora_cierre_viaje:
                km_recorridos = st.session_state.km_final_viaje - st.session_state.km_salida_patio
                
                registro_final = {
                    "Fecha Registro": obtener_hora_mexico(),
                    "Unidad": st.session_state.unidad_activa,
                    "Operador": st.session_state.operador_activo,
                    "Cliente ID": st.session_state.get("sel_cliente_e2", "N/D"),
                    "Sucursal / Destino": st.session_state.get("sel_sucursal_e2", "N/D"),
                    "KM Salida Patio": st.session_state.km_salida_patio,
                    "Hora Salida Patio": st.session_state.hora_salida_patio,
                    "KM Llegada Pedrera": st.session_state.km_llegada_pedrera,
                    "Hora Llegada Pedrera": st.session_state.hora_llegada_pedrera,
                    "KM Salida Pedrera": st.session_state.km_salida_pedrera,
                    "Hora Salida Pedrera": st.session_state.hora_salida_pedrera,
                    "KM Llegada Cliente": st.session_state.km_llegada_cliente,
                    "Hora Llegada Cliente": st.session_state.hora_llegada_cliente,
                    "KM Salida Cliente": st.session_state.km_salida_cliente,
                    "Hora Salida Cliente": st.session_state.hora_salida_cliente,
                    "KM Final Viaje": st.session_state.km_final_viaje,
                    "Hora Cierre Viaje": st.session_state.hora_cierre_viaje,
                    "KM Totales Recorridos": km_recorridos,
                    "Latitud Destino": st.session_state.cli_latitud,
                    "Longitud Destino": st.session_state.cli_longitud
                }
                
                guardar_en_excel(registro_final, ARCHIVO_REGISTROS)
                st.session_state.ruta_guardada = True
                guardar_borrador_en_disco()
                st.success("✅ ¡Viaje guardado exitosamente en el historial!")
            else:
                st.warning("Debes registrar la hora de cierre del viaje antes de guardar.")
    else:
        st.success("✔️ Este viaje ya ha sido guardado en el historial.")
        if st.button("🔄 Iniciar Nuevo Viaje (Reiniciar Formulario)"):
            if os.path.exists(ARCHIVO_BORRADOR):
                try:
                    os.remove(ARCHIVO_BORRADOR)
                except:
                    pass
            st.session_state.etapa_idx = 0
            st.session_state.ruta_guardada = False
            st.session_state.hora_salida_patio = ""
            st.session_state.hora_llegada_pedrera = ""
            st.session_state.hora_salida_pedrera = ""
            st.session_state.hora_llegada_cliente = ""
            st.session_state.hora_salida_cliente = ""
            st.session_state.hora_cierre_viaje = ""
            st.rerun()

    if st.button("⬅️ Volver a Etapa Cliente"):
        st.session_state.etapa_idx = 2
        guardar_borrador_en_disco()
        st.rerun()
