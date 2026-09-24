import os
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone

# ==========================================
# 1. CONFIGURACIÓN Y PERSISTENCIA DE DATOS
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
    """Garantiza que la longitud sea negativa para que caiga en América/México."""
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
    """Guarda el estado actual de catálogos y coordenadas en el archivo Excel permanente."""
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
                        "latitud": info.get("latitud", 25.844412),
                        "longitud": asegurar_longitud_negativa(info.get("longitud", -100.293652))
                    })
            pd.DataFrame(clientes_data).to_excel(writer, sheet_name="Clientes", index=False)
    except Exception as e:
        print(f"Error al guardar catálogos en disco: {e}")

def cargar_catalogos_de_disco():
    """Carga los catálogos y coordenadas reales desde el disco si el archivo existe."""
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
                        num_cli = str(row.get("num_cliente", row.iloc[0] if len(row)>0 else "S/N"))
                        suc_nombre = str(row.get("nombre", row.iloc[1] if len(row)>1 else "Principal"))
                        
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
                                "latitud": lat_val,
                                "longitud": asegurar_longitud_negativa(lon_val)
                            }
                    if lista_c:
                        st.session_state.lista_clientes = lista_c
                        st.session_state.db_clientes = nueva_db
        except Exception as e:
            print(f"Error al leer catálogos de disco: {e}")

# ==========================================
# 2. CONFIGURACIÓN INICIAL DE LA APP Y ESTADOS
# ==========================================
st.set_page_config(page_title="Fletes AVE Control", layout="wide")

st.title("🚛 Control de Fletes AVE")

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
OPERADORES_DEFAULT = ["Octavio Rodrigo Serrano Saavedra", "Juan Carlos Arenas", "Pedro Martinez", "Jose Luis Rodriguez"]
CLIENTES_DEFAULT = ["3185", "1198", "1555"]
DB_CLIENTES_DEFAULT = {
    "3185": {
        "HOLCIM MEXICO OPERACIONES (Escobedo)": {"latitud": 25.847237, "longitud": -100.293652},
        "HOLCIM MEXICO OPERACIONES (Litos)": {"latitud": 25.898032, "longitud": -100.200360}
    },
    "1198": {
        "ORGANIZACIÓN Y SERVICIO PARA LA CONSTRUCCION": {"latitud": 25.829361, "longitud": -100.230889}
    },
    "1555": {
        "GRUPO PERFIMEXA SA DE CV": {"latitud": 26.004936, "longitud": -100.296356}
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

# Cargar desde disco al iniciar la app para mantener persistencia
if "catalogos_cargados" not in st.session_state:
    if os.path.exists(ARCHIVO_CATALOGOS):
        cargar_catalogos_de_disco()
    else:
        guardar_catalogos_en_disco()
    st.session_state.catalogos_cargados = True

# Variables de sesión para coordenadas activas
for var, val in [
    ("cli_latitud", 25.844412), ("cli_longitud", -100.395043),
    ("unidad_activa", "C-24"), ("operador_activo", "")
]:
    if var not in st.session_state:
        st.session_state[var] = val

def actualizar_campos_cliente():
    """Actualiza dinámicamente la latitud y longitud según el cliente y sucursal seleccionados."""
    c_sel = st.session_state.get("sel_cliente_e2")
    s_sel = st.session_state.get("sel_sucursal_e2")
    if c_sel in st.session_state.db_clientes and s_sel in st.session_state.db_clientes[c_sel]:
        info = st.session_state.db_clientes[c_sel][s_sel]
        st.session_state.cli_latitud = float(info.get("latitud", 25.844412))
        st.session_state.cli_longitud = asegurar_longitud_negativa(info.get("longitud", -100.395043))

# ==========================================
# 3. BARRA LATERAL: ADMIN (¡DEBE IR AQUÍ ARRIBA!)
# ==========================================
st.sidebar.title("🔒 Acceso Administrador")
pin_input = st.sidebar.text_input("Ingresa PIN de Administrador:", type="password")
es_admin = (pin_input == PIN_ADMIN)

if es_admin:
    st.sidebar.success("🔑 Modo Administrador Activo")
else:
    if pin_input != "":
        st.sidebar.error("❌ PIN Incorrecto")

# ==========================================
# 4. NAVEGACIÓN POR PESTAÑAS
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
# ETAPA 1: SALIDA DE PATIO BASE (CARGA MASIVA LIMPIA)
# ==========================================
if st.session_state.etapa_idx == 0:
    st.header("Etapa 1: Salida de Patio Base y Carga de Clientes")

    if es_admin:
        with st.expander("📂 Sube o Actualiza tu Archivo de Clientes con Coordenadas"):
            st.markdown("Sube tu archivo `.xlsx` o `.csv` para actualizar el catálogo de clientes de forma permanente.")
            
            archivo_excel = st.file_uploader("Sube tu archivo de clientes", type=["xlsx", "xls", "csv"], key="uploader_masivo_clientes")
            
            if archivo_excel is not None:
                try:
                    if archivo_excel.name.endswith('.csv'):
                        df_preview = pd.read_csv(archivo_excel)
                    else:
                        df_preview = pd.read_excel(archivo_excel)
                        
                    if not df_preview.empty:
                        columnas_disponibles = list(df_preview.columns)
                        
                        def buscar_indice(palabras_clave):
                            for idx, col in enumerate(columnas_disponibles):
                                if any(p in str(col).lower() for p in palabras_clave):
                                    return idx
                            return 0

                        st.markdown("🗺️ **Mapeo Automático de Columnas:**")
                        c1, c2 = st.columns(2)
                        with c1:
                            col_cli = st.selectbox("Columna de Cliente / ID:", columnas_disponibles, index=buscar_indice(["cliente", "id", "num"]))
                            col_dir = st.selectbox("Columna de Sucursal / Nombre:", columnas_disponibles, index=buscar_indice(["nombre", "sucursal", "direccion"]))
                        with c2:
                            col_lat = st.selectbox("Columna de LATITUD:", columnas_disponibles, index=buscar_indice(["lat"]))
                            col_lon = st.selectbox("Columna de LONGITUD:", columnas_disponibles, index=buscar_indice(["lon"]))
                        
                        st.markdown("")
                        if st.button("📥 Procesar y Cargar Archivo"):
                            nueva_db = {}
                            lista_c = []
                            contador_registros = 0
                            
                            for _, row in df_preview.iterrows():
                                nc = str(row[col_cli]).strip() if pd.notna(row[col_cli]) else ""
                                nd = str(row[col_dir]).strip() if pd.notna(row[col_dir]) else "Principal"
                                
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
                                        "latitud": lat_val,
                                        "longitud": asegurar_longitud_negativa(lon_val)
                                    }
                                    contador_registros += 1

                            if lista_c:
                                st.session_state.lista_clientes = lista_c
                                st.session_state.db_clientes = nueva_db
                                guardar_catalogos_en_disco()
                                
                                # Confirmación clara solicitada
                                st.success(f"✅ ¡Archivo cargado y agregado con éxito! Se procesaron **{len(lista_c)} clientes** y **{contador_registros} sucursales/ubicaciones** correctamente.")
                except Exception as e:
                    st.error(f"Error al procesar el archivo: {e}")

    st.success("Configuración cargada correctamente. Selecciona tu unidad y operador para continuar.")
    st.session_state.unidad_activa = st.selectbox("Unidad:", st.session_state.lista_unidades)
    st.session_state.operador_activo = st.selectbox("Operador:", st.session_state.lista_operadores)

# ==========================================
# ETAPA 2: LLEGADA / SALIDA CLIENTE
# ==========================================
elif st.session_state.etapa_idx == 2:
    st.header("Etapa 2: Llegada / Salida con Cliente")

    if st.session_state.lista_clientes:
        sel_cliente = st.selectbox("Selecciona el Cliente:", st.session_state.lista_clientes, key="sel_cliente_e2", on_change=actualizar_campos_cliente)
        
        sucursales_disponibles = list(st.session_state.db_clientes.get(sel_cliente, {}).keys())
        if sucursales_disponibles:
            sel_sucursal = st.selectbox("Selecciona la Sucursal / Dirección:", sucursales_disponibles, key="sel_sucursal_e2", on_change=actualizar_campos_cliente)
            
            actualizar_campos_cliente()

            st.info(f"📍 **Coordenadas Actuales del Cliente Seleccionado:**\n- **Latitud:** `{st.session_state.cli_latitud}`\n- **Longitud:** `{st.session_state.cli_longitud}`")
        else:
            st.warning("Este cliente no tiene sucursales registradas.")
    else:
        st.warning("No hay clientes cargados. Sube tu archivo Excel en la Etapa 1.")

# ==========================================
# OTRAS ETAPAS (Intermedia y Cierre)
# ==========================================
else:
    st.info("Navegando en la etapa seleccionada. Utiliza las pestañas superiores para cambiar de sección.")
