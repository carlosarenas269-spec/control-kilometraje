import os
import json
import base64
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
from PIL import Image

# ==========================================
# CONFIGURACIÓN Y PERSISTENCIA DE DATOS
# ==========================================
ARCHIVO_REGISTROS = "registros_rutas.xlsx"
ARCHIVO_CATALOGOS = "catalogos_guardados.xlsx"
ARCHIVO_INCIDENCIAS = "registros_incidencias.xlsx"
ARCHIVO_BORRADOR = "borrador_viaje_activo.xlsx"
ARCHIVO_CONFIG = "db_config_rutas.json"
CARPETA_BANNER = "Banner_Principal"
CARPETA_LOGO = "Logo_Barra"
CARPETA_FOTOS_OPERADORES = "Fotos_Operadores"

os.makedirs(CARPETA_BANNER, exist_ok=True)
os.makedirs(CARPETA_LOGO, exist_ok=True)
os.makedirs(CARPETA_FOTOS_OPERADORES, exist_ok=True)

def cargar_configuracion():
    config_default = {
        "ancho_portada": 100, 
        "titulo_app": "Control de Fletes AVE",
        "texto_barra": "PANEL DE NAVEGACIÓN FLETES AVE",
        "texto_bienvenida": "✨ ¡Bienvenido al sistema operativo de rutas! Selecciona una opción en los botones superiores para continuar.",
        "alineacion_titulo": "Izquierda",
        "margen_superior": 0,
        "tamano_foto_operador": 120
    }
    if os.path.exists(ARCHIVO_CONFIG):
        try:
            with open(ARCHIVO_CONFIG, "r") as f:
                cfg = json.load(f)
                for k, v in config_default.items():
                    if k not in cfg:
                        cfg[k] = v
                return cfg
        except Exception:
            return config_default
    return config_default

def guardar_configuracion(config):
    try:
        with open(ARCHIVO_CONFIG, "w") as f:
            json.dump(config, f)
    except Exception:
        pass

config_actual = cargar_configuracion()

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

# Valores predeterminados base
UNIDADES_DEFAULT = ["C-24", "C-25", "C-26", "C-27", "C-28", "T-01", "T-02", "C-13"]
OPERADORES_DEFAULT = [
    "Octavio Rodrigo Serrano Saavedra",
    "Juan Carlos Arenas",
    "Pedro Martinez",
    "Jose Luis Rodriguez",
    "Luis Gonzalez Mendez",
    "Brandon Heriberto Medina Marroquin",
    "Oscar Garza Salas",
    "Cesar Dagoberto Galvan",
    "Mario Alberto Trejo"
]

ASIG_DEFAULT = {
    "C-24": "Octavio Rodrigo Serrano Saavedra",
    "C-25": "Juan Carlos Arenas",
    "C-26": "Luis Gonzalez Mendez",
    "C-27": "Juan Carlos Arenas",
    "C-28": "Brandon Heriberto Medina Marroquin",
    "T-01": "Pedro Martinez",
    "T-02": "Jose Luis Rodriguez",
    "C-13": "Oscar Garza Salas"
}

def guardar_borrador_de_disco():
    """Guarda el estado actual del viaje en un archivo temporal para evitar pérdidas por recargas."""
    try:
        datos_borrador = {
            "etapa_idx": st.session_state.get("etapa_idx", 0),
            "unidad_activa": st.session_state.get("unidad_activa", UNIDADES_DEFAULT[0]),
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
                st.session_state.unidad_activa = str(row.get("unidad_activa", UNIDADES_DEFAULT[0]))
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
            
            asig_data = [{"Unidad": u, "Operador": op} for u, op in st.session_state.asig_unidad_operador.items()]
            pd.DataFrame(asig_data).to_excel(writer, sheet_name="Asignacion", index=False)
            
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

            if "Asignacion" in xl.sheet_names:
                df_a = pd.read_excel(ARCHIVO_CATALOGOS, sheet_name="Asignacion")
                if not df_a.empty and "Unidad" in df_a.columns and "Operador" in df_a.columns:
                    st.session_state.asig_unidad_operador = {}
                    for _, row in df_a.iterrows():
                        u = str(row["Unidad"]).strip()
                        op = str(row["Operador"]).strip()
                        st.session_state.asig_unidad_operador[u] = op
            
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
# FUNCIONES DE CALLBACK REQUERIDAS ANTES DE SU USO
# ==========================================
def al_cambiar_unidad():
    """Actualiza automáticamente tanto la unidad activa como su operador correspondiente basado en el diccionario de asignaciones."""
    uni = st.session_state.get("sel_unidad_e1")
    if uni:
        st.session_state.unidad_activa = uni
        op_asignado = st.session_state.asig_unidad_operador.get(uni, st.session_state.lista_operadores[0])
        st.session_state.operador_activo = op_asignado
    guardar_borrador_de_disco()

# ==========================================
# CONFIGURACIÓN INICIAL DE LA APP
# ==========================================
st.set_page_config(page_title="Fletes AVE Control", layout="wide")

PIN_ADMIN = "1234"

# ==========================================
# BARRA LATERAL: CONTROL DE ACCESO ADMIN & DISEÑO
# ==========================================
st.sidebar.title("🔒 Acceso Administrador")
pin_input = st.sidebar.text_input("Ingresa PIN de Administrador:", type="password")
es_admin = (pin_input == PIN_ADMIN)

if es_admin:
    st.sidebar.success("🔑 Modo Administrador Activo")
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("🎨 Personalización y Estilos")
    
    with st.sidebar.form("form_admin_estilos"):
        nuevo_titulo = st.text_input("Título Principal:", value=config_actual.get("titulo_app", "Control de Fletes AVE"))
        nuevo_texto_barra = st.text_input("Texto Barra Navegación:", value=config_actual.get("texto_barra", "PANEL DE NAVEGACIÓN FLETES AVE"))
        nuevo_texto_bienv = st.text_area("Mensaje de Bienvenida:", value=config_actual.get("texto_bienvenida", "✨ ¡Bienvenido al sistema operativo de rutas!"))
        
        alineaciones = ["Izquierda", "Centro", "Derecha"]
        align_actual = config_actual.get("alineacion_titulo", "Izquierda")
        idx_align = alineaciones.index(align_actual) if align_actual in alineaciones else 0
        nueva_alineacion = st.selectbox("Alineación del Título:", options=alineaciones, index=idx_align)
        nuevo_margen = st.slider("Margen Superior (px):", min_value=0, max_value=50, value=int(config_actual.get("margen_superior", 0)), step=5)
        
        btn_save_textos = st.form_submit_button("💾 Guardar Cambios de Texto")
        if btn_save_textos:
            config_actual["titulo_app"] = nuevo_titulo
            config_actual["texto_barra"] = nuevo_texto_barra
            config_actual["texto_bienvenida"] = nuevo_texto_bienv
            config_actual["alineacion_titulo"] = nueva_alineacion
            config_actual["margen_superior"] = nuevo_margen
            guardar_configuracion(config_actual)
            st.success("✅ ¡Estilos actualizados!")
            st.rerun()

    st.sidebar.markdown("---")
    st.sidebar.subheader("🖼️ Imagen de Portada")
    foto_portada_subida = st.sidebar.file_uploader("Sube nueva portada:", type=["jpg", "jpeg", "png"], key="up_portada_admin")
    if foto_portada_subida is not None:
        if st.sidebar.button("💾 Guardar Nueva Portada"):
            try:
                ext = os.path.splitext(foto_portada_subida.name)[1]
                for f_ant in os.listdir(CARPETA_BANNER):
                    if f_ant.startswith("portada_fletes"):
                        try: os.remove(os.path.join(CARPETA_BANNER, f_ant))
                        except: pass
                ruta_portada = os.path.join(CARPETA_BANNER, f"portada_fletes{ext}")
                with open(ruta_portada, "wb") as f:
                    f.write(foto_portada_subida.getbuffer())
                st.sidebar.success("✅ ¡Portada guardada con éxito!")
                st.rerun()
            except Exception as e:
                st.sidebar.error(f"Error: {e}")

    ancho_portada_val = config_actual.get("ancho_portada", 100)
    ancho_portada_slider = st.sidebar.slider("Ancho de Portada (%):", min_value=30, max_value=100, value=int(ancho_portada_val), step=5)
    if st.sidebar.button("💾 Guardar Ancho de Portada"):
        config_actual["ancho_portada"] = ancho_portada_slider
        guardar_configuracion(config_actual)
        st.sidebar.success("✅ ¡Tamaño actualizado!")
        st.rerun()

    st.sidebar.markdown("---")
    st.sidebar.subheader("🛡️ Logotipo de la Barra")
    logo_subido = st.sidebar.file_uploader("Sube logotipo (png/jpg):", type=["jpg", "jpeg", "png"], key="up_logo_admin")
    if logo_subido is not None:
        if st.sidebar.button("💾 Guardar Logotipo"):
            try:
                ext = os.path.splitext(logo_subido.name)[1]
                for f_ant in os.listdir(CARPETA_LOGO):
                    try: os.remove(os.path.join(CARPETA_LOGO, f_ant))
                    except: pass
                ruta_logo = os.path.join(CARPETA_LOGO, f"logo_barra{ext}")
                with open(ruta_logo, "wb") as f:
                    f.write(logo_subido.getbuffer())
                st.sidebar.success("✅ ¡Logotipo guardado!")
                st.rerun()
            except Exception as e:
                st.sidebar.error(f"Error: {e}")

    st.sidebar.markdown("---")
    st.sidebar.subheader("👤 Fotos de Operadores")
    op_foto_sel = st.sidebar.selectbox("Selecciona Operador a Editar:", options=st.session_state.get("lista_operadores", []))
    foto_op_subida = st.sidebar.file_uploader(f"Foto para {op_foto_sel}:", type=["jpg", "jpeg", "png"], key="up_foto_op")
    
    if foto_op_subida is not None:
        if st.sidebar.button("💾 Guardar Foto de Operador"):
            try:
                ext = os.path.splitext(foto_op_subida.name)[1]
                nombre_limpio = "".join([c if c.isalnum() else "_" for c in op_foto_sel])
                for f_ant in os.listdir(CARPETA_FOTOS_OPERADORES):
                    if f_ant.startswith(nombre_limpio + "."):
                        try: os.remove(os.path.join(CARPETA_FOTOS_OPERADORES, f_ant))
                        except: pass
                ruta_foto_op = os.path.join(CARPETA_FOTOS_OPERADORES, f"{nombre_limpio}{ext}")
                with open(ruta_foto_op, "wb") as f:
                    f.write(foto_op_subida.getbuffer())
                st.sidebar.success(f"✅ ¡Foto guardada para {op_foto_sel}!")
                st.rerun()
            except Exception as e:
                st.sidebar.error(f"Error: {e}")

    tamano_actual_foto = config_actual.get("tamano_foto_operador", 120)
    nuevo_tamano_foto = st.sidebar.slider("Tamaño Foto Operador (px):", min_value=60, max_value=250, value=int(tamano_actual_foto), step=10)
    if st.sidebar.button("💾 Guardar Tamaño de Foto"):
        config_actual["tamano_foto_operador"] = nuevo_tamano_foto
        guardar_configuracion(config_actual)
        st.sidebar.success("✅ ¡Tamaño de foto actualizado!")
        st.rerun()

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
# RENDERIZADO DE CABECERA PERSONALIZABLE
# ==========================================
titulo_app_val = config_actual.get("titulo_app", "Control de Fletes AVE")
align_sel = config_actual.get("alineacion_titulo", "Izquierda")
margen_sup = config_actual.get("margen_superior", 0)

if align_sel == "Centro":
    align_style = "text-align: center;"
elif align_sel == "Derecha":
    align_style = "text-align: right;"
else:
    align_style = "text-align: left;"

st.markdown(f"""
    <div style="margin-top: {margen_sup}px; margin-bottom: 10px; {align_style}">
        <h1 style="display: inline-block; font-size: 2.2rem; font-weight: 700; color: #31333F;">🚛 {titulo_app_val}</h1>
    </div>
""", unsafe_allow_html=True)

archivos_banner = os.listdir(CARPETA_BANNER) if os.path.exists(CARPETA_BANNER) else []
portadas_candidatas = [f for f in archivos_banner if f.startswith("portada_fletes")]
ancho_portada_final = config_actual.get("ancho_portada", 100)

if portadas_candidatas:
    ruta_portada_activa = os.path.join(CARPETA_BANNER, portadas_candidatas[0])
    try:
        img_p = Image.open(ruta_portada_activa)
        if ancho_portada_final < 100:
            margen_vacio = (100 - ancho_portada_final) / 2
            _, col_img, _ = st.columns([margen_vacio, ancho_portada_final, margen_vacio])
            with col_img:
                st.image(img_p, use_container_width=True)
        else:
            st.image(img_p, use_container_width=True)
    except Exception:
        pass
else:
    try:
        st.image(r"C:\Users\CarlosArenas\Desktop\logo_fletes_ave.jpg", width=120)
    except Exception:
        try:
            st.image("logo_fletes_ave.jpg", width=120)
        except Exception:
            pass

html_logo_barra = ""
if os.path.exists(CARPETA_LOGO):
    archivos_l = os.listdir(CARPETA_LOGO)
    if archivos_l:
        ruta_l_activa = os.path.join(CARPETA_LOGO, archivos_l[0])
        try:
            with open(ruta_l_activa, "rb") as img_file:
                encoded_string = base64.b64encode(img_file.read()).decode()
                ext_l = archivos_l[0].split('.')[-1]
                html_logo_barra = f'<img src="data:image/{ext_l};base64,{encoded_string}" style="height: 26px; vertical-align: middle; margin-right: 10px; border-radius: 4px;">'
        except Exception:
            pass

texto_barra_val = config_actual.get("texto_barra", "PANEL DE NAVEGACIÓN FLETES AVE")
st.markdown(f"""
    <div style="background-color: rgba(20, 24, 33, 0.85); padding: 10px; border-radius: 6px; margin-top: 10px; margin-bottom: 20px; text-align: center; border: 1px solid rgba(255, 255, 255, 0.15);">
        <p style="color: #f0f2f6; margin: 0; font-size: 1rem; font-weight: 700; display: inline-block;">
            {html_logo_barra}{texto_barra_val}
        </p>
    </div>
""", unsafe_allow_html=True)

opciones_pestañas = [
    "🟢 Etapa 1: Salida Patio Base",
    "🟡 Etapa Intermedia: Pedrera (Llegada/Salida)",
    "🟠 Etapa 2: Llegada/Salida Cliente",
    "🔴 Etapa 3: Cierre de Ruta / Regreso"
]

if "etapa_idx" not in st.session_state:
    st.session_state.etapa_idx = 0

CLIENTES_DEFAULT = ["3185", "1198", "1555"]
DB_CLIENTES_DEFAULT = {
    "3185": {
        "HOLCIM MEXICO OPERACIONES (Escobedo)": {
            "contacto": "GUADALUPE",
            "telefono": "8112345678",
            "latitud": 25.847237,
            "longitud": -100.293652
        },
        "HOLCIM MEXICO OPERACIONES (Litos)": {
            "contacto": "ING. IVAN",
            "telefono": "8187654321",
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

if "asig_unidad_operador" not in st.session_state:
    st.session_state.asig_unidad_operador = dict(ASIG_DEFAULT)

if "lista_clientes" not in st.session_state:
    st.session_state.lista_clientes = list(CLIENTES_DEFAULT)

if "db_clientes" not in st.session_state:
    st.session_state.db_clientes = dict(DB_CLIENTES_DEFAULT)

if "catalogos_cargados" not in st.session_state:
    if os.path.exists(ARCHIVO_CATALOGOS):
        cargar_catalogos_de_disco()
    else:
        guardar_catalogos_en_disco()
    cargar_borrador_de_disco()
    st.session_state.catalogos_cargados = True

unidad_inicial_def = st.session_state.lista_unidades[0] if st.session_state.lista_unidades else "C-24"
operador_inicial_def = st.session_state.asig_unidad_operador.get(unidad_inicial_def, st.session_state.lista_operadores[0])

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
    ("unidad_activa", unidad_inicial_def),
    ("operador_activo", operador_inicial_def)
]:
    if var not in st.session_state:
        st.session_state[var] = val

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

    col_u1, col_u2, col_foto = st.columns([2, 2, 1.2])
    
    with col_u1:
        idx_u = st.session_state.lista_unidades.index(st.session_state.unidad_activa) if st.session_state.unidad_activa in st.session_state.lista_unidades else 0
        unidad_seleccionada = st.selectbox(
            "Selecciona la Unidad:", 
            options=st.session_state.lista_unidades, 
            index=idx_u, 
            key="sel_unidad_e1",
            on_change=al_cambiar_unidad
        )
    
    with col_u2:
        operador_sugerido = st.session_state.asig_unidad_operador.get(unidad_seleccionada, st.session_state.lista_operadores[0])
        st.session_state.operador_activo = operador_sugerido
        
        idx_op = st.session_state.lista_operadores.index(operador_sugerido) if operador_sugerido in st.session_state.lista_operadores else 0
        operador_seleccionado = st.selectbox(
            "Operador Asignado (Vínculo Automático):",
            options=st.session_state.lista_operadores,
            index=idx_op,
            key="sel_operador_e1"
        )
        st.session_state.operador_activo = operador_seleccionado

    with col_foto:
        st.markdown("<p style='font-size: 0.85rem; font-weight:600; margin-bottom:5px;'>Fotografía Operador:</p>", unsafe_allow_html=True)
        op_actual = st.session_state.get("operador_activo", "")
        nombre_limpio = "".join([c if c.isalnum() else "_" for c in op_actual])
        
        foto_encontrada = None
        if os.path.exists(CARPETA_FOTOS_OPERADORES):
            for archivo in os.listdir(CARPETA_FOTOS_OPERADORES):
                if archivo.startswith(nombre_limpio + "."):
                    foto_encontrada = os.path.join(CARPETA_FOTOS_OPERADORES, archivo)
                    break
        
        tamano_f = config_actual.get("tamano_foto_operador", 120)
        
        if foto_encontrada and os.path.exists(foto_encontrada):
            try:
                img_op = Image.open(foto_encontrada)
                st.image(img_op, width=tamano_f)
            except Exception:
                st.info("Sin foto")
        else:
            st.markdown(f"""
                <div style="width: {tamano_f}px; height: {tamano_f}px; background-color: #f0f2f6; border: 2px dashed #cccccc; border-radius: 8px; display: flex; align-items: center; justify-content: center; text-align: center; padding: 5px;">
                    <span style="font-size: 0.75rem; color: #666;">Sin foto cargada</span>
                </div>
            """, unsafe_allow_html=True)

    st.info(f"🚚 Unidad Activa: **{st.session_state.unidad_activa}** | 👤 Operador Vinculado: **{st.session_state.operador_activo}**")

    if es_admin:
        st.markdown("---")
        st.subheader("⚙️ Panel de Administración y Configuración de Catálogos")
        
        tab_admin1, tab_admin2, tab_admin3, tab_admin4 = st.tabs([
            "➕ 1. Alta Manual", 
            "🔗 2. Asignación Unidad-Operador", 
            "🗑️ 3. Eliminar Catálogos", 
            "📊 4. Reportes e Historial"
        ])
        
        with tab_admin1:
            with st.form("form_alta_manual"):
                st.markdown("#### Agregar elementos nuevos al sistema")
                nuevo_tipo = st.selectbox("¿Qué deseas agregar?", ["Unidad", "Operador", "Cliente / Dirección"], key="alta_tipo_v4")
                
                if nuevo_tipo == "Cliente / Dirección":
                    num_cli_input = st.text_input("Número de Cliente (ej. 3185):", key="alta_num_cli_v4")
                    nombre_dir_input = st.text_input("Nombre / Sucursal (ej. HOLCIM (Escobedo)):", key="alta_nom_dir_v4")
                    contacto_input = st.text_input("Contacto:", key="alta_contacto_v4")
                    tel_input = st.text_input("Teléfono:", key="alta_tel_v4")
                    lat_input = st.number_input("Latitud:", value=25.847237, format="%.6f", key="alta_lat_v4")
                    lon_input = st.number_input("Longitud (ej. -100.293652):", value=-100.293652, format="%.6f", key="alta_lon_v4")
                else:
                    nuevo_nombre = st.text_input("Escribe el nombre o código:", key="alta_nombre_v4")
                
                btn_submit_alta = st.form_submit_button("➕ Guardar en Catálogo")
                if btn_submit_alta:
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
                            cargar_catalogos_de_disco()
                            st.success(f"✅ ¡Dirección agregada al cliente {nc} con éxito!")
                            st.rerun()
                        else:
                            st.warning("⚠️ Completa el número de cliente y la sucursal.")
                    else:
                        if nuevo_nombre.strip():
                            val = nuevo_nombre.strip()
                            if nuevo_tipo == "Unidad" and val not in st.session_state.lista_unidades:
                                st.session_state.lista_unidades.append(val)
                            elif nuevo_tipo == "Operador" and val not in st.session_state.lista_operadores:
                                st.session_state.lista_operadores.append(val)
                            guardar_catalogos_en_disco()
                            cargar_catalogos_de_disco()
                            st.success(f"✅ ¡'{val}' agregado al catálogo con éxito!")
                            st.rerun()
                        else:
                            st.warning("⚠️ El campo de texto no puede estar vacío.")

        with tab_admin2:
            with st.form("form_asignacion"):
                st.markdown("#### Vincular Operador Fijo por Unidad")
                asig_uni_sel = st.selectbox("Selecciona Unidad:", st.session_state.lista_unidades, key="asig_uni_v4")
                op_actual_asig = st.session_state.asig_unidad_operador.get(asig_uni_sel, st.session_state.lista_operadores[0])
                idx_op_def = st.session_state.lista_operadores.index(op_actual_asig) if op_actual_asig in st.session_state.lista_operadores else 0
                asig_op_sel = st.selectbox("Operador Habitual:", st.session_state.lista_operadores, index=idx_op_def, key="asig_op_v4")
                
                btn_submit_asig = st.form_submit_button("🔗 Guardar Asignación Fija")
                if btn_submit_asig:
                    st.session_state.asig_unidad_operador[asig_uni_sel] = asig_op_sel
                    guardar_catalogos_en_disco()
                    cargar_catalogos_de_disco()
                    st.success(f"✅ ¡Asignación guardada! La unidad {asig_uni_sel} quedó vinculada a {asig_op_sel}.")
                    st.rerun()

        with tab_admin3:
            with st.form("form_eliminar_catalogo"):
                st.markdown("#### Eliminar elementos de los catálogos")
                tipo_del = st.selectbox("¿Qué deseas eliminar?", ["Unidad", "Operador", "Cliente / Sucursal"], key="del_tipo_v4")
                
                if tipo_del == "Unidad":
                    uni_del = st.selectbox("Selecciona Unidad:", st.session_state.lista_unidades, key="del_uni_v4")
                elif tipo_del == "Operador":
                    op_del = st.selectbox("Selecciona Operador:", st.session_state.lista_operadores, key="del_op_v4")
                else:
                    cli_del = st.selectbox("Selecciona Cliente:", st.session_state.lista_clientes, key="del_cli_v4")
                    sucursales_del_list = list(st.session_state.db_clientes.get(cli_del, {}).keys())
                    suc_del = st.selectbox("Selecciona Sucursal:", sucursales_del_list if sucursales_del_list else [""], key="del_suc_v4")
                
                btn_submit_del = st.form_submit_button("🗑️ Eliminar Elemento")
                if btn_submit_del:
                    if tipo_del == "Unidad":
                        if len(st.session_state.lista_unidades) > 1:
                            st.session_state.lista_unidades.remove(uni_del)
                            if uni_del in st.session_state.asig_unidad_operador:
                                del st.session_state.asig_unidad_operador[uni_del]
                            guardar_catalogos_en_disco()
                            cargar_catalogos_de_disco()
                            st.success(f"🗑️ Unidad {uni_del} eliminada.")
                            st.rerun()
                        else:
                            st.warning("⚠️ Debe haber al menos una unidad en el sistema.")
                    elif tipo_del == "Operador":
                        if len(st.session_state.lista_operadores) > 1:
                            st.session_state.lista_operadores.remove(op_del)
                            guardar_catalogos_en_disco()
                            cargar_catalogos_de_disco()
                            st.success(f"🗑️ Operador {op_del} eliminado.")
                            st.rerun()
                        else:
                            st.warning("⚠️ Debe haber al menos un operador en el sistema.")
                    else:
                        if cli_del in st.session_state.db_clientes and suc_del in st.session_state.db_clientes[cli_del]:
                            del st.session_state.db_clientes[cli_del][suc_del]
                            if not st.session_state.db_clientes[cli_del]:
                                del st.session_state.db_clientes[cli_del]
                                if cli_del in st.session_state.lista_clientes:
                                    st.session_state.lista_clientes.remove(cli_del)
                            guardar_catalogos_en_disco()
                            cargar_catalogos_de_disco()
                            st.success(f"🗑️ Sucursal {suc_del} eliminada.")
                            st.rerun()

        with tab_admin4:
            st.markdown("#### Historial General de Rutas")
            if os.path.exists(ARCHIVO_REGISTROS):
                try:
                    df_reg = pd.read_excel(ARCHIVO_REGISTROS)
                    st.dataframe(df_reg, use_container_width=True)
                    with open(ARCHIVO_REGISTROS, "rb") as f_excel:
                        st.download_button("📥 Descargar Reporte Completo (Excel)", f_excel, file_name="registros_rutas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.info("ℹ️ Aún no hay registros de rutas completadas.")
            else:
                st.info("ℹ️ Aún no hay registros de rutas completadas.")

    st.markdown("---")
    st.subheader("Kilometraje y Hora de Salida")
    
    col_k1, col_h1 = st.columns(2)
    with col_k1:
        st.session_state.km_salida_patio = st.number_input("Kilometraje Salida Patio:", value=float(st.session_state.get("km_salida_patio", 145826.0)), format="%.1f")
    with col_h1:
        if not st.session_state.hora_salida_patio:
            st.session_state.hora_salida_patio = obtener_hora_mexico()
        st.text_input("Hora Salida Patio:", value=st.session_state.hora_salida_patio, disabled=True)

    if st.button("🚀 Registrar Salida y Continuar a Pedrera", type="primary"):
        guardar_borrador_de_disco()
        st.session_state.etapa_idx = 1
        st.success("✅ ¡Salida registrada con éxito!")
        st.rerun()

# ==========================================
# ETAPA INTERMEDIA: PEDRERA (LLEGADA/SALIDA)
# ==========================================
elif st.session_state.etapa_idx == 1:
    st.header("Etapa Intermedia: Pedrera (Llegada / Salida)")
    st.info(f"🚚 Unidad Activa: **{st.session_state.unidad_activa}** | 👤 Operador: **{st.session_state.operador_activo}**")

    st.markdown("---")
    st.subheader("1️⃣ Llegada a Pedrera")
    col_kp1, col_hp1 = st.columns(2)
    with col_kp1:
        st.session_state.km_llegada_pedrera = st.number_input("Km Llegada Pedrera:", value=float(st.session_state.get("km_llegada_pedrera", 145826.0)), format="%.1f")
    with col_hp1:
        if not st.session_state.hora_llegada_pedrera:
            st.session_state.hora_llegada_pedrera = obtener_hora_mexico()
        st.text_input("Hora Llegada Pedrera:", value=st.session_state.hora_llegada_pedrera, disabled=True)

    if st.button("📍 Confirmar Llegada a Pedrera"):
        guardar_borrador_de_disco()
        st.success("✅ Llegada a pedrera guardada.")

    st.markdown("---")
    st.subheader("2️⃣ Salida de Pedrera (Cargado)")
    col_kp2, col_hp2 = st.columns(2)
    with col_kp2:
        st.session_state.km_salida_pedrera = st.number_input("Km Salida Pedrera:", value=float(st.session_state.get("km_salida_pedrera", 145826.0)), format="%.1f")
    with col_hp2:
        if not st.session_state.hora_salida_pedrera:
            st.session_state.hora_salida_pedrera = obtener_hora_mexico()
        st.text_input("Hora Salida Pedrera:", value=st.session_state.hora_salida_pedrera, disabled=True)

    col_btn_ant, col_btn_sig = st.columns(2)
    with col_btn_ant:
        if st.button("⬅️ Volver a Etapa 1"):
            st.session_state.etapa_idx = 0
            st.rerun()
    with col_btn_sig:
        if st.button("➡️ Avanzar a Cliente", type="primary"):
            guardar_borrador_de_disco()
            st.session_state.etapa_idx = 2
            st.rerun()

# ==========================================
# ETAPA 2: LLEGADA/SALIDA CLIENTE (CON MAPA Y ACCIONES DE LLAMADA/WHATSAPP)
# ==========================================
elif st.session_state.etapa_idx == 2:
    st.header("Etapa 2: Llegada y Salida con el Cliente")
    st.info(f"🚚 Unidad Activa: **{st.session_state.unidad_activa}** | 👤 Operador: **{st.session_state.operador_activo}**")

    st.markdown("---")
    st.subheader("🏢 Selección de Cliente y Destino")

    col_c1, col_c2 = st.columns(2)
    with col_c1:
        cliente_sel = st.selectbox("Número de Cliente:", options=st.session_state.lista_clientes, key="sel_cliente_e2")
    with col_c2:
        sucursales_dict = st.session_state.db_clientes.get(cliente_sel, {})
        nombres_sucursales = list(sucursales_dict.keys())
        sucursal_sel = st.selectbox("Sucursal / Destino:", options=nombres_sucursales if nombres_sucursales else ["Principal"], key="sel_sucursal_e2")

    if nombres_sucursales and sucursal_sel in sucursales_dict:
        info_s = sucursales_dict[sucursal_sel]
        st.session_state.cli_contacto = info_s.get("contacto", "")
        st.session_state.cli_telefono = info_s.get("telefono", "")
        st.session_state.cli_latitud = info_s.get("latitud", 25.844412)
        st.session_state.cli_longitud = info_s.get("longitud", -100.293652)

    # Tarjeta de información con botones de Llamada y WhatsApp
    tel_limpio = "".join([c for c in str(st.session_state.cli_telefono) if c.isdigit()])
    link_whatsapp = f"https://wa.me/52{tel_limpio}" if tel_limpio else "#"
    link_llamada = f"tel:{st.session_state.cli_telefono}" if st.session_state.cli_telefono else "#"

    st.markdown(f"""
        <div style="background-color: #f8f9fa; padding: 12px; border-radius: 6px; border-left: 4px solid #ff4b4b; margin-bottom: 15px;">
            <p style="margin: 0; font-size: 0.95rem;"><b>👤 Contacto:</b> {st.session_state.cli_contacto or 'N/D'} | <b>📞 Teléfono:</b> {st.session_state.cli_telefono or 'N/D'}</p>
            <p style="margin: 5px 0 10px 0; font-size: 0.85rem; color: #555;">📍 Ubicación: Lat: {st.session_state.cli_latitud}, Lon: {st.session_state.cli_longitud}</p>
            <div>
                <a href="{link_llamada}" target="_blank" style="background-color: #28a745; color: white; padding: 6px 12px; border-radius: 4px; text-decoration: none; font-size: 0.85rem; font-weight: bold; margin-right: 10px;">📞 Llamar al Cliente</a>
                <a href="{link_whatsapp}" target="_blank" style="background-color: #25D366; color: white; padding: 6px 12px; border-radius: 4px; text-decoration: none; font-size: 0.85rem; font-weight: bold;">💬 Enviar WhatsApp</a>
            </div>
        </div>
    """, unsafe_allow_html=True)

    # Mostrar Mapa interactivo con la ubicación según Latitud y Longitud
    st.markdown("#### 🗺️ Mapa de Ubicación del Destino")
    df_mapa = pd.DataFrame({
        'lat': [st.session_state.cli_latitud],
        'lon': [st.session_state.cli_longitud]
    })
    st.map(df_mapa, zoom=14)

    st.markdown("---")
    st.subheader("1️⃣ Llegada al Cliente (Descarga)")
    col_kc1, col_hc1 = st.columns(2)
    with col_kc1:
        st.session_state.km_llegada_cliente = st.number_input("Km Llegada Cliente:", value=float(st.session_state.get("km_llegada_cliente", 145826.0)), format="%.1f")
    with col_hc1:
        if not st.session_state.hora_llegada_cliente:
            st.session_state.hora_llegada_cliente = obtener_hora_mexico()
        st.text_input("Hora Llegada Cliente:", value=st.session_state.hora_llegada_cliente, disabled=True)

    if st.button("📍 Confirmar Llegada al Cliente"):
        guardar_borrador_de_disco()
        st.success("✅ Llegada registrada.")

    st.markdown("---")
    st.subheader("2️⃣ Salida del Cliente (Desocupado)")
    col_kc2, col_hc2 = st.columns(2)
    with col_kc2:
        st.session_state.km_salida_cliente = st.number_input("Km Salida Cliente:", value=float(st.session_state.get("km_salida_cliente", 145826.0)), format="%.1f")
    with col_hc2:
        if not st.session_state.hora_salida_cliente:
            st.session_state.hora_salida_cliente = obtener_hora_mexico()
        st.text_input("Hora Salida Cliente:", value=st.session_state.hora_salida_cliente, disabled=True)

    col_btn_ant2, col_btn_sig2 = st.columns(2)
    with col_btn_ant2:
        if st.button("⬅️ Volver a Pedrera"):
            st.session_state.etapa_idx = 1
            st.rerun()
    with col_btn_sig2:
        if st.button("➡️ Ir al Cierre de Ruta", type="primary"):
            guardar_borrador_de_disco()
            st.session_state.etapa_idx = 3
            st.rerun()

# ==========================================
# ETAPA 3: CIERRE DE RUTA / REGRESO
# ==========================================
elif st.session_state.etapa_idx == 3:
    st.header("Etapa 3: Cierre de Ruta / Regreso a Base")
    st.info(f"🚚 Unidad Activa: **{st.session_state.unidad_activa}** | 👤 Operador: **{st.session_state.operador_activo}**")

    st.markdown("---")
    col_kf, col_hf = st.columns(2)
    with col_kf:
        st.session_state.km_final_viaje = st.number_input("Km Final Regreso a Patio:", value=float(st.session_state.get("km_final_viaje", 145826.0)), format="%.1f")
    with col_hf:
        if not st.session_state.hora_cierre_viaje:
            st.session_state.hora_cierre_viaje = obtener_hora_mexico()
        st.text_input("Hora Cierre:", value=st.session_state.hora_cierre_viaje, disabled=True)

    km_totales = st.session_state.km_final_viaje - st.session_state.km_salida_patio
    st.metric(label="📊 Kilómetros Totales Recorridos", value=f"{km_totales:.1f} km")

    col_b_ant, col_b_fin = st.columns(2)
    with col_b_ant:
        if st.button("⬅️ Volver a Cliente"):
            st.session_state.etapa_idx = 2
            st.rerun()
    with col_b_fin:
        if not st.session_state.ruta_guardada:
            if st.button("💾 Guardar y Concluir Ruta Definitivamente", type="primary"):
                registro_final = {
                    "Fecha Registro": obtener_hora_mexico(),
                    "Unidad": st.session_state.unidad_activa,
                    "Operador": st.session_state.operador_activo,
                    "Km Salida Patio": st.session_state.km_salida_patio,
                    "Hora Salida Patio": st.session_state.hora_salida_patio,
                    "Km Llegada Pedrera": st.session_state.km_llegada_pedrera,
                    "Hora Llegada Pedrera": st.session_state.hora_llegada_pedrera,
                    "Km Salida Pedrera": st.session_state.km_salida_pedrera,
                    "Hora Salida Pedrera": st.session_state.hora_salida_pedrera,
                    "Km Llegada Cliente": st.session_state.km_llegada_cliente,
                    "Hora Llegada Cliente": st.session_state.hora_llegada_cliente,
                    "Km Salida Cliente": st.session_state.km_salida_cliente,
                    "Hora Salida Cliente": st.session_state.hora_salida_cliente,
                    "Km Final Viaje": st.session_state.km_final_viaje,
                    "Hora Cierre Viaje": st.session_state.hora_cierre_viaje,
                    "Km Totales": km_totales
                }
                guardar_en_excel(registro_final, ARCHIVO_REGISTROS)
                st.session_state.ruta_guardada = True
                if os.path.exists(ARCHIVO_BORRADOR):
                    try: os.remove(ARCHIVO_BORRADOR)
                    except: pass
                st.success("🎉 ¡Ruta guardada exitosamente en el sistema!")
        else:
            st.info("ℹ️ Esta ruta ya fue guardada previamente.")

    if st.button("🔄 Iniciar Nuevo Viaje (Reiniciar)"):
        for k in list(st.session_state.keys()):
            if k not in ["lista_unidades", "lista_operadores", "asig_unidad_operador", "lista_clientes", "db_clientes", "catalogos_cargados"]:
                del st.session_state[k]
        if os.path.exists(ARCHIVO_BORRADOR):
            try: os.remove(ARCHIVO_BORRADOR)
            except: pass
        st.success("🔄 Formulario reiniciado.")
        st.rerun()