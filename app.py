import os
import json
import base64
import re
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
        "texto_bienvenida": "✨ ¡Bienvenido al sistema operativo de rutas!",
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
    tz_mexico = timezone(timedelta(hours=-6))
    return datetime.now(tz_mexico).strftime("%Y-%m-%d %H:%M:%S")

def asegurar_longitud_negativa(val):
    try:
        fval = float(val)
        return -abs(fval)
    except Exception:
        return -100.293652

def extraer_lat_lon_de_texto(texto):
    if not texto:
        return 25.847237, -100.293652
    texto_str = str(texto).strip()
    match_at = re.search(r'@(-?\d+\.\d+),(-?\d+\.\d+)', texto_str)
    if match_at:
        try:
            return float(match_at.group(1)), asegurar_longitud_negativa(float(match_at.group(2)))
        except:
            pass
    match_q = re.search(r'(?:q=|ll=)(-?\d+\.\d+)[,\s]+(-?\d+\.\d+)', texto_str)
    if match_q:
        try:
            return float(match_q.group(1)), asegurar_longitud_negativa(float(match_q.group(2)))
        except:
            pass
    match_coords = re.findall(r'(-?\d{2}\.\d+)[,\s]+(-?\d{2,3}\.\d+)', texto_str)
    if match_coords:
        try:
            return float(match_coords[0][0]), asegurar_longitud_negativa(float(match_coords[0][1]))
        except:
            pass
    return 25.847237, -100.293652

def guardar_en_excel(data_dict, archivo_excel=ARCHIVO_REGISTROS):
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

UNIDADES_DEFAULT = ["C-24", "C-25", "C-26", "C-27", "C-28", "T-01", "T-02", "C-13"]
OPERADORES_DEFAULT = [
    "Octavio Rodrigo Serrano Saavedra", "Juan Carlos Arenas", "Pedro Martinez",
    "Jose Luis Rodriguez", "Luis Gonzalez Mendez", "Brandon Heriberto Medina Marroquin",
    "Oscar Garza Salas", "Cesar Dagoberto Galvan", "Mario Alberto Trejo"
]
ASIG_DEFAULT = {
    "C-24": "Octavio Rodrigo Serrano Saavedra", "C-25": "Juan Carlos Arenas",
    "C-26": "Luis Gonzalez Mendez", "C-27": "Juan Carlos Arenas",
    "C-28": "Brandon Heriberto Medina Marroquin", "T-01": "Pedro Martinez",
    "T-02": "Jose Luis Rodriguez", "C-13": "Oscar Garza Salas"
}

def guardar_borrador_de_disco():
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
                        "num_cliente": num_cli, "nombre": suc_nombre,
                        "contacto": info.get("contacto", ""), "telefono": info.get("telefono", ""),
                        "latitud": info.get("latitud", 25.844412), "longitud": info.get("longitud", -100.293652)
                    })
            pd.DataFrame(clientes_data).to_excel(writer, sheet_name="Clientes", index=False)
    except Exception as e:
        print(f"Error al guardar catálogos: {e}")

def cargar_catalogos_de_disco():
    if os.path.exists(ARCHIVO_CATALOGOS):
        try:
            xl = pd.ExcelFile(ARCHIVO_CATALOGOS)
            if "Unidades" in xl.sheet_names:
                df_u = pd.read_excel(ARCHIVO_CATALOGOS, sheet_name="Unidades")
                col_u = "Unidades" if "Unidades" in df_u.columns else df_u.columns[0]
                unidades_disco = df_u[col_u].dropna().astype(str).tolist()
                if unidades_disco: st.session_state.lista_unidades = unidades_disco
            if "Operadores" in xl.sheet_names:
                df_o = pd.read_excel(ARCHIVO_CATALOGOS, sheet_name="Operadores")
                col_o = "Operadores" if "Operadores" in df_o.columns else df_o.columns[0]
                operadores_disco = df_o[col_o].dropna().astype(str).tolist()
                if operadores_disco: st.session_state.lista_operadores = operadores_disco
            if "Asignacion" in xl.sheet_names:
                df_a = pd.read_excel(ARCHIVO_CATALOGOS, sheet_name="Asignacion")
                if not df_a.empty and "Unidad" in df_a.columns and "Operador" in df_a.columns:
                    st.session_state.asig_unidad_operador = {}
                    for _, row in df_a.iterrows():
                        st.session_state.asig_unidad_operador[str(row["Unidad"]).strip()] = str(row["Operador"]).strip()
            if "Clientes" in xl.sheet_names:
                df_c = pd.read_excel(ARCHIVO_CATALOGOS, sheet_name="Clientes")
                if not df_c.empty:
                    nueva_db = {}
                    lista_c = []
                    for _, row in df_c.iterrows():
                        num_cli = str(row.get("num_cliente", row.get("ID_Cliente", row.iloc[0] if len(row)>0 else "S/N"))).strip().replace(".0", "")
                        suc_nombre = str(row.get("nombre", row.get("Nombre_Cliente", row.iloc[1] if len(row)>1 else "Principal")))
                        if num_cli and num_cli.lower() != 'nan':
                            if num_cli not in lista_c: lista_c.append(num_cli)
                            if num_cli not in nueva_db: nueva_db[num_cli] = {}
                            try: lat_val = float(row.get("latitud", 25.844412))
                            except: lat_val = 25.844412
                            try: lon_val = float(row.get("longitud", -100.395043))
                            except: lon_val = -100.395043
                            
                            contacto_val = str(row.get("contacto", ""))
                            if contacto_val.lower() == 'nan': contacto_val = ""
                            telefono_val = str(row.get("telefono", ""))
                            if telefono_val.lower() == 'nan': telefono_val = ""

                            nueva_db[num_cli][suc_nombre] = {
                                "contacto": contacto_val,
                                "telefono": telefono_val,
                                "latitud": lat_val,
                                "longitud": asegurar_longitud_negativa(lon_val)
                            }
                    if lista_c:
                        st.session_state.lista_clientes = lista_c
                        st.session_state.db_clientes = nueva_db
        except Exception as e:
            print(f"Error al leer catálogos: {e}")

# ==========================================
# CALLBACKS DE SINCRONIZACIÓN REACTIVA
# ==========================================
def al_cambiar_unidad():
    uni = st.session_state.get("sel_unidad_e1")
    if uni:
        st.session_state.unidad_activa = uni
        op_asignado = st.session_state.asig_unidad_operador.get(uni, st.session_state.lista_operadores[0])
        st.session_state.operador_activo = op_asignado
    guardar_borrador_de_disco()

def actualizar_datos_cliente():
    cli = str(st.session_state.get("sel_cliente_e2", "")).strip().replace(".0", "")
    suc = st.session_state.get("sel_sucursal_e2", "")
    
    if cli and cli in st.session_state.db_clientes:
        sucursales = list(st.session_state.db_clientes[cli].keys())
        sucursales = [s for s in sucursales if str(s).lower() != 'nan' and str(s).strip() != '']
        if not sucursales:
            sucursales = ["Principal"]
            
        if suc not in sucursales:
            st.session_state.sel_sucursal_e2 = sucursales[0]
            suc = sucursales[0]
        
        info_sucursal = st.session_state.db_clientes[cli].get(suc, {})
        st.session_state.cli_contacto = info_sucursal.get("contacto", "")
        st.session_state.cli_telefono = info_sucursal.get("telefono", "")
        st.session_state.cli_latitud = float(info_sucursal.get("latitud", 25.844412))
        st.session_state.cli_longitud = float(info_sucursal.get("longitud", -100.395043))

def al_cambiar_cliente_e2():
    cli = str(st.session_state.get("sel_cliente_e2", "")).strip().replace(".0", "")
    if cli and cli in st.session_state.db_clientes:
        sucursales = list(st.session_state.db_clientes[cli].keys())
        sucursales = [s for s in sucursales if str(s).lower() != 'nan' and str(s).strip() != '']
        if sucursales:
            st.session_state.sel_sucursal_e2 = sucursales[0]
    actualizar_datos_cliente()
    st.rerun()

def al_cambiar_sucursal_e2():
    actualizar_datos_cliente()

# ==========================================
# CONFIGURACIÓN INICIAL DE LA APP
# ==========================================
st.set_page_config(page_title="Control de Fletes AVE", layout="wide")
PIN_ADMIN = "1234"

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
        nueva_alineacion = st.selectbox("Alineación del Título:", options=alineaciones, index=alineaciones.index(align_actual) if align_actual in alineaciones else 0)
        nuevo_margen = st.slider("Margen Superior (px):", min_value=0, max_value=50, value=int(config_actual.get("margen_superior", 0)), step=5)
        if st.form_submit_button("💾 Guardar Cambios de Texto"):
            config_actual["titulo_app"] = nuevo_titulo
            config_actual["texto_barra"] = nuevo_texto_barra
            config_actual["texto_bienvenida"] = nuevo_texto_bienv
            config_actual["alineacion_titulo"] = nueva_alineacion
            config_actual["margen_superior"] = nuevo_margen
            guardar_configuracion(config_actual)
            st.success("✅ ¡Estilos actualizados!")
            st.rerun()

    st.sidebar.markdown("---")
    st.sidebar.subheader("🖼 Imagen de Portada")
    foto_portada_subida = st.sidebar.file_uploader("Sube nueva portada:", type=["jpg", "jpeg", "png"], key="up_portada_admin")
    if foto_portada_subida is not None and st.sidebar.button("💾 Guardar Nueva Portada"):
        try:
            ext = os.path.splitext(foto_portada_subida.name)[1]
            for f_ant in os.listdir(CARPETA_BANNER):
                if f_ant.startswith("portada_fletes"):
                    try: os.remove(os.path.join(CARPETA_BANNER, f_ant))
                    except: pass
            with open(os.path.join(CARPETA_BANNER, f"portada_fletes{ext}"), "wb") as f:
                f.write(foto_portada_subida.getbuffer())
            st.sidebar.success("✅ ¡Portada guardada!")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"Error: {e}")

    ancho_portada_slider = st.sidebar.slider("Ancho de Portada (%):", min_value=30, max_value=100, value=int(config_actual.get("ancho_portada", 100)), step=5)
    if st.sidebar.button("💾 Guardar Ancho de Portada"):
        config_actual["ancho_portada"] = ancho_portada_slider
        guardar_configuracion(config_actual)
        st.sidebar.success("✅ ¡Tamaño actualizado!")
        st.rerun()

    st.sidebar.markdown("---")
    st.sidebar.subheader("🛡️ Logotipo de la Barra")
    logo_subido = st.sidebar.file_uploader("Sube logotipo:", type=["jpg", "jpeg", "png"], key="up_logo_admin")
    if logo_subido is not None and st.sidebar.button("💾 Guardar Logotipo"):
        try:
            ext = os.path.splitext(logo_subido.name)[1]
            for f_ant in os.listdir(CARPETA_LOGO):
                try: os.remove(os.path.join(CARPETA_LOGO, f_ant))
                except: pass
            with open(os.path.join(CARPETA_LOGO, f"logo_barra{ext}"), "wb") as f:
                f.write(logo_subido.getbuffer())
            st.sidebar.success("✅ ¡Logotipo guardado!")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"Error: {e}")

    st.sidebar.markdown("---")
    st.sidebar.subheader("👤 Fotos de Operadores")
    op_foto_sel = st.sidebar.selectbox("Selecciona Operador a Editar:", options=st.session_state.get("lista_operadores", []))
    foto_op_subida = st.sidebar.file_uploader(f"Foto para {op_foto_sel}:", type=["jpg", "jpeg", "png"], key="up_foto_op")
    if foto_op_subida is not None and st.sidebar.button("💾 Guardar Foto de Operador"):
        try:
            ext = os.path.splitext(foto_op_subida.name)[1]
            nombre_limpio = "".join([c if c.isalnum() else "_" for c in op_foto_sel])
            for f_ant in os.listdir(CARPETA_FOTOS_OPERADORES):
                if f_ant.startswith(nombre_limpio + "."):
                    try: os.remove(os.path.join(CARPETA_FOTOS_OPERADORES, f_ant))
                    except: pass
            with open(os.path.join(CARPETA_FOTOS_OPERADORES, f"{nombre_limpio}{ext}"), "wb") as f:
                f.write(foto_op_subida.getbuffer())
            st.sidebar.success(f"✅ ¡Foto guardada para {op_foto_sel}!")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"Error: {e}")

    nuevo_tamano_foto = st.sidebar.slider("Tamaño Foto Operador (px):", min_value=60, max_value=250, value=int(config_actual.get("tamano_foto_operador", 120)), step=10)
    if st.sidebar.button("💾 Guardar Tamaño de Foto"):
        config_actual["tamano_foto_operador"] = nuevo_tamano_foto
        guardar_configuracion(config_actual)
        st.sidebar.success("✅ ¡Actualizado!")
        st.rerun()
else:
    if pin_input != "":
        st.sidebar.error("❌ PIN Incorrecto")
    else:
        st.sidebar.info("Modo Usuario (Sólo lectura)")

st.sidebar.markdown("---")
st.sidebar.subheader("⚠️ Reporte de Incidencias")
with st.sidebar.expander("Registrar Incidencia / Parada"):
    tipo_incidencia = st.selectbox("Tipo de Incidencia:", ["Accidente", "Retraso", "Falla mecánica", "Ponchadura", "Hora de comida", "Otro"])
    comentario_incidencia = st.text_area("Detalles / Comentarios:")
    if st.button("🚨 Guardar Incidencia"):
        data_inc = {
            "Fecha/Hora": obtener_hora_mexico(),
            "Unidad": st.session_state.get("unidad_activa", "N/D"),
            "Operador": st.session_state.get("operador_activo", "N/D"),
            "Tipo Incidencia": tipo_incidencia,
            "Detalles": comentario_incidencia
        }
        guardar_incidencia_en_excel(data_inc)
        st.sidebar.success("✅ Incidencia registrada correctamente.")

if es_admin and os.path.exists(ARCHIVO_INCIDENCIAS):
    with open(ARCHIVO_INCIDENCIAS, "rb") as file:
        st.sidebar.download_button("📥 Descargar Excel de Incidencias", data=file, file_name="registros_incidencias.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ==========================================
# ENCABEZADO DE LA APLICACIÓN
# ==========================================
titulo_app_val = config_actual.get("titulo_app", "Control de Fletes AVE")
align_sel = config_actual.get("alineacion_titulo", "Izquierda")
margen_sup = config_actual.get("margen_superior", 0)

align_style = "text-align: center;" if align_sel == "Centro" else ("text-align: right;" if align_sel == "Derecha" else "text-align: left;")
st.markdown(f"""
    <div style="margin-top: {margen_sup}px; margin-bottom: 10px; {align_style}">
        <h1 style="display: inline-block; font-size: 2.2rem; font-weight: 700; color: #31333F;">🚛 {titulo_app_val}</h1>
    </div>
""", unsafe_allow_html=True)

archivos_banner = os.listdir(CARPETA_BANNER) if os.path.exists(CARPETA_BANNER) else []
portadas_candidatas = [f for f in archivos_banner if f.startswith("portada_fletes")]
ancho_portada_final = config_actual.get("ancho_portada", 100)

if portadas_candidatas:
    try:
        img_p = Image.open(os.path.join(CARPETA_BANNER, portadas_candidatas[0]))
        if ancho_portada_final < 100:
            margen_vacio = (100 - ancho_portada_final) / 2
            _, col_img, _ = st.columns([margen_vacio, ancho_portada_final, margen_vacio])
            with col_img: st.image(img_p, use_container_width=True)
        else:
            st.image(img_p, use_container_width=True)
    except Exception:
        pass

html_logo_barra = ""
if os.path.exists(CARPETA_LOGO):
    archivos_l = os.listdir(CARPETA_LOGO)
    if archivos_l:
        try:
            with open(os.path.join(CARPETA_LOGO, archivos_l[0]), "rb") as img_file:
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

options_pestañas = [
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
        "HOLCIM MEXICO OPERACIONES (Escobedo)": {"contacto": "GUADALUPE", "telefono": "8112345678", "latitud": 25.847237, "longitud": -100.293652},
        "HOLCIM MEXICO OPERACIONES (Litos)": {"contacto": "ING. IVAN", "telefono": "8187654321", "latitud": 25.898032, "longitud": -100.200360}
    },
    "1198": {
        "ORGANIZACIÓN Y SERVICIO PARA LA CONSTRUCCION, S.A. DE C.V.": {"contacto": "", "telefono": "2381915291", "latitud": 25.829361, "longitud": -100.230889}
    },
    "1555": {
        "GRUPO PERFIMEXA SA DE CV": {"contacto": "", "telefono": "8115783462", "latitud": 26.004936, "longitud": -100.296356}
    }
}

if "lista_unidades" not in st.session_state: st.session_state.lista_unidades = list(UNIDADES_DEFAULT)
if "lista_operadores" not in st.session_state: st.session_state.lista_operadores = list(OPERADORES_DEFAULT)
if "asig_unidad_operador" not in st.session_state: st.session_state.asig_unidad_operador = dict(ASIG_DEFAULT)
if "lista_clientes" not in st.session_state: st.session_state.lista_clientes = list(CLIENTES_DEFAULT)
if "db_clientes" not in st.session_state: st.session_state.db_clientes = dict(DB_CLIENTES_DEFAULT)

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

if "sel_cliente_e2" not in st.session_state or st.session_state.sel_cliente_e2 not in st.session_state.lista_clientes:
    st.session_state.sel_cliente_e2 = st.session_state.lista_clientes[0]

cli_inicial_limpio = str(st.session_state.sel_cliente_e2).strip().replace(".0", "")
sucursales_iniciales = list(st.session_state.db_clientes.get(cli_inicial_limpio, {"Principal": {}}).keys())
sucursales_iniciales = [s for s in sucursales_iniciales if str(s).lower() != 'nan' and str(s).strip() != '']
if not sucursales_iniciales:
    sucursales_iniciales = ["Principal"]

if "sel_sucursal_e2" not in st.session_state or st.session_state.sel_sucursal_e2 not in sucursales_iniciales:
    st.session_state.sel_sucursal_e2 = sucursales_iniciales[0]

actualizar_datos_cliente()

pestana_activa = st.radio("Selecciona la Etapa:", options_pestañas, index=st.session_state.etapa_idx, horizontal=True)
st.session_state.etapa_idx = options_pestañas.index(pestana_activa)
st.markdown("---")

# ==========================================
# PANEL DE ADMINISTRACIÓN
# ==========================================
if es_admin:
    st.markdown("### ⚙️ Panel de Administración y Configuración de Catálogos")
    tab_admin1, tab_admin2, tab_admin3, tab_admin4 = st.tabs([
        "➕ 1. Alta y Edición", "🔗 2. Asignación Unidad-Operador", "🗑️ 3. Eliminar Catálogos", "📊 4. Reportes y Edición de Registros"
    ])
    
    with tab_admin1:
        sub_modo = st.radio("Método de Ingreso:", ["Registro Manual Individual", "📂 Subida Masiva (Excel / CSV)"], horizontal=True)
        tipo_gestion = st.selectbox("¿Qué deseas gestionar?", ["Unidad", "Operador", "Cliente y Sucursal"])
        
        if sub_modo == "Registro Manual Individual":
            if tipo_gestion == "Unidad":
                nueva_uni = st.text_input("Código de la Unidad:")
                if st.button("💾 Guardar Unidad") and nueva_uni.strip():
                    if nueva_uni.strip() not in st.session_state.lista_unidades:
                        st.session_state.lista_unidades.append(nueva_uni.strip())
                        guardar_catalogos_en_disco()
                        st.success("✅ Unidad agregada.")
                        st.rerun()
            elif tipo_gestion == "Operador":
                nuevo_op = st.text_input("Nombre del Operador:")
                if st.button("💾 Guardar Operador") and nuevo_op.strip():
                    if nuevo_op.strip() not in st.session_state.lista_operadores:
                        st.session_state.lista_operadores.append(nuevo_op.strip())
                        guardar_catalogos_en_disco()
                        st.success("✅ Operador agregado.")
                        st.rerun()
            elif tipo_gestion == "Cliente y Sucursal":
                with st.form("form_alta_cliente"):
                    num_cli = st.text_input("Número de Cliente (ID):").strip().replace(".0", "")
                    nom_suc = st.text_input("Nombre de la Sucursal / Razón Social:")
                    contacto = st.text_input("Persona de Contacto:")
                    telefono = st.text_input("Teléfono:")
                    map_input = st.text_input("Coordenadas GPS o Link de Google Maps:")
                    if st.form_submit_button("💾 Guardar Cliente / Sucursal") and num_cli and nom_suc.strip():
                        lat, lon = extraer_lat_lon_de_texto(map_input)
                        if num_cli not in st.session_state.db_clientes: st.session_state.db_clientes[num_cli] = {}
                        if num_cli not in st.session_state.lista_clientes: st.session_state.lista_clientes.append(num_cli)
                        st.session_state.db_clientes[num_cli][nom_suc.strip()] = {
                            "contacto": contacto.strip(), "telefono": telefono.strip(), "latitud": lat, "longitud": lon
                        }
                        guardar_catalogos_en_disco()
                        st.success("✅ Cliente/Sucursal guardado con éxito.")
                        st.rerun()
        else:
            archivo_subido_masivo = st.file_uploader(f"Selecciona archivo para {tipo_gestion}(s):", type=["xlsx", "xls", "csv"])
            if archivo_subido_masivo is not None:
                try:
                    df_masivo = pd.read_csv(archivo_subido_masivo) if archivo_subido_masivo.name.endswith(".csv") else pd.read_excel(archivo_subido_masivo)
                    st.dataframe(df_masivo.head())
                    if st.button(f"🚀 Procesar e Importar"):
                        contador_exito = 0
                        for _, row in df_masivo.iterrows():
                            if tipo_gestion == "Unidad":
                                u_limpia = str(row.iloc[0]).strip()
                                if u_limpia and u_limpia not in st.session_state.lista_unidades:
                                    st.session_state.lista_unidades.append(u_limpia)
                                    contador_exito += 1
                            elif tipo_gestion == "Operador":
                                op_limpio = str(row.iloc[0]).strip()
                                if op_limpio and op_limpio not in st.session_state.lista_operadores:
                                    st.session_state.lista_operadores.append(op_limpio)
                                    contador_exito += 1
                        guardar_catalogos_en_disco()
                        st.success(f"✅ Se importaron {contador_exito} registros correctamente.")
                        st.rerun()
                except Exception as e:
                    st.error(f"Error al procesar archivo masivo: {e}")

    with tab_admin2:
        st.subheader("🔗 Asignación Predeterminada Unidad ➡️ Operador")
        with st.form("form_asignacion_uo"):
            nuevas_asig = {}
            for u in st.session_state.lista_unidades:
                op_actual = st.session_state.asig_unidad_operador.get(u, st.session_state.lista_operadores[0])
                idx_op = st.session_state.lista_operadores.index(op_actual) if op_actual in st.session_state.lista_operadores else 0
                sel_op = st.selectbox(f"Unidad {u}:", options=st.session_state.lista_operadores, index=idx_op, key=f"asig_u_{u}")
                nuevas_asig[u] = sel_op
            if st.form_submit_button("💾 Guardar Asignaciones"):
                st.session_state.asig_unidad_operador = nuevas_asig
                guardar_catalogos_en_disco()
                st.success("✅ Asignaciones actualizadas correctamente.")
                st.rerun()

    with tab_admin3:
        st.subheader("🗑️️ Eliminación de Elementos de Catálogos")
        tipo_borrar = st.selectbox("Selecciona qué deseas eliminar:", ["Unidad", "Operador", "Cliente/Sucursal"])
        if tipo_borrar == "Unidad":
            uni_a_borrar = st.selectbox("Selecciona Unidad a eliminar:", options=st.session_state.lista_unidades)
            if st.button("🗑️ Eliminar Unidad Seleccionada"):
                if len(st.session_state.lista_unidades) > 1:
                    st.session_state.lista_unidades.remove(uni_a_borrar)
                    if uni_a_borrar in st.session_state.asig_unidad_operador:
                        del st.session_state.asig_unidad_operador[uni_a_borrar]
                    guardar_catalogos_en_disco()
                    st.success(f"✅ Unidad {uni_a_borrar} eliminada.")
                    st.rerun()
                else:
                    st.error("No puedes eliminar todas las unidades.")
        elif tipo_borrar == "Operador":
            op_a_borrar = st.selectbox("Selecciona Operador a eliminar:", options=st.session_state.lista_operadores)
            if st.button("🗑️ Eliminar Operador Seleccionado"):
                if len(st.session_state.lista_operadores) > 1:
                    st.session_state.lista_operadores.remove(op_a_borrar)
                    guardar_catalogos_en_disco()
                    st.success(f"✅ Operador {op_a_borrar} eliminado.")
                    st.rerun()
                else:
                    st.error("No puedes eliminar todos los operadores.")
        elif tipo_borrar == "Cliente/Sucursal":
            cli_a_borrar = st.selectbox("Selecciona Cliente ID:", options=st.session_state.lista_clientes)
            sucursales_cli = list(st.session_state.db_clientes.get(cli_a_borrar, {}).keys())
            suc_a_borrar = st.selectbox("Selecciona Sucursal:", options=sucursales_cli)
            if st.button("🗑️ Eliminar Sucursal Seleccionada"):
                if cli_a_borrar in st.session_state.db_clientes and suc_a_borrar in st.session_state.db_clientes[cli_a_borrar]:
                    del st.session_state.db_clientes[cli_a_borrar][suc_a_borrar]
                    if not st.session_state.db_clientes[cli_a_borrar]:
                        del st.session_state.db_clientes[cli_a_borrar]
                        if cli_a_borrar in st.session_state.lista_clientes:
                            st.session_state.lista_clientes.remove(cli_a_borrar)
                    guardar_catalogos_en_disco()
                    st.success("✅ Sucursal eliminada correctamente.")
                    st.rerun()

    with tab_admin4:
        st.subheader("📊 Historial General y Edición de Registros")
        
        if os.path.exists(ARCHIVO_REGISTROS):
            df_registros = pd.read_excel(ARCHIVO_REGISTROS)
            
            if not df_registros.empty:
                st.dataframe(df_registros, use_container_width=True)
                
                with open(ARCHIVO_REGISTROS, "rb") as file:
                    st.download_button(
                        label="📥 Descargar Reporte Completo (Excel)",
                        data=file,
                        file_name="registros_rutas.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                st.markdown("---")
                st.subheader("🛠️ Modificar o Eliminar Registro Seleccionado")
                
                df_registros["Etiqueta_Fila"] = df_registros.index.astype(str) + " - Unidad: " + df_registros.get("Unidad", "N/D").astype(str) + " (" + df_registros.get("Fecha", "Sin Fecha").astype(str) + ")"
                
                fila_seleccionada_label = st.selectbox("Selecciona el registro a modificar o eliminar:", options=df_registros["Etiqueta_Fila"].tolist())
                
                if fila_seleccionada_label:
                    idx_seleccionado = int(fila_seleccionada_label.split(" - ")[0])
                    registro_actual = df_registros.loc[idx_seleccionado]
                    
                    with st.form(key=f"form_editar_registro_{idx_seleccionado}"):
                        st.write(f"Editando el registro con índice: **{idx_seleccionado}**")
                        
                        datos_editados = {}
                        col1, col2 = st.columns(2)
                        
                        for i, col in enumerate(df_registros.columns):
                            if col == "Etiqueta_Fila":
                                continue
                            valor_actual = registro_actual[col]
                            val_str = "" if pd.isna(valor_actual) else str(valor_actual)
                            
                            with (col1 if i % 2 == 0 else col2):
                                datos_editados[col] = st.text_input(f"{col}:", value=val_str)
                        
                        col_btn1, col_btn2 = st.columns(2)
                        with col_btn1:
                            submit_modificar = st.form_submit_button("💾 Guardar Cambios")
                        with col_btn2:
                            submit_eliminar = st.form_submit_button("🗑️ Eliminar este Registro")
                            
                        if submit_modificar:
                            for col, val in datos_editados.items():
                                df_registros.at[idx_seleccionado, col] = val
                            
                            if "Etiqueta_Fila" in df_registros.columns:
                                df_registros = df_registros.drop(columns=["Etiqueta_Fila"])
                                
                            df_registros.to_excel(ARCHIVO_REGISTROS, index=False)
                            st.success(f"✅ ¡Registro {idx_seleccionado} modificado con éxito!")
                            st.rerun()
                            
                        if submit_eliminar:
                            df_registros = df_registros.drop(idx_seleccionado).reset_index(drop=True)
                            
                            if "Etiqueta_Fila" in df_registros.columns:
                                df_registros = df_registros.drop(columns=["Etiqueta_Fila"])
                                
                            df_registros.to_excel(ARCHIVO_REGISTROS, index=False)
                            st.warning(f"🗑️ ¡Registro {idx_seleccionado} eliminado correctamente!")
                            st.rerun()
            else:
                st.info("No hay registros guardados en el archivo de rutas todavía.")
        else:
            st.info("Aún no se ha generado el archivo de registros de rutas.")

# ==========================================
# ETAPAS DE OPERACIÓN (1 A 3)
# ==========================================
if st.session_state.etapa_idx == 0:
    st.markdown("### 🟢 Etapa 1: Salida de Patio Base")
    with st.form("form_etapa1"):
        lista_uni = st.session_state.lista_unidades
        idx_u = lista_uni.index(st.session_state.unidad_activa) if st.session_state.unidad_activa in lista_uni else 0
        unidad_sel = st.selectbox("Selecciona la Unidad:", options=lista_uni, index=idx_u, key="sel_unidad_e1", on_change=al_cambiar_unidad)
        
        op_sugerido = st.session_state.asig_unidad_operador.get(unidad_sel, st.session_state.lista_operadores[0])
        st.session_state.operador_activo = op_sugerido
        st.text_input("Operador Asignado:", value=op_sugerido, disabled=True)
        
        # Mostrar foto del operador si existe
        nombre_limpio_op = "".join([c if c.isalnum() else "_" for c in op_sugerido])
        foto_encontrada = None
        if os.path.exists(CARPETA_FOTOS_OPERADORES):
            for f_op in os.listdir(CARPETA_FOTOS_OPERADORES):
                if f_op.startswith(nombre_limpio_op + "."):
                    foto_encontrada = os.path.join(CARPETA_FOTOS_OPERADORES, f_op)
                    break
        if foto_encontrada:
            tam_foto = config_actual.get("tamano_foto_operador", 120)
            st.image(foto_encontrada, width=tam_foto, caption=f"Operador: {op_sugerido}")

        km_salida = st.number_input("Kilometraje de Salida de Patio:", min_value=0.0, value=float(st.session_state.km_salida_patio), step=1.0)
        
        if st.form_submit_button("💾 Guardar y Avanzar a Pedrera ➡️"):
            st.session_state.unidad_activa = unidad_sel
            st.session_state.km_salida_patio = km_salida
            st.session_state.hora_salida_patio = obtener_hora_mexico()
            st.session_state.etapa_idx = 1
            guardar_borrador_de_disco()
            st.success("✅ Salida de Patio registrada con éxito.")
            st.rerun()

elif st.session_state.etapa_idx == 1:
    st.markdown("### 🟡 Etapa Intermedia: Pedrera (Llegada / Salida)")
    with st.form("form_etapa2"):
        st.write(f"**Unidad Activa:** {st.session_state.unidad_activa} | **Operador:** {st.session_state.operador_activo}")
        
        col_llegada, col_salida = st.columns(2)
        with col_llegada:
            km_llegada_p = st.number_input("Km Llegada a Pedrera:", min_value=0.0, value=float(st.session_state.km_llegada_pedrera), step=1.0)
            if st.form_submit_button("📍 Registrar Hora Llegada Pedrera"):
                st.session_state.km_llegada_pedrera = km_llegada_p
                st.session_state.hora_llegada_pedrera = obtener_hora_mexico()
                guardar_borrador_de_disco()
                st.success(f"✅ Llegada a Pedrera registrada: {st.session_state.hora_llegada_pedrera}")
                st.rerun()
                
        with col_salida:
            km_salida_p = st.number_input("Km Salida de Pedrera (Cargado):", min_value=0.0, value=float(st.session_state.km_salida_pedrera), step=1.0)
            if st.form_submit_button("🚀 Registrar Hora Salida Pedrera"):
                st.session_state.km_salida_pedrera = km_salida_p
                st.session_state.hora_salida_pedrera = obtener_hora_mexico()
                st.session_state.etapa_idx = 2
                guardar_borrador_de_disco()
                st.success(f"✅ Salida de Pedrera registrada: {st.session_state.hora_salida_pedrera}. Avanzando a Cliente...")
                st.rerun()

elif st.session_state.etapa_idx == 2:
    st.markdown("### 🟠 Etapa 2: Llegada / Salida con el Cliente")
    with st.form("form_etapa3"):
        st.write(f"**Unidad Activa:** {st.session_state.unidad_activa} | **Operador:** {st.session_state.operador_activo}")
        
        lista_c_str = [str(c).strip().replace(".0", "") for c in st.session_state.lista_clientes]
        idx_cli = lista_c_str.index(str(st.session_state.sel_cliente_e2)) if str(st.session_state.sel_cliente_e2) in lista_c_str else 0
        
        sel_c = st.selectbox("Número de Cliente:", options=lista_c_str, index=idx_cli, key="sel_cliente_e2", on_change=al_cambiar_cliente_e2)
        
        sucursales_disp = list(st.session_state.db_clientes.get(sel_c, {"Principal": {}}).keys())
        sucursales_disp = [s for s in sucursales_disp if str(s).lower() != 'nan' and str(s).strip() != '']
        if not sucursales_disp: sucursales_disp = ["Principal"]
        
        idx_suc = sucursales_disp.index(st.session_state.sel_sucursal_e2) if st.session_state.sel_sucursal_e2 in sucursales_disp else 0
        sel_s = st.selectbox("Sucursal / Razón Social:", options=sucursales_disp, index=idx_suc, key="sel_sucursal_e2", on_change=al_cambiar_sucursal_e2)
        
        st.text_input("Contacto en Sucursal:", value=st.session_state.cli_contacto, disabled=True)
        st.text_input("Teléfono:", value=st.session_state.cli_telefono, disabled=True)
        
        col_cl_llegada, col_cl_salida = st.columns(2)
        with col_cl_llegada:
            km_ll_cli = st.number_input("Km Llegada a Cliente:", min_value=0.0, value=float(st.session_state.km_llegada_cliente), step=1.0)
            if st.form_submit_button("📍 Registrar Llegada con Cliente"):
                st.session_state.km_llegada_cliente = km_ll_cli
                st.session_state.hora_llegada_cliente = obtener_hora_mexico()
                guardar_borrador_de_disco()
                st.success("✅ Llegada con cliente registrada.")
                st.rerun()
                
        with col_cl_salida:
            km_sal_cli = st.number_input("Km Salida de Cliente (Descargado):", min_value=0.0, value=float(st.session_state.km_salida_cliente), step=1.0)
            if st.form_submit_button("🏁 Registrar Salida de Cliente"):
                st.session_state.km_salida_cliente = km_sal_cli
                st.session_state.hora_salida_cliente = obtener_hora_mexico()
                st.session_state.etapa_idx = 3
                guardar_borrador_de_disco()
                st.success("✅ Salida de cliente registrada. Avanzando a Cierre de Ruta...")
                st.rerun()

elif st.session_state.etapa_idx == 3:
    st.markdown("### 🔴 Etapa 3: Cierre de Ruta / Regreso a Patio")
    with st.form("form_etapa4"):
        st.write(f"**Unidad Activa:** {st.session_state.unidad_activa} | **Operador:** {st.session_state.operador_activo}")
        km_final = st.number_input("Kilometraje Final (Regreso a Patio):", min_value=0.0, value=float(st.session_state.km_final_viaje), step=1.0)
        
        if st.form_submit_button("💾 Guardar y Finalizar Ruta Completa"):
            st.session_state.km_final_viaje = km_final
            st.session_state.hora_cierre_viaje = obtener_hora_mexico()
            
            # Consolidar registro para guardar en Excel de rutas
            registro_final = {
                "Fecha": datetime.now().strftime("%Y-%m-%d"),
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
                "Hora Cierre Viaje": st.session_state.hora_cierre_viaje
            }
            guardar_en_excel(registro_final)
            
            # Reiniciar estado para un nuevo viaje
            st.session_state.etapa_idx = 0
            st.session_state.ruta_guardada = True
            if os.path.exists(ARCHIVO_BORRADOR):
                try: os.remove(ARCHIVO_BORRADOR)
                except: pass
                
            st.success("🎉 ¡Ruta cerrada y guardada exitosamente en el historial!")
            st.rerun()
