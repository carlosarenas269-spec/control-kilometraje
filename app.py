import os
import json
import base64
import re
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
from PIL import Image
from sqlalchemy import create_engine

# ==========================================
# CONFIGURACIÓN Y PERSISTENCIA DE SUPABASE Y DATOS
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

@st.cache_resource
def obtener_motor_db():
    try:
        if "DATABASE_URL" in st.secrets:
            db_url = st.secrets["DATABASE_URL"]
            if db_url.startswith("postgres://"):
                db_url = db_url.replace("postgres://", "postgresql+psycopg2://", 1)
            elif db_url.startswith("postgresql://") and "+psycopg2" not in db_url:
                db_url = db_url.replace("postgresql://", "postgresql+psycopg2://", 1)
            return create_engine(db_url)
    except Exception:
        pass
    return None

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
            "km_salida_cliente": st.session_state.get("km_salida_cliente", 14
