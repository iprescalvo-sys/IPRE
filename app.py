import streamlit as st
from datetime import datetime
import pytz
import io
import smtplib
from email.message import EmailMessage
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from google import genai
import time
import streamlit.components.v1 as components

# ==========================================
# 1. CONFIGURACIÓN INICIAL Y SECRETOS
# ==========================================
st.set_page_config(page_title="Estudio Interacción IA - UC", page_icon="💬", layout="centered")

# Configura estas credenciales en los Secrets de Streamlit (.streamlit/secrets.toml)
# GEMINI_API_KEY = "AIzaSy..."
# EMAIL_SENDER = "tu_correo@gmail.com"
# EMAIL_PASSWORD = "xxxx xxxx xxxx xxxx" # Contraseña de aplicación de 16 letras
# EMAIL_RECEIVER = "investigador@uc.cl"

ZONA_HORARIA = pytz.timezone("America/Santiago")

def obtener_hora():
    return datetime.now(ZONA_HORARIA).strftime("%H:%M")

# Mantenemos abierta la conexión con Google para ahorrar tiempo de autenticación
@st.cache_resource
def obtener_cliente_ia():
    return genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

def responder_con_ia(historial_ia, pregunta_usuario):
    """Llama a la API de Gemini con velocidad optimizada y respaldo anti-caídas."""
    client = obtener_cliente_ia()
    
    contexto = (
        "Eres un asistente participando en un estudio académico de la PUC. "
        "Responde de forma clara, informativa, CORDIAL Y BREVE (máximo 2 párrafos). "
        "No agregues preguntas de cierre porque el sistema las añade automáticamente.\n\n"
    )
    
    # Solo enviamos el historial reciente para que lea más rápido
    historial_reciente = historial_ia[-3:] if len(historial_ia) > 3 else historial_ia
    for h in historial_reciente:
        contexto += f"{h['rol']}: {h['texto']}\n"
        
    contexto += f"Usuario: {pregunta_usuario}\nAsistente:"
    
    # Sistema de reintentos rápidos (Fail-safe)
    for intento in range(3):
        try:
            response = client.models.generate_content(
                model="gemini-3.8-flash",
                contents=contexto,
                config={"max_output_tokens": 200} # Respuesta corta y rápida
            )
            return response.text.strip()
        except Exception as e:
            error_str = str(e)
            # Si el servidor está saturado, esperamos 1.5 segundos y reintentamos en silencio
            if "503" in error_str or "429" in error_str or "UNAVAILABLE" in error_str:
                time.sleep(1.5)
                continue
            else:
                break # Si es otro tipo de error grave, sale del bucle

    # =========================================================================
    # RESPUESTA COMODÍN (SALVAVIDAS DEL ESTUDIO)
    # Si después de los intentos Google sigue caído, enviamos este texto.
    # =========================================================================
    return "Comprendo tu punto. Es una perspectiva interesante, ¿podrías darme un ejemplo o profundizar un poco más en tu idea?"

# ==========================================
# 2. GENERACIÓN DEL EXCEL Y ENVÍO POR CORREO
# ==========================================
def generar_excel_en_memoria(datos_demograficos, registros_act):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Registro Interaccion"

    # Fila 1: Respuestas de Preguntas de datos en columnas distintas
    headers_meta = ["P1: Edad", "P2: Género", "P3: Colegio", "
