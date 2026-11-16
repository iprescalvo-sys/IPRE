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

def responder_con_ia(historial_ia, pregunta_usuario):
    """Llama a la API de Gemini (versión prioritaria) con reintentos para cortes de red."""
    client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
    contexto = (
        "Eres un asistente participando en un estudio académico de la PUC. "
        "Responde de forma clara, informativa y cordial a la pregunta del usuario sobre sus temas de interés. "
        "No agregues preguntas de cierre porque el sistema las añade automáticamente.\n\n"
    )
    for h in historial_ia:
        contexto += f"{h['rol']}: {h['texto']}\n"
    contexto += f"Usuario: {pregunta_usuario}\nAsistente:"
    
    # Modelo estable para cuentas con plan
    modelo = "gemini-1.5-flash"
    max_reintentos = 3
    
    ultimo_error = ""
    for intento in range(max_reintentos):
        try:
            response = client.models.generate_content(
                model=modelo,
                contents=contexto
            )
            return response.text.strip()
        except Exception as e:
            ultimo_error = str(e)
            time.sleep(2)  # Pausa breve si hay inestabilidad en la red
            continue
            
    # Si todo falla (ej: se cae tu internet o Google Cloud), muestra un error claro
    return f"Experimenté un problema de conexión al procesar tu solicitud. Por favor, intenta de nuevo. (Nota: {ultimo_error})"

# ==========================================
# 2. GENERACIÓN DEL EXCEL Y ENVÍO POR CORREO
# ==========================================
def generar_excel_en_memoria(datos_demograficos, registros_act):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Registro Interaccion"

    # Fila 1: Respuestas de Preguntas de datos en columnas distintas
    headers_meta = ["P1: Edad", "P2: Género", "P3: Colegio", "P4: Carrera", "P5: Rendimiento", "P6: Región"]
    valores_meta = [
        datos_demograficos.get("p1", ""),
        datos_demograficos.get("p2", ""),
        datos_demograficos.get("p3", ""),
        datos_demograficos.get("p4", ""),
        datos_demograficos.get("p5", ""),
        datos_demograficos.get("p6", "")
    ]

    for col_idx, (h, val) in enumerate(zip(headers_meta, valores_meta), start=1):
        cell = ws.cell(row=1, column=col_idx, value=f"{h}: {val}")
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

    # Fila 3: Encabezados de la tabla
    table_headers = [
        "Tiempo Chatbot",
        "Mensajes del Chatbot desde la Act 1",
        "Tiempo Usuario",
        "Respuestas del Usuario desde la Act 1",
        "Actividad Correspondiente"
    ]
    
    thin_border = Border(
        left=Side(
