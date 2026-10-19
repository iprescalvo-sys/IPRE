import streamlit as st
from datetime import datetime
import pytz
import io
import smtplib
from email.message import EmailMessage
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from google import genai

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
    """Llama a la API de Gemini solo para responder las preguntas de los temas (Act 2.3 y Act 3)."""
    client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
    contexto = (
        "Eres un asistente participando en un estudio académico de la PUC. "
        "Responde de forma clara, informativa y cordial a la pregunta del usuario sobre sus temas de interés. "
        "No agregues preguntas de cierre porque el sistema las añade automáticamente.\n\n"
    )
    for h in historial_ia:
        contexto += f"{h['rol']}: {h['texto']}\n"
    contexto += f"Usuario: {pregunta_usuario}\nAsistente:"
    
    # Lista de modelos en orden de prioridad por si alguno cambia de nombre
    modelos = ["gemini-3.8-flash", "gemini-3-flash-preview", "gemini-2.0-flash"]
    
    ultimo_error = ""
    for modelo in modelos:
        try:
            response = client.models.generate_content(
                model=modelo,
                contents=contexto
            )
            return response.text.strip()
        except Exception as e:
            ultimo_error = str(e)
            continue
            
    return f"Interesante pregunta. (Nota técnica: {ultimo_error})"

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
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF')
    )

    for col_idx, th in enumerate(table_headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=th)
        cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin_border

    # Colores por actividad según las instrucciones
    def obtener_estilo_actividad(act_nombre):
        if act_nombre.startswith("Act 1"):
            return "D9E1F2", "000000"  # Azul muy claro
        elif act_nombre.startswith("Act 2"):
            return "B4C6E7", "000000"  # Azul claro
        elif act_nombre.startswith("Act 3"):
            return "8EA9DB", "000000"  # Azul medio
        else:
            return "2F5597", "FFFFFF"  # Azul oscuro ejecutivo con texto blanco (Act 4)

    for row_idx, r in enumerate(registros_act, start=4):
        bg_color, txt_color = obtener_estilo_actividad(r["act"])
        fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        font = Font(name="Calibri", size=10, color=txt_color)

        valores_fila = [r["t_bot"], r["m_bot"], r["t_user"], r["r_user"], r["act"]]
        for col_idx, val in enumerate(valores_fila, start=1):
            c = ws.cell(row=row_idx, column=col_idx, value=val)
            c.fill = fill
            c.font = font
            c.border = thin_border
            if col_idx in [1, 3, 5]:
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                c.alignment = Alignment(vertical="center", wrap_text=True)

    anchos = {"A": 18, "B": 50, "C": 18, "D": 50, "E": 20, "F": 25}
    for col_letter, width in anchos.items():
        ws.column_dimensions[col_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def enviar_excel_por_correo(excel_bytes, datos_demograficos):
    """Envía el Excel silenciosamente al correo del investigador."""
    try:
        msg = EmailMessage()
        timestamp_str = datetime.now(ZONA_HORARIA).strftime("%Y%m%d_%H%M%S")
        msg["Subject"] = f"[Estudio IA UC] Nuevo registro participante - {timestamp_str}"
        msg["From"] = st.secrets["EMAIL_SENDER"]
        msg["To"] = st.secrets["EMAIL_RECEIVER"]
        
        cuerpo = (
            f"Se ha completado una nueva interacción en el estudio.\n\n"
            f"Resumen del participante:\n"
            f"- Edad: {datos_demograficos.get('p1', 'N/A')}\n"
            f"- Género: {datos_demograficos.get('p2', 'N/A')}\n"
            f"- Colegio: {datos_demograficos.get('p3', 'N/A')}\n"
            f"- Carrera: {datos_demograficos.get('p4', 'N/A')}\n"
            f"- Rendimiento: {datos_demograficos.get('p5', 'N/A')}\n"
            f"- Región: {datos_demograficos.get('p6', 'N/A')}\n"
        )
        msg.set_content(cuerpo)

        nombre_archivo = f"Participante_{timestamp_str}.xlsx"
        msg.add_attachment(
            excel_bytes.getvalue(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=nombre_archivo
        )

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(st.secrets["EMAIL_SENDER"], st.secrets["EMAIL_PASSWORD"])
            smtp.send_message(msg)
        return True
    except Exception as e:
        print(f"Error enviando correo: {e}")
        return False

# ==========================================
# 3. GESTIÓN DE ESTADO Y FLUJO DE ACTIVIDADES
# ==========================================
TEXTO_CONSENTIMIENTO = """**Profundidad en la interacción con la Inteligencia Artificial Generativa**
**Pontificia Universidad Católica de Chile – Escuela de Ingeniería**

Su participación consistirá en una única intervención en la cual debe trabajar con una Inteligencia Artificial generativa. Esto no le deberá tomar más de 10 o 15 minutos aproximadamente. Antes de empezar usted deberá indicar: “acepto participar”, o “no acepto participar” Al escribir “acepto participar”, usted da su consentimiento para que la información recopilada se utilice para únicamente fines académicos y/o de investigación. Por lo tanto, antes de señalar “acepto participar”, lea cuidadosamente el texto a continuación:

- Su participación es absolutamente voluntaria, por lo que no hay obligación alguna de participar en este estudio. Más aún, puede dejar de participar en cualquier momento, no estando obligado a responder pregunta alguna.
- La información proporcionada por usted se utilizará solo para el propósito de esta investigación, y las publicaciones que resultan de ella.
- Asimismo, los datos podrán ser utilizados en investigaciones futuras que sigan la misma línea investigativa de este estudio.
- Los datos serán almacenados por un período de 1 año, una vez finalizada la investigación, siendo la fecha de término estimada agosto de 2027.
- Los datos obtenidos en esta investigación y los datos individuales son estrictamente confidenciales y anónimos. No se divulgará información individual alguna en las presentaciones y publicaciones de los resultados.
- Toda la información obtenida será protegida a través de claves de acceso, permitiendo así solo el equipo que trabaja en el proyecto pueda revisarlo. Toda la información obtenida será destruida luego de 1 año de finalizado el proyecto.
- No hay beneficios directos asociados a la participación en el estudio.
- No hay riesgos asociados a su participación.

Considerando lo anterior ¿Está Ud. dispuesto a realizar la actividad?"""

if "etapa" not in st.session_state:
    st.session_state.etapa = "ACT_0_CONSENTIMIENTO"
    st.session_state.mensajes_chat = [{"rol": "assistant", "texto": TEXTO_CONSENTIMIENTO}]
    st.session_state.ultimo_mensaje_bot = TEXTO_CONSENTIMIENTO
    st.session_state.ultima_hora_bot = obtener_hora()
    st.session_state.datos_demograficos = {}
    st.session_state.registros_act = [] # Guarda desde Act 1 en adelante
    st.session_state.historial_ia = []
    st.session_state.temas_usuario = ""
    st.session_state.iteracion_act3 = 1
    st.session_state.correo_enviado = False

def enviar_mensaje_bot(texto):
    st.session_state.mensajes_chat.append({"rol": "assistant", "texto": texto})
    st.session_state.ultimo_mensaje_bot = texto
    st.session_state.ultima_hora_bot = obtener_hora()

def registrar_interaccion(respuesta_usuario, hora_usuario, etiqueta_actividad):
    st.session_state.registros_act.append({
        "t_bot": st.session_state.ultima_hora_bot,
        "m_bot": st.session_state.ultimo_mensaje_bot,
        "t_user": hora_usuario,
        "r_user": respuesta_usuario,
        "act": etiqueta_actividad
    })

def procesar_entrada(texto_usuario):
    hora_usr = obtener_hora()
    st.session_state.mensajes_chat.append({"rol": "user", "texto": texto_usuario})
    etapa = st.session_state.etapa

    # REGLA DE TÉRMINO INMEDIATO ("Tomate")
    if texto_usuario.strip().lower() == "tomate" and etapa != "FINALIZADO":
        if etapa.startswith("ACT_"):
            registrar_interaccion(texto_usuario, hora_usr, "Act 4 (Término Tomate)")
        despedida = "Muchas gracias por tu tiempo y participación. ¡Que tengas un excelente día!"
        enviar_mensaje_bot(despedida)
        st.session_state.etapa = "FINALIZADO"
        return

    # FLUJO PASO A PASO
    if etapa == "ACT_0_CONSENTIMIENTO":
        if texto_usuario == "Acepto participar":
            st.session_state.etapa = "P1"
            enviar_mensaje_bot("Pregunta 1: ¿Qué edad tiene usted? (Años)")
        else:
            enviar_mensaje_bot("Entendido. Muchas gracias por su tiempo. Hasta luego.")
            st.session_state.etapa = "CANCELADO"

    elif etapa == "P1":
        st.session_state.datos_demograficos["p1"] = texto_usuario
        st.session_state.etapa = "P2"
        enviar_mensaje_bot("Pregunta 2: ¿Cuál es su género?")

    elif etapa == "P2":
        st.session_state.datos_demograficos["p2"] = texto_usuario
        st.session_state.etapa = "P3"
        enviar_mensaje_bot("Pregunta 3: Usted estudió en:")

    elif etapa == "P3":
        st.session_state.datos_demograficos["p3"] = texto_usuario
        st.session_state.etapa = "P4"
        enviar_mensaje_bot("Pregunta 4: Usted es alumno de la carrera:")

    elif etapa == "P4":
        st.session_state.datos_demograficos["p4"] = texto_usuario
        st.session_state.etapa = "P5"
        enviar_mensaje_bot("Pregunta 5: Usted se considera:")

    elif etapa == "P5":
        st.session_state.datos_demograficos["p5"] = texto_usuario
        st.session_state.etapa = "P6"
        enviar_mensaje_bot("Pregunta 6: ¿A qué región pertenece (donde ha vivido la mayoría de su vida o donde vive actualmente)?")

    elif etapa == "P6":
        st.session_state.datos_demograficos["p6"] = texto_usuario
        st.session_state.etapa = "ACT_1"
        enviar_mensaje_bot("Descríbete en un par de líneas. Eres libre de indicar lo que te parezca más representativo de ti")

    elif etapa == "ACT_1":
        registrar_interaccion(texto_usuario, hora_usr, "Act 1")
        # Conteo exacto de caracteres en Python (>= 150)
        if len(texto_usuario) >= 150:
            st.session_state.etapa = "ACT_2_1"
            enviar_mensaje_bot("Cuéntame sobre temas que te interesen")
        else:
            enviar_mensaje_bot("¿Podrías profundizar un poco más en tu descripción y contarme más detalles sobre ti?")

    elif etapa == "ACT_2_1":
        registrar_interaccion(texto_usuario, hora_usr, "Act 2")
        st.session_state.temas_usuario = texto_usuario
        st.session_state.etapa = "ACT_2_2"
        enviar_mensaje_bot(f"Hazme una pregunta de algo que te interese de alguno de los temas anteriores:\n{texto_usuario}")

    elif etapa == "ACT_2_2":
        registrar_interaccion(texto_usuario, hora_usr, "Act 2")
        respuesta_ia = responder_con_ia(st.session_state.historial_ia, texto_usuario)
        st.session_state.historial_ia.append({"rol": "Usuario", "texto": texto_usuario})
        st.session_state.historial_ia.append({"rol": "Asistente", "texto": respuesta_ia})
        msg_completo = f"{respuesta_ia}\n\n¿Te gustaría continuar con la conversación o terminarla?"
        st.session_state.etapa = "ACT_2_3_DECISION"
        enviar_mensaje_bot(msg_completo)

    elif etapa == "ACT_2_3_DECISION":
        if texto_usuario == "Continuar":
            registrar_interaccion(texto_usuario, hora_usr, f"Act 3.{st.session_state.iteracion_act3}")
            st.session_state.etapa = "ACT_3_PREGUNTA"
            enviar_mensaje_bot("¿Qué te gustaría profundizar sobre el tema?")
        elif texto_usuario == "Terminar":
            registrar_interaccion(texto_usuario, hora_usr, "Act 4.1")
            st.session_state.etapa = "ACT_4_1"
            enviar_mensaje_bot("Muchas gracias por tu tiempo. ¿Qué te hizo querer terminar la interacción?")

    elif etapa == "ACT_3_PREGUNTA":
        etiqueta = f"Act 3.{st.session_state.iteracion_act3}"
        registrar_interaccion(texto_usuario, hora_usr, etiqueta)
        respuesta_ia = responder_con_ia(st.session_state.historial_ia, texto_usuario)
        st.session_state.historial_ia.append({"rol": "Usuario", "texto": texto_usuario})
        st.session_state.historial_ia.append({"rol": "Asistente", "texto": respuesta_ia})
        st.session_state.iteracion_act3 += 1
        msg_completo = f"{respuesta_ia}\n\n¿Te gustaría continuar con la conversación o terminarla?"
        st.session_state.etapa = "ACT_2_3_DECISION"
        enviar_mensaje_bot(msg_completo)

    elif etapa == "ACT_4_1":
        registrar_interaccion(texto_usuario, hora_usr, "Act 4.2")
        st.session_state.etapa = "FINALIZADO"
        enviar_mensaje_bot("Muchas gracias nuevamente por haber participado en este estudio. ¡Que tengas un excelente día!")

# ==========================================
# 4. RENDERIZADO DE LA INTERFAZ Y BOTONES
# ==========================================
etapa_actual = st.session_state.etapa
etapas_texto_libre = ["P1", "P4", "P6", "ACT_1", "ACT_2_1", "ACT_2_2", "ACT_3_PREGUNTA", "ACT_4_1"]

# 1. Capturar el texto libre ANTES de dibujar los mensajes (SIN usar st.rerun)
if etapa_actual in etapas_texto_libre:
    entrada_usuario = st.chat_input("Escribe tu respuesta aquí...")
    if entrada_usuario:
        procesar_entrada(entrada_usuario)
        etapa_actual = st.session_state.etapa  # Actualizamos la etapa inmediatamente

# 2. Dibujar todos los mensajes del chat
for idx, msg in enumerate(st.session_state.mensajes_chat):
    if idx == 0 and st.session_state.etapa != "ACT_0_CONSENTIMIENTO":
        with st.expander("📄 Ver texto de Consentimiento Informado (Aceptado)"):
            st.markdown(msg["texto"])
    else:
        with st.chat_message(msg["rol"]):
            st.markdown(msg["texto"])

# 3. Mostrar botones interactivos cuando corresponda según la etapa
opciones_botones = {
    "ACT_0_CONSENTIMIENTO": ["Acepto participar", "No acepto participar"],
    "P2": ["Mujer", "Hombre", "Otro", "No desea responder"],
    "P3": ["Colegio municipal o público", "Colegio particular subvencionado", "Colegio particular pagado"],
    "P5": ["Muy buen alumno", "Bueno o del promedio superior", "Regular o del promedio inferior"],
    "ACT_2_3_DECISION": ["Continuar", "Terminar"]
}

if etapa_actual in opciones_botones:
    st.write("Seleccione una opción:")
    cols = st.columns(len(opciones_botones[etapa_actual]))
    for idx, opcion in enumerate(opciones_botones[etapa_actual]):
        if cols[idx].button(opcion, use_container_width=True):
            procesar_entrada(opcion)
            st.rerun()

# 4. Al finalizar, enviar el Excel automáticamente por correo en segundo plano
if etapa_actual == "FINALIZADO" and not st.session_state.correo_enviado:
    excel_io = generar_excel_en_memoria(st.session_state.datos_demograficos, st.session_state.registros_act)
    exito = enviar_excel_por_correo(excel_io, st.session_state.datos_demograficos)
    st.session_state.correo_enviado = True
    if exito:
        st.success("Tus respuestas han sido registradas exitosamente. Ya puedes cerrar esta ventana.")

# 5. Ancla invisible y Auto-Scroll universal para todas las versiones de Streamlit
st.markdown('<div id="fin-del-chat"></div>', unsafe_allow_html=True)
components.html(
    f"""
    <script>
        // Turno actual: {len(st.session_state.mensajes_chat)}
        setTimeout(function() {{
            const ancla = window.parent.document.getElementById('fin-del-chat');
            if (ancla) {{
                ancla.scrollIntoView({{ behavior: 'smooth', block: 'end' }});
            }}
        }}, 100);
    </script>
    """,
    height=0,
)
