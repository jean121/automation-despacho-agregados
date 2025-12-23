import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def enviar_correo():
    remitente = "ealpiste@unicon.com.pe"
    contraseña = "L1m4.2025$%"
    destinatario = "ealpiste@unicon.com.pe"

    asunto = "Correo programado"
    cuerpo = "Hola, este correo se envió automáticamente."

    mensaje = MIMEMultipart()
    mensaje["From"] = remitente
    mensaje["To"] = destinatario
    mensaje["Subject"] = asunto
    mensaje.attach(MIMEText(cuerpo, "plain"))

    try:
        servidor = smtplib.SMTP("smtp.gmail.com", 587)
        servidor.starttls()
        servidor.login(remitente, contraseña)
        servidor.send_message(mensaje)
        servidor.quit()
        print("Correo enviado con éxito.")
    except Exception as e:
        print(f"Error al enviar el correo: {e}")

# Esta parte hace que se ejecute la función
if __name__ == "__main__":
    enviar_correo()