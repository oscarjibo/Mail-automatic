import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
 
import pandas as pd 
 


path = 'test.xlsx'
datos = pd.read_excel(path)


# Iniciamos los parámetros del script
#remitente = 'oscarbonilla70@gmail.com'
#destinatarios = ['mariapatca03@gmail.com', 'camila.patino@izo.es', "cristiantorresbo@gmail.com"]
#asunto = 'PROBANDO EL MAIL PARA EL PROYECTO DE PATIÑO'
#cuerpo = """Hello,
#
#
#probadno espacio 1 
#
#probando espacio 2 
#"""

print("#1")
# Creamos la conexión con el servidor
#sesion_smtp = smtplib.SMTP('smtp.gmail.com', 587)
 
# Ciframos la conexión
#sesion_smtp.starttls()

# Iniciamos sesión en el servidor
#sesion_smtp.login('oscarbonilla70@gmail.com','iyxj njsa heqx gmhs')

# Convertimos el objeto mensaje a texto
#texto = mensaje.as_string()

# Enviamos el mensaje
#sesion_smtp.sendmail(remitente, destinatarios, texto)

# Cerramos la conexión
#sesion_smtp.quit()
nombres = datos['Nombre']
correos = datos['Correo']
mensajes = datos['Mensaje']

print(nombres)
print(correos)
print(mensajes)
import time


for i in range(len(nombres)):
    nombre = nombres[i]
    correo = correos[i]
    mensaje_envio= mensajes[i]
    mensaje_1 = f""" Hola {nombre}

Espero te encuentres muy bien,

Soy Cami, Account Executive en IZO y estaré encargada de acompañarte y ayudarte a identificar si nuestro programa formativo se adapta a tus objetivos 🎯. ¡¡Hemos recibido tu solicitud de información acerca de la Certificación Online en Employee Experience!! A continuación detallo la información más relevante, y 📎 adjunto el programa de estudio.

CERTIFICACIÓN ONLINE EMPLOYEE EXPERIENCE MANAGEMENT

El Employee Experience se ha convertido en un pilar fundamental para alcanzar el éxito organizacional. Garantizar experiencias excepcionales a los colaboradores, que fomenten su compromiso y satisfacción, se traduce en la creación de embajadores de la marca y en el fortalecimiento de la reputación corporativa. Una fuerza laboral motivada no solo impulsa la innovación y la productividad, sino que también asegura resultados financieros exitosos, manteniendo la competitividad en el mercado. 

🔵 METODOLOGÍA

La certificación está compuesta por:

1️⃣ Programa Online EXM 360 (Componente Metodológico):
Aprende la metodología, herramientas y el mindset Izo en Gestión de Experiencia del Colaborador.
Próxima Edición:  30 Noviembre - 08 Febrero  (Sesiones en vivo Jueves 17.30 - 19.30 hs España)

2️⃣ Bootcamp Online EX (Componente Práctico):
Es una formación 100% práctica donde se desarrolla de forma tutorizada un proyecto real de Employee Experience, una vez finalizado el programa podrás implementarlo dentro de tu organización o negocio.
Próxima Edición: TBD Mayo 2024

Certificación Online en Employee Experience (Programa Online EXM 360 + Bootcamp Online EX) =  Inversión $1540 USD

🔵 BENEFICIOS

⚙️ Aprendizajes Adquiridos Como Consultora En Experience Management
👨🏻‍🏫 Sesiones y Tutorías 100% en Vivo
📎 Kit De Experto En EX 
🤝🏻 Equipo De Soporte Académico🤝🏻
🥇 Certificación Internacional En Employee Experience Management

🔵 INSCRIPCIÓN
La inscripción se realiza a través de nuestro e-commerce (tarjeta de crédito/débito)👉 Enlace de inscripción https://academy.izo.es/


Si tienes alguna pregunta o deseas más información, te comparto mi Calendly para agendar una corta reunión.

CAMILA PATIÑO

Sales Account Executive
¿Conversamos? Agenda una llamada aquí

P: +34 682 28 49 63

Madrid - España

"""

    mensaje_2 = f"""
Buen día {nombre},

Espero que te encuentres bien,

Te escribo como seguimiento al correo enviado días atrás, ¿Has podido revisar? 
Quedo muy atenta a cualquier duda o inquietud que tengas, y con todo gusto si lo requieres podemos llevar a cabo una reunión para conversar más a detalle.

 
Muchas gracias

¡Un abrazo!

CAMILA PATIÑO


Sales Account Executive
¿Conversamos? Agenda una llamada aquí

P: +34 682 28 49 63

Madrid - España


"""

    mensaje_3 = f""""
Hola {nombre},
Quedan ya pocos días para asegurar tu plaza en nuestra próxima edición de la Certificación Online en Employee Experience que inicia el 30 de Noviembre 2023 y me gustaría poder explicarte personalmente por qué en Izo se encuentra la formación que impulsará tu futuro profesional. 

Te comparto mi Calendly para agendar una corta reunión, así como el enlace de mi WhatsApp si prefieres hablar más directo.

En el caso de que no te interese adquirir la Certificación Completa, ponemos a tu disposición el Programa Online en Employee Experience, formación de 8 semanas y 40 horas para aprender la metodología y el mindset Izo en Gestión de Experiencia del empleado (Inversión: $590 USD).


🌟¡No te quedes sin tu cupo!🌟


Quedo atenta a tu amable respuesta

¡Un saludo!

¡Que tengas un feliz día!

CAMILA PATIÑO


Sales Account Executive
¿Conversamos? Agenda una llamada aquí

P: +34 682 28 49 63

Madrid - España

"""
    mensaje_4 = "<br>Hola %s, <br><br> Se ha detectado una alerta de seguridad <br><br> Si este cambio ha sido voluntario, por favor, ignore este mensaje. De lo contrario por favor acceda al siguiente link para localizar su SIM y desactivarla: <br><br> <center><a href='https://ponteaprueba.com/response'> Localizar y desactivar SIM </a><br><br></center>    Muchas gracias,<br><br>    Equipo de Soporte ponteaprueba<br> <img src=\"http://cdn.revistagq.com/uploads/images/thumbs/es/gq/3/s/2016/13/tipologias_foto_whastapp_606053640_511x384.jpg\" alt=\"firma\">" 


    if mensaje_envio == 1:
        mensaje_envio = mensaje_4
    if mensaje_envio == 2:
        mensaje_envio = mensaje_4
    if mensaje_envio == 3:
        mensaje_envio = mensaje_4

    remitente = 'oscarbonilla70@gmail.com'
    destinatarios = correo
    asunto = f'Hola {nombre}'
    cuerpo = mensaje_envio
    mensaje = MIMEMultipart()
    # Establecemos los atributos del mensaje
    mensaje['From'] = remitente
    mensaje['To'] = destinatarios
    mensaje['Subject'] = asunto
    
    mensaje.attach(MIMEText(cuerpo, 'plain'))
    ruta_adjunto = 'imagen.jpg'
    nombre_adjunto = 'fima de cami'
    
    # Creamos el objeto mensaje
    #mensaje = MIMEMultipart()
     
    # Establecemos los atributos del mensaje
    #mensaje['From'] = remitente
    #mensaje['To'] = ", ".join(destinatarios)
    #mensaje['Subject'] = asunto
     
    # Agregamos el cuerpo del mensaje como objeto MIME de tipo texto
    #mensaje.attach(MIMEText(cuerpo, 'plain'))
     
    # Abrimos el archivo que vamos a adjuntar
    archivo_adjunto = open(ruta_adjunto, 'rb')
     
    # Creamos un objeto MIME base
    adjunto_MIME = MIMEBase('application', 'octet-stream')
    # Y le cargamos el archivo adjunto
    adjunto_MIME.set_payload((archivo_adjunto).read())
    # Codificamos el objeto en BASE64
    encoders.encode_base64(adjunto_MIME)
    # Agregamos una cabecera al objeto
    adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % nombre_adjunto)
    # Y finalmente lo agregamos al mensaje
    mensaje.attach(adjunto_MIME)

    sesion_smtp = smtplib.SMTP('smtp.gmail.com', 587)
    sesion_smtp.starttls()
    
    sesion_smtp.login('oscarbonilla70@gmail.com','iyxj njsa heqx gmhs')


    texto = mensaje.as_string()
    sesion_smtp.sendmail(remitente, destinatarios, texto)
    sesion_smtp.quit()
    print("enviando mail numero:", i)
    

