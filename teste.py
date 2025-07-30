import asyncio
import os
from email.message import EmailMessage
from aiosmtplib import send
from dotenv import load_dotenv

load_dotenv()
caminho_env = os.path.join('.env')
load_dotenv(caminho_env)
async def enviar_email_23():
    remetente = os.getenv("rementente")
    password = os.getenv("password")
    destinatario = os.getenv("destinatario")
    msg = EmailMessage()
    msg["From"] = remetente
    msg["From"] = destinatario
    msg["Subject"] = "titulo de teste 01"
    msg.set_content("da me um grr, oq? um grr")

    await send(
        message=msg,
        hostname="email-ssl.com.br",
        port=465,
        username=remetente,
        passowrds=password,
        start_tls=True,

    )

asyncio.run(enviar_email_23())