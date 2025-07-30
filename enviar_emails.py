import os
from email.message import EmailMessage
import smtplib
import pandas as pd
import tempfile

#-> pega as informações para enviar o email no arquivo processar_pastas.py e envia o email.
#-> servidor correspodente ao do remetente (IMPORTANTE)

def enviar_email_com_df(para: str, assunto: str, corpo: str, df: pd.DataFrame, remetente: str, senha: str, nome_original: str, smtp=None): 
    msg = EmailMessage()
    msg['From'] = remetente
    msg['To'] = para
    msg['Subject'] = assunto
    msg.set_content(corpo)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        temp_path = tmp.name

    with pd.ExcelWriter(temp_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    with open(temp_path, 'rb') as f:
        msg.add_attachment(
            f.read(),
            maintype='application',
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename=nome_original
        )

    os.unlink(temp_path)

    if smtp is None:
        with smtplib.SMTP_SSL('email-ssl.com.br', 465) as server:
            server.login(remetente, senha)
            server.send_message(msg)
    else:
        smtp.send_message(msg)

    

  


    
