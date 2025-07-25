import os
from email.message import EmailMessage
import smtplib
import pandas as pd
from datetime import time, datetime
from relatorio import salva_relatorio
import time

#-> pega as informações para enviar o email no arquivo processar_pastas.py e envia o email.
#-> servidor correspodente ao do remetente (IMPORTANTE)

def enviar_email_com_df(para: str, assunto: str, corpo: str, df: pd.DataFrame, remetente: str, senha: str, nome_original: str):

    msg = EmailMessage()
    msg['From'] = remetente
    msg['To'] = para
    msg['Subject'] = assunto
    msg.set_content(corpo)

    temp_file = "temp_envio.xlsx"
    df.to_excel(temp_file, index=False)

    with open(temp_file, 'rb') as file:
        msg.add_attachment(
            file.read(),
            maintype='application',
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename=nome_original
        )

    with smtplib.SMTP_SSL('email-ssl.com.br', 465) as smtp:
        smtp.login(remetente, senha)
        smtp.send_message(msg)
        print(f"E-mail enviado para {para} com o anexo '{nome_original}'.")

    os.remove(temp_file)
    

  


    
