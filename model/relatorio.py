from typing import List
from google.oauth2.credentials import Credentials
import os
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.errors import HttpError

#-> Função de gerar o relatorio na planilha do google.
#-> necessario que a pasta configs esteja no mesmo diretorio do executavel.
#-> as colunas do relatorio são fixas com>: a data da execução, nome do bot, df e tempo de execução.


def salva_relatorio(row: List[List]) -> None:
        SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
        SAMPLE_SPREADSHEET_ID = "15gGHm67_W5maIas-4_YPSzE6R5f_CNJGcza_BJFlNBk"
        SAMPLE_RANGE_NAME = "Página1!A{}:D1000000"

        creds = None
        if os.path.exists("configs/token.json"):
            creds = Credentials.from_authorized_user_file("configs/token.json", SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file("configs/client_secret.json", SCOPES)
                creds = flow.run_local_server(port=0)

            with open("configs/token.json", "w") as token:
                token.write(creds.to_json())

        try:
            service = build("sheets", "v4", credentials=creds)
            sheet = service.spreadsheets()

            result = sheet.values().get(
                spreadsheetId=SAMPLE_SPREADSHEET_ID,
                range=SAMPLE_RANGE_NAME.format(2)
            ).execute()
            values = result.get("values", [])

            idx = 2 + len(values)

            sheet.values().update(
                spreadsheetId=SAMPLE_SPREADSHEET_ID,
                range=SAMPLE_RANGE_NAME.format(idx),
                valueInputOption='USER_ENTERED',
                body={"values": row}
            ).execute()

            print(f"Relatório salvo na planilha Google Sheets na linha {idx}.")

        except HttpError as err:
            print(f"Erro ao acessar Google Sheets: {err}")

