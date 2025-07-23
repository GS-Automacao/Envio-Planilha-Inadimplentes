import os
import time
from dotenv import load_dotenv
from datetime import datetime
from relatorio import salva_relatorio
from processar_pastas import processar_pastas
from validador import validador_horario


# Carrega credenciais
load_dotenv()
REMETENTE = os.getenv("usuario")
SENHA = os.getenv("senha")

# Executa processo
caminho_arquivo = r"Y:\DADOS\INADIMPLENCIA\BASE\Contas a Receber em aberto - GS.xlsx"
nome_aba = "Consolidado"
coluna = "U"
validador_horario(caminho_arquivo, nome_aba, coluna)
DIRETORIO_BASE = r"Y:\DADOS\INADIMPLENCIA"
time_inicio = time.time()
processar_pastas(DIRETORIO_BASE, REMETENTE, SENHA)

time_fim = time.time()
tempo_total = time_fim - time_inicio
data = datetime.now().strftime("%d/%m/%Y")
row = [[data, "Envio Emails Inadimplencia", "", tempo_total]]
salva_relatorio(row)
