import os
import time
from dotenv import load_dotenv
from datetime import datetime
from relatorio import salva_relatorio
from processar_pastas import processar_pastas
from atualizar_planilha import atualizar_planilha_geral
from validador import validador_horario
from enviar_emails import enviar_email_com_df

# Carrega credenciais
# Executa processo
#1) puxa a função para atualizar a planilha geral
#2) puxa a função para validação do horario
#3) processa as pastas, atualiza, ler e envia o email
#4) salva o relatorio de tempo de execução e o log de erro no mesmo diretorio do arquivo

load_dotenv()
caminho_env = os.path.join(os.path.dirname(__file__), '.env')

load_dotenv(dotenv_path=caminho_env)
atualizar_planilha_geral()
caminho_arquivo = os.getenv('caminho_base')
nome_aba = "Consolidado"
coluna = "U"

validador_horario(caminho_arquivo, nome_aba, coluna)

REMETENTE = os.getenv("usuario")
SENHA = os.getenv("senha")
DIRETORIO_BASE = os.getenv('DIRETORIO_BASE')
time_inicio = time.time()
processar_pastas(DIRETORIO_BASE, REMETENTE, SENHA)

time_fim = time.time()
tempo_total = time_fim - time_inicio
data = datetime.now().strftime("%d/%m/%Y")
row = [[data, "Envio Emails Inadimplencia", "", tempo_total]]
salva_relatorio(row)
print("Cleyton agradece pela execução dele :)")

