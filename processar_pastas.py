import os
from datetime import datetime, date, timedelta
from tqdm import tqdm
import pandas as pd
from atualizar_planilha import atualizar_planilha_excel
from ler_planilha import ler_excel_em_dataframe
from enviar_emails import enviar_email_com_df



def processar_pastas(diretorio_base: str, remetente: str, senha: str):
    erros = []  # Lista para armazenar os erros

    pastas = [p for p in os.listdir(diretorio_base) if os.path.isdir(os.path.join(diretorio_base, p))]

    for nome_pasta in tqdm(pastas, desc="Processando pastas"):
        caminho_pasta = os.path.join(diretorio_base, nome_pasta)

        planilhas = [f for f in os.listdir(caminho_pasta) if f.lower().endswith(('.xlsx', '.xls'))]
        if not planilhas:
            print(f" Nenhuma planilha na pasta: {nome_pasta}")
            continue

        caminho_excel = os.path.join(caminho_pasta, planilhas[0])
        destinatario = os.getenv(nome_pasta.strip())

        if not destinatario:
            msg = f"Destinatário não encontrado para '{nome_pasta}'."
            print(f" {msg} Pulando.")
            erros.append({
                "Pasta": nome_pasta,
                "Erro": msg,
                "DataHora": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            })
            continue


        try:
            print(f"Atualizando: {caminho_excel}")
            atualizar_planilha_excel(caminho_excel)
            
            print(f"Lendo: {caminho_excel}")
            df = ler_excel_em_dataframe(caminho_excel)
            um_dia = timedelta(days=1)
            data_atual = datetime.now()
            data_anterior = data_atual - um_dia
            data_anterior = data_anterior.strftime("%d/%m/%Y")
            data_hoje = date.today().strftime("%d/%m/%Y")
            
            
            corpo_email = f"""
        Bom dia,


            Segue em anexo os dados atualizados da inadimplência! 
            
            São considerados contas a receber até o vencimento em {data_anterior}.

            Este relatório é gerado diariamente para acompanhamento dos valores em aberto por cliente e data de vencimento.

            
            
        Qualquer dúvida, entrar em contato com (85)99825-2426 -  Helpdesk GS
            """

            enviar_email_com_df(
                para=destinatario,
                assunto=f"Inadimplência - {nome_pasta} - {data_hoje}",
                corpo = corpo_email,
                df=df,
                remetente=remetente,
                senha=senha,
                nome_original=os.path.basename(caminho_excel)
            )

           

        except Exception as e:
            msg = f"Erro ao processar {nome_pasta}: {e}"
            print(msg)
            erros.append({
                "Pasta": nome_pasta,
                "Erro": str(e),
                "DataHora": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                
            })

    # Salvar log de erros (se houver)
    if erros:
        df_erros = pd.DataFrame(erros)
        log_path = os.path.join(os.getcwd(), "log_erros.xlsx")
        df_erros.to_excel(log_path, index=False)
        print(f"\nLog de erros salvo em: {log_path}")
    else:
        print("\nTodas as pastas foram processadas sem erros.")
