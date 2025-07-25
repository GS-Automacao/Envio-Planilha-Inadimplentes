import os
from datetime import datetime, date, timedelta
from tqdm import tqdm
import pandas as pd
from atualizar_planilha import atualizar_planilha_excel
from ler_planilha import ler_excel_em_dataframe
from enviar_emails import enviar_email_com_df
from dotenv import load_dotenv

#Explicação da função:
#-> Processar_pastas: chama a função para atualizar a pasta geral
#  de forma prioritaria e depois atualiza e envia as outras pastas, incluive a geral.
#-> E gera o log de erro geral.

load_dotenv()
def processar_pastas(diretorio_base: str, remetente: str, senha: str):
    erros = []  # armazena erros
    try:
        caminho_geral = os.path.join(diretorio_base, "GERAL", "Contas a Receber em aberto - Geral.xlsx")
        print(f"\nAtualizando planilha GERAL: {caminho_geral}")
        atualizar_planilha_excel(caminho_geral)
    except Exception as e:
        msg = f"Erro ao atualizar a planilha da GERAL: {e}"
        print(msg)
        erros.append({"Pasta": "GERAL", "Erro": str(e), "DataHora": datetime.now().strftime("%d/%m/%Y %H:%M:%S")})
            
    pastas = [p for p in os.listdir(diretorio_base) if os.path.isdir(os.path.join(diretorio_base, p)) and p.upper() != ""] #-> se quiser pular alguma pasta especifica, basta colocar entre as aspas.
       
    for nome_pasta in tqdm(pastas, desc="\nProcessando pastas"):    
        caminho_pasta = os.path.join(diretorio_base, nome_pasta)

        planilhas = [f for f in os.listdir(caminho_pasta) if f.lower().endswith(('.xlsx', '.xls'))]
        if not planilhas:
            print(f" Nenhuma planilha na pasta: {nome_pasta}")
            continue

        caminho_excel = os.path.join(caminho_pasta, planilhas[0])
        destinatario = os.getenv(nome_pasta.strip())

        if not destinatario:
            msg = f"Destinatário não encontrado para '{nome_pasta}'."
            print(f"\n {msg} Pulando.")
            erros.append({
                "Pasta": nome_pasta,
                "Erro": msg,
                "DataHora": datetime.now().strftime("%d/%m/%Y %H:%M:%S") #-> add essas colunas ao log de erro.
            })
            continue


        try:
            print(f"\nAtualizando: {caminho_excel}")
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
            # informações do email, destinatario, titulo, corpo acima, a planilha correspodente, remetente, senha e nome original do arquivo.
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

    # Salvar log de erros (se houver) -> gera o log de erros.
    if erros:
        df_erros = pd.DataFrame(erros)
        log_path = os.path.join(os.getcwd(), "log_erros.xlsx")
        df_erros.to_excel(log_path, index=False)
        print(f"\nLog de erros salvo em: {log_path}")
    else:
        print("\nTodas as pastas foram processadas sem erros.")
