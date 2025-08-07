import os
import time
import win32com.client
from dotenv import load_dotenv
import pythoncom

load_dotenv()
def atualizar_planilha_geral():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    plan_geral = os.getenv('caminho_geral')
    pythoncom.CoInitialize()

    try:
        plan_atualizada = excel.Workbooks.Open(os.path.abspath(plan_geral))
        plan_atualizada.RefreshAll()
        time.sleep(5) 
        plan_atualizada.Save()
        plan_atualizada.Close(SaveChanges=True)
    except Exception as e:
        print(f"erro ao atualizar a planilha geral: {e}")
    finally:
        excel.Quit()

def atualizar_planilha_excel(caminho_excel: str):
    """Atualiza uma planilha Excel via COM."""
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    pythoncom.CoInitialize()
    try:
        wb = excel.Workbooks.Open(os.path.abspath(caminho_excel))
        wb.RefreshAll()
        time.sleep(5)
        wb.Save()
        wb.Close(SaveChanges=True)
    except Exception as e:
        print(f"Erro ao atualizar a planilha: {caminho_excel}\n{e}")
    finally:
        excel.Quit()



