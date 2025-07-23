import os
import time
import win32com.client

def atualizar_planilha_excel(caminho_excel: str):
    """Atualiza uma planilha Excel via COM."""
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(os.path.abspath(caminho_excel))
    wb.RefreshAll()
    time.sleep(5)
    wb.Save()
    wb.Close()
    excel.Quit()



