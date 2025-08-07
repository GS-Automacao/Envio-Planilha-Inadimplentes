import pandas as pd

#-> le a planilha excel de cada arquivo

def ler_excel_em_dataframe(caminho_excel: str) -> pd.DataFrame:
    return pd.read_excel(caminho_excel, sheet_name=0)
