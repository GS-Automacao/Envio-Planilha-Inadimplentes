import pandas as pd

def ler_excel_em_dataframe(caminho_excel: str) -> pd.DataFrame:
    return pd.read_excel(caminho_excel, sheet_name=0)
