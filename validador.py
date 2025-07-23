import openpyxl
from datetime import datetime

def validador_horario(caminho_arquivo, nome_aba, coluna):
    try: 
        workbook = openpyxl.load_workbook(caminho_arquivo, data_only=True)
        sheet = workbook[nome_aba]
        data_atual = datetime.now().date()
        datas_invalidas = []

        for row in range(2, sheet.max_row + 1):
            try:
                celula = sheet[f"{coluna}{row}"]
                valor_celula = celula.value

                if valor_celula is None:
                    continue  
                if isinstance(valor_celula, datetime):
                    data_celula = valor_celula.date()
                elif isinstance(valor_celula, str):
                    
                    try:
                        data_celula = datetime.strptime(valor_celula.strip(), '%d/%m/%Y').date()
                    except ValueError:
                        try:
                            data_celula = datetime.strptime(valor_celula.strip(), '%Y-%m-%d').date()
                        except ValueError:
                            raise ValueError("Formato desconhecido")
                else:
                    raise TypeError("Tipo de dado não reconhecido")

                if data_celula != data_atual:
                    datas_invalidas.append((row, data_celula))

            except (ValueError, TypeError):
                datas_invalidas.append((row, "Formato inválido"))

        return datas_invalidas

    except FileNotFoundError:
        print(f"Erro: arquivo não encontrado em {caminho_arquivo}")
        return None
    except KeyError:
        print(f"Erro: Aba '{nome_aba}' não encontrada no arquivo.")
        return None
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")
        return None

caminho = r"Y:\DADOS\INADIMPLENCIA\BASE\Contas a Receber em aberto - GS.xlsx"
aba = "Consolidado"
coluna = "U"

datas_invalidas = validador_horario(caminho, aba, coluna)

if datas_invalidas is not None:
    if datas_invalidas:
        print("Datas inválidas encontradas:")
        for linha, data in datas_invalidas:
            print(f"Linha {linha}: {data}")
    else:
        print("Todas as datas na coluna correspondem à data atual.")