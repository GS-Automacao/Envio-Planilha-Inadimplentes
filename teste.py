from dotenv import load_dotenv
import os

load_dotenv()

caminho_env = os.path.join(os.path.dirname(__file__), '.env')
load_dotenv(dotenv_path=caminho_env)

variavel = os.getenv('diretorio_base')

print(variavel)
print(caminho_env)