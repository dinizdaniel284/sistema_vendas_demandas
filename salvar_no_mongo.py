import os
from dotenv import load_dotenv
from pymongo import MongoClient
import pandas as pd

# Carrega variáveis do .env
load_dotenv()

# Conexão com o MongoDB
mongo_uri = os.getenv("MONGO_URI")
client = MongoClient(mongo_uri)
db = client["sistema_vendas"]
colecao_vendas = db["vendas"]

# Caminho da planilha
caminho_planilha = os.path.join("dados", "planilha.xlsx")

# Verifica se a planilha existe
if not os.path.exists(caminho_planilha):
    print("❌ Planilha não encontrada.")
    exit()

# Lê os dados da planilha
df = pd.read_excel(caminho_planilha)

# Converte DataFrame em dicionários
dados = df.to_dict(orient="records")

# Insere no MongoDB
if dados:
    colecao_vendas.insert_many(dados)
    print(f"✅ {len(dados)} registros inseridos no MongoDB!")
else:
    print("⚠️ Nenhum dado encontrado na planilha.")
