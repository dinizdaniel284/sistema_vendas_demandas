import os
import pandas as pd
from dotenv import load_dotenv
from pymongo import MongoClient

# Carrega variáveis do .env
load_dotenv()

# Conecta ao MongoDB
mongo_uri = os.getenv("MONGO_URI")
client = MongoClient(mongo_uri)

db = client["sistema_vendas"]
colecao_vendas = db["vendas"]

# Caminho para planilha
caminho_planilha = os.path.join("dados", "planilha.xlsx")

# Verifica se a planilha existe
if not os.path.exists(caminho_planilha):
    print("❌ Planilha não encontrada. Gere os dados primeiro.")
    exit()

# Lê os dados da planilha
df = pd.read_excel(caminho_planilha)

# Análise 1: Total de vendas
total_vendas = df['Quantidade'].sum()

# Análise 2: Produto mais vendido
produto_mais_vendido = df.groupby('Produto')['Quantidade'].sum().idxmax()

# Análise 3: Faturamento por produto
df['Faturamento'] = df['Quantidade'] * df['Preco_Unitario']
faturamento_por_produto = df.groupby('Produto')['Faturamento'].sum()

# Análise 4: Faturamento por data
faturamento_por_data = df.groupby('Data')['Faturamento'].sum()

# Salva os dados no MongoDB (remove antigos para evitar duplicatas)
colecao_vendas.delete_many({})
colecao_vendas.insert_many(df.to_dict(orient="records"))

# Cria diretório para relatórios
os.makedirs("relatorios", exist_ok=True)

# Gera o relatório em .txt
with open(os.path.join("relatorios", "relatorio_resumo.txt"), "w", encoding="utf-8") as f:
    f.write(f"📊 RELATÓRIO DE VENDAS\n\n")
    f.write(f"Total de vendas (itens): {total_vendas}\n")
    f.write(f"Produto mais vendido: {produto_mais_vendido}\n\n")
    f.write("Faturamento por Produto:\n")
    f.write(faturamento_por_produto.to_string())
    f.write("\n\nFaturamento por Data:\n")
    f.write(faturamento_por_data.to_string())

print("✅ Relatório gerado com sucesso em 'relatorios/relatorio_resumo.txt'")
print("✅ Dados armazenados no MongoDB com sucesso!")
