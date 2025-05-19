import tkinter as tk
from tkinter import messagebox, scrolledtext
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

import os
import pandas as pd
from pymongo import MongoClient
from dotenv import load_dotenv

load_dotenv()

# Configuração MongoDB
mongo_uri = os.getenv("MONGO_URI")
client = MongoClient(mongo_uri)
db = client["sistema_vendas"]
colecao_vendas = db["vendas"]

# Caminho planilha e diretórios
CAMINHO_PLANILHA = os.path.join("dados", "planilha.xlsx")
PASTA_RELATORIOS = "relatorios"
os.makedirs(PASTA_RELATORIOS, exist_ok=True)

class App(ttk.Window):
    def __init__(self):
        super().__init__(themename="superhero")  # tema moderno
        self.title("Sistema de Vendas - Dashboard")
        self.geometry("700x500")

        # Label topo
        self.label_titulo = ttk.Label(self, text="Sistema de Vendas e Demandas", font=("Helvetica", 18, "bold"))
        self.label_titulo.pack(pady=10)

        # Botões
        self.btn_gerar_relatorio = ttk.Button(self, text="📊 Gerar Relatório", command=self.gerar_relatorio)
        self.btn_gerar_relatorio.pack(pady=5, fill=X, padx=20)

        self.btn_salvar_mongo = ttk.Button(self, text="📤 Salvar Dados no MongoDB", command=self.salvar_dados_mongo)
        self.btn_salvar_mongo.pack(pady=5, fill=X, padx=20)

        self.btn_mostrar_relatorio = ttk.Button(self, text="📄 Mostrar Relatório", command=self.mostrar_relatorio)
        self.btn_mostrar_relatorio.pack(pady=5, fill=X, padx=20)

        # Caixa texto para relatório
        self.texto_relatorio = scrolledtext.ScrolledText(self, width=80, height=15)
        self.texto_relatorio.pack(pady=10, padx=20)

    def gerar_relatorio(self):
        if not os.path.exists(CAMINHO_PLANILHA):
            messagebox.showerror("Erro", "Planilha não encontrada. Gere os dados primeiro.")
            return

        df = pd.read_excel(CAMINHO_PLANILHA)

        total_vendas = df['Quantidade'].sum()
        produto_mais_vendido = df.groupby('Produto')['Quantidade'].sum().idxmax()
        df['Faturamento'] = df['Quantidade'] * df['Preco_Unitario']
        faturamento_por_produto = df.groupby('Produto')['Faturamento'].sum()
        faturamento_por_data = df.groupby('Data')['Faturamento'].sum()

        texto = (
            f"📊 RELATÓRIO DE VENDAS\n\n"
            f"Total de vendas (itens): {total_vendas}\n"
            f"Produto mais vendido: {produto_mais_vendido}\n\n"
            f"Faturamento por Produto:\n{faturamento_por_produto.to_string()}\n\n"
            f"Faturamento por Data:\n{faturamento_por_data.to_string()}"
        )

        # Salva em arquivo
        with open(os.path.join(PASTA_RELATORIOS, "relatorio_resumo.txt"), "w", encoding="utf-8") as f:
            f.write(texto)

        self.texto_relatorio.delete("1.0", tk.END)
        self.texto_relatorio.insert(tk.END, texto)
        messagebox.showinfo("Sucesso", "Relatório gerado e exibido com sucesso!")

    def salvar_dados_mongo(self):
        if not os.path.exists(CAMINHO_PLANILHA):
            messagebox.showerror("Erro", "Planilha não encontrada.")
            return

        df = pd.read_excel(CAMINHO_PLANILHA)
        registros = df.to_dict(orient="records")
        resultado = colecao_vendas.insert_many(registros)
        messagebox.showinfo("Sucesso", f"{len(resultado.inserted_ids)} registros inseridos no MongoDB!")

    def mostrar_relatorio(self):
        caminho = os.path.join(PASTA_RELATORIOS, "relatorio_resumo.txt")
        if not os.path.exists(caminho):
            messagebox.showwarning("Aviso", "Relatório não encontrado. Gere o relatório primeiro.")
            return

        with open(caminho, "r", encoding="utf-8") as f:
            conteudo = f.read()

        self.texto_relatorio.delete("1.0", tk.END)
        self.texto_relatorio.insert(tk.END, conteudo)


if __name__ == "__main__":
    app = App()
    app.mainloop()
