import pandas as pd
import os
from datetime import datetime

def gerar_relatorio():
    # Caminho do arquivo Excel
    caminho_planilha = os.path.join('dados', 'planilha.xlsx')

    # Verifica se o arquivo existe
    if not os.path.exists(caminho_planilha):
        print("❌ Arquivo planilha.xlsx não encontrado em 'dados/'.")
        return

    # Leitura da planilha
    df = pd.read_excel(caminho_planilha)

    # Verifica se as colunas esperadas existem
    colunas_esperadas = ['Produto', 'Quantidade', 'Preço Unitário']
    if not all(coluna in df.columns for coluna in colunas_esperadas):
        print("❌ A planilha não contém as colunas esperadas:", colunas_esperadas)
        return

    # Cálculos
    df['Total Item'] = df['Quantidade'] * df['Preço Unitário']
    total_itens = df['Quantidade'].sum()
    valor_total = df['Total Item'].sum()

    # Data e hora do relatório
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    # Conteúdo do relatório
    relatorio = (
        f"📝 RELATÓRIO DE VENDAS\n"
        f"Gerado em: {agora}\n"
        f"{'-'*40}\n"
        f"Total de itens vendidos: {total_itens}\n"
        f"Valor total arrecadado: R$ {valor_total:,.2f}\n"
        f"{'-'*40}\n\n"
        f"DETALHAMENTO POR PRODUTO:\n\n"
    )

    for index, linha in df.iterrows():
        relatorio += (
            f"{linha['Produto']}: {linha['Quantidade']} und. x "
            f"R$ {linha['Preço Unitário']:.2f} = R$ {linha['Total Item']:.2f}\n"
        )

    # Criar pasta relatorios se não existir
    pasta_relatorios = 'relatorios'
    os.makedirs(pasta_relatorios, exist_ok=True)

    # Caminho do arquivo de relatório
    caminho_relatorio = os.path.join(pasta_relatorios, 'relatorio_vendas.txt')

    # Salva o relatório em arquivo .txt
    with open(caminho_relatorio, 'w', encoding='utf-8') as arquivo:
        arquivo.write(relatorio)

    print(f"✅ Relatório gerado com sucesso em: {caminho_relatorio}")

# Executa a função
if __name__ == "__main__":
    gerar_relatorio()
