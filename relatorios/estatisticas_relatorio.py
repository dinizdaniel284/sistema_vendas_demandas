import pandas as pd
import os

def gerar_estatisticas():
    # Caminho da planilha
    caminho_planilha = os.path.join('dados', 'planilha.xlsx')

    # Verifica se a planilha existe
    if not os.path.exists(caminho_planilha):
        print("❌ Planilha não encontrada!")
        return

    # Carregar planilha
    df = pd.read_excel(caminho_planilha)

    # Criar nova coluna de valor total por item
    df['Valor Total'] = df['Quantidade'] * df['Preço Unitário']

    # Total de vendas
    total_vendas = df['Valor Total'].sum()

    # Produto mais vendido (maior quantidade)
    produto_mais_vendido = df.loc[df['Quantidade'].idxmax(), 'Produto']

    # Produto mais lucrativo
    produto_mais_lucrativo = df.loc[df['Valor Total'].idxmax(), 'Produto']

    # Número total de itens vendidos
    total_itens = df['Quantidade'].sum()

    # Produto com menor quantidade vendida
    produto_menos_vendido = df.loc[df['Quantidade'].idxmin(), 'Produto']

    # Exibir no terminal
    print("📊 RELATÓRIO DE ESTATÍSTICAS")
    print(f"💵 Total de Vendas: R$ {total_vendas:,.2f}")
    print(f"📦 Produto Mais Vendido: {produto_mais_vendido}")
    print(f"💰 Produto Mais Lucrativo: {produto_mais_lucrativo}")
    print(f"🧾 Total de Itens Vendidos: {total_itens}")
    print(f"📉 Produto com Menor Saída: {produto_menos_vendido}")

    # Salvar relatório em .txt
    caminho_relatorio = os.path.join('relatorios', 'relatorio_estatisticas.txt')
    with open(caminho_relatorio, 'w', encoding='utf-8') as f:
        f.write("📊 RELATÓRIO DE ESTATÍSTICAS\n")
        f.write(f"💵 Total de Vendas: R$ {total_vendas:,.2f}\n")
        f.write(f"📦 Produto Mais Vendido: {produto_mais_vendido}\n")
        f.write(f"💰 Produto Mais Lucrativo: {produto_mais_lucrativo}\n")
        f.write(f"🧾 Total de Itens Vendidos: {total_itens}\n")
        f.write(f"📉 Produto com Menor Saída: {produto_menos_vendido}\n")

    print(f"\n📝 Relatório salvo em: {caminho_relatorio}")

if __name__ == '__main__':
    gerar_estatisticas()
