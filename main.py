import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF

def load_data(file_path):
    # Carregar a planilha com todas as variáveis na primeira linha
    df = pd.read_excel(file_path, header=0)
    return df

def process_data(df):
    # Garantir que DATAVENDA está no formato correto
    df['DATAVENDA'] = pd.to_datetime(df['DATAVENDA'], errors='coerce')

    # Filtrar dados para julho e agosto
    df_julho = df[df['DATAVENDA'].dt.month == 7]
    df_agosto = df[df['DATAVENDA'].dt.month == 8]

    # Contar o número de vendas por dia
    vendas_julho = df_julho.groupby(df_julho['DATAVENDA'].dt.date)['NUMVENDA'].count()
    vendas_agosto = df_agosto.groupby(df_agosto['DATAVENDA'].dt.date)['NUMVENDA'].count()

    # Calcular receita total por mês
    df['mes'] = df['DATAVENDA'].dt.to_period('M')
    receita_mensal = df.groupby('mes')['RECEITA'].sum()

    # Analisar estados
    estado_vendas = df['ESTADO'].value_counts(normalize=True) * 100

    # Encontrar o produto mais vendido
    produto_mais_vendido = df['TITULO'].value_counts().idxmax()
    sku_produto_mais_vendido = df[df['TITULO'] == produto_mais_vendido]['SKU'].values[0]
    qtd_produto_mais_vendido = df['TITULO'].value_counts().max()

    # Calcular receita gerada pelo produto mais vendido
    receita_produto_mais_vendido = df[df['TITULO'] == produto_mais_vendido]['RECEITA'].sum()

    return vendas_julho, vendas_agosto, receita_mensal, estado_vendas, produto_mais_vendido, sku_produto_mais_vendido, qtd_produto_mais_vendido, receita_produto_mais_vendido

def generate_reports(vendas_julho, vendas_agosto, receita_mensal, estado_vendas, produto_mais_vendido, sku_produto_mais_vendido, qtd_produto_mais_vendido, receita_produto_mais_vendido):
    # Plotar gráficos
    plt.figure(figsize=(14, 7))

    # Gráfico de vendas por dia
    plt.subplot(1, 2, 1)
    vendas_julho.plot(kind='bar', color='skyblue', label='Julho')
    vendas_agosto.plot(kind='bar', color='salmon', label='Agosto')
    plt.title('Número de Vendas por Dia')
    plt.xlabel('Dia')
    plt.ylabel('Número de Vendas')
    plt.legend()
    plt.grid(True)

    # Gráfico de receita por mês
    plt.subplot(1, 2, 2)
    receita_mensal.plot(kind='bar', color='lightgreen')
    plt.title('Receita por Mês')
    plt.xlabel('Mês')
    plt.ylabel('Receita (R$)')
    plt.grid(True)

    plt.tight_layout()
    plt.savefig('relatorio_graficos.png')
    plt.close()

    # Criar relatório PDF
    pdf = FPDF()
    pdf.add_page()

    # Adicionar título principal
    pdf.set_font("Arial", 'B', size=18)
    pdf.cell(0, 12, txt="Relatório de Vendas e Receita", ln=True, align='C')
    pdf.ln(20)  # Adicionar espaço após o título

    # Adicionar gráficos ao PDF
    pdf.image('relatorio_graficos.png', x=10, y=pdf.get_y(), w=190)
    pdf.ln(100)  # Adicionar espaço após os gráficos

    # Adicionar seção de vendas por dia
    pdf.set_font("Arial", 'B', size=14)
    pdf.cell(0, 8, txt="Número de Vendas por Dia (Julho e Agosto):", ln=True, align='C')
    pdf.ln(10)  # Adicionar espaço antes da tabela

    # Adicionar vendas de julho
    pdf.set_font("Arial", size=12)
    pdf.cell(0, 14, txt="Vendas de Julho Mercado Livre:", ln=True, align='L')
    pdf.cell(0, 10, txt="Data            |    Vendas", ln=True, align='L')
    pdf.cell(0, 10, txt="--------------------------------", ln=True, align='L')
    for date, count in vendas_julho.items():
        pdf.cell(0, 10, txt=f"{date} | {count} vendas", ln=True, align='L')
    pdf.ln(20)  # Adicionar espaço

    # Adicionar vendas de agosto
    pdf.cell(0, 10, txt="Vendas de Agosto Mercado Livre:", ln=True, align='L')
    pdf.cell(0, 10, txt="Data            |    Vendas", ln=True, align='L')
    pdf.cell(0, 10, txt="--------------------------------", ln=True, align='L')
    for date, count in vendas_agosto.items():
        pdf.cell(0, 10, txt=f"{date} | {count} vendas", ln=True, align='L')

    # Adicionar seção de receita
    pdf.ln(20)  # Adicionar espaço
    pdf.set_font("Arial", 'B', size=14)
    pdf.cell(0, 10, txt="Receita por Mês:", ln=True, align='C')
    pdf.ln(10)  # Adicionar espaço antes da tabela

    pdf.set_font("Arial", size=12)
    pdf.cell(0, 10, txt="Mês        | Receita (R$)", ln=True, align='L')
    pdf.cell(0, 10, txt="-----------------------", ln=True, align='L')
    for period, total in receita_mensal.items():
        pdf.cell(0, 10, txt=f"{period} | R$ {total:.2f}", ln=True, align='L')

    # Adicionar seção de estado
    pdf.ln(20)  # Adicionar espaço
    pdf.set_font("Arial", 'B', size=14)
    pdf.cell(0, 10, txt="Porcentagem de Vendas por Estado:", ln=True, align='C')
    pdf.ln(10)  # Adicionar espaço antes da tabela

    pdf.set_font("Arial", size=12)
    pdf.cell(0, 10, txt="Estado     | Percentual (%)", ln=True, align='L')
    pdf.cell(0, 10, txt="-----------------------", ln=True, align='L')
    for estado, percent in estado_vendas.items():
        pdf.cell(0, 10, txt=f"{estado} | {percent:.2f}%", ln=True, align='L')

    # Adicionar seção do produto mais vendido
    pdf.ln(20)  # Adicionar espaço
    pdf.set_font("Arial", 'B', size=14)
    pdf.cell(0, 10, txt="Produto Mais Vendido:", ln=True, align='C')
    pdf.ln(10)  # Adicionar espaço antes da tabela

    # Adicionar tabela para o produto mais vendido
    pdf.set_font("Arial", 'B', size=12)
    pdf.cell(100, 10, 'Produto', border=1, align='C')
    pdf.cell(90, 10, 'SKU', border=1, ln=True, align='C')

    pdf.set_font("Arial", size=10)
    pdf.cell(100, 10, produto_mais_vendido, border=1, align='C')
    pdf.cell(90, 10, sku_produto_mais_vendido, border=1, ln=True, align='C')

    # Adicionar seção detalhada sobre o produto mais vendido
    pdf.ln(20)  # Adicionar espaço
    pdf.set_font("Arial", 'B', size=14)
    pdf.cell(0, 10, txt="Detalhes do Produto Mais Vendido:", ln=True, align='C')
    pdf.ln(10)  # Adicionar espaço antes da tabela

    pdf.set_font("Arial", size=12)
    pdf.cell(0, 10, txt=f"Produto: {produto_mais_vendido}", ln=True, align='L')
    pdf.cell(0, 10, txt=f"SKU: {sku_produto_mais_vendido}", ln=True, align='L')
    pdf.cell(0, 10, txt=f"Quantidade Vendida: {qtd_produto_mais_vendido}", ln=True, align='L')
    pdf.cell(0, 10, txt=f"Receita Gerada deste produto: R$ {receita_produto_mais_vendido:.2f}", ln=True, align='L')

    # Salvar PDF
    pdf.output('relatorio_final.pdf')

    print("Relatório PDF 'relatorio_final.pdf' gerado com sucesso!")

def main():
    file_path = 'mlb.xlsx'

    # Carregar os dados
    df = load_data(file_path)

    # Processar os dados
    vendas_julho, vendas_agosto, receita_mensal, estado_vendas, produto_mais_vendido, sku_produto_mais_vendido, qtd_produto_mais_vendido, receita_produto_mais_vendido = process_data(df)

    # Gerar relatórios
    generate_reports(vendas_julho, vendas_agosto, receita_mensal, estado_vendas, produto_mais_vendido, sku_produto_mais_vendido, qtd_produto_mais_vendido, receita_produto_mais_vendido)

if __name__ == "__main__":
    main()