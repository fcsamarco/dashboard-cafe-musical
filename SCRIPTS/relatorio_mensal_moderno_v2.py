import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from fpdf import FPDF
import os

cores = ['#1f77b4', '#2ca02c', '#d62728']

def salvar_grafico(dados, titulo, nome_coluna, valor_coluna, filename):
    plt.figure(figsize=(8, 5))
    sns.barplot(x=dados[nome_coluna], y=dados[valor_coluna], palette=cores)
    plt.title(titulo)
    plt.xlabel(nome_coluna)
    plt.ylabel(valor_coluna)
    plt.xticks(rotation=45, ha='right')
    plt.grid(True, linestyle='--', alpha=0.6)
    plt.tight_layout()
    plt.savefig(filename, facecolor='white')
    plt.close()

def gerar_relatorio_mensal(compras_df, ano, mes, output_dir='relatorios'):
    dados_mes = compras_df[(compras_df['Ano'] == ano) & (compras_df['Mes'] == mes)]
    if dados_mes.empty:
        print(f"Nenhum dado para {mes}/{ano}")
        return

    os.makedirs(output_dir, exist_ok=True)
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, f'Relatório - {mes}/{ano}', ln=True)

    total_mes = dados_mes['TOTAL'].sum()
    pdf.set_font("Arial", '', 12)
    pdf.multi_cell(0, 8, f"Total Comprado: R$ {total_mes:.2f}")

    resumo_tipo = dados_mes.groupby('tipo').agg({'TOTAL': 'sum'}).reset_index()
    resumo_tipo['PERCENTUAL'] = resumo_tipo['TOTAL'] / resumo_tipo['TOTAL'].sum() * 100

    grafico_file = f"{output_dir}/grafico_{ano}_{mes}.png"
    salvar_grafico(resumo_tipo, f'Compras por Categoria - {mes}/{ano}', 'tipo', 'TOTAL', grafico_file)
    pdf.image(grafico_file, x=10, y=None, w=180)
    os.remove(grafico_file)

    pdf.ln(5)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, 'Resumo:', ln=True)

    for _, row in resumo_tipo.iterrows():
        pdf.set_font("Arial", '', 12)
        texto = f"{row['tipo']}: R$ {row['TOTAL']:.2f} ("
        pdf.write(5, texto)
        pdf.set_text_color(0, 0, 255)  # azul
        pdf.set_font("Arial", 'B', 12)
        pdf.write(5, f"{row['PERCENTUAL']:.1f}%")
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", '', 12)
        pdf.write(5, ")
")

    pdf_file = f"{output_dir}/relatorio_{ano}_{mes}.pdf"
    pdf.output(pdf_file)
    print(f"Relatório gerado: {pdf_file}")

compras_df = pd.read_excel('itens.xlsx', sheet_name='compras da semana')
compras_df.columns = compras_df.columns.str.strip()
compras_df['Data'] = pd.to_datetime(compras_df['Data'])
compras_df['Semana'] = compras_df['Data'].dt.isocalendar().week
compras_df['Ano'] = compras_df['Data'].dt.year
compras_df['Mes'] = compras_df['Data'].dt.month

for (ano, mes) in compras_df[['Ano', 'Mes']].drop_duplicates().sort_values(['Ano', 'Mes']).values:
    gerar_relatorio_mensal(compras_df, ano, mes)

input("Pressione Enter para sair...")