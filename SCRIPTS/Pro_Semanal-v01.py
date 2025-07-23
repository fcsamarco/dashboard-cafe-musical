import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from fpdf import FPDF
import os

# Cores padrão
cores = ['#1f77b4', '#2ca02c', '#d62728']

def formatar_valor(valor):
    return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

def salvar_grafico_barras(dados, titulo, nome_coluna, valor_coluna, filename):
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

def salvar_grafico_pizza(dados, titulo, valor_coluna, nome_coluna, filename):
    plt.figure(figsize=(6, 6))
    plt.pie(dados[valor_coluna], labels=dados[nome_coluna], autopct='%1.1f%%', startangle=90, colors=cores)
    plt.title(titulo)
    plt.tight_layout()
    plt.savefig(filename, facecolor='white')
    plt.close()

def gerar_relatorio_mensal(compras_df, custos_df, ano, mes, output_dir='relatorios'):
    dados_mes = compras_df[(compras_df['Ano'] == ano) & (compras_df['Mes'] == mes)]
    custos_mes = custos_df[(custos_df['Ano'] == ano) & (custos_df['Mes'] == mes)]

    if dados_mes.empty:
        print(f"Nenhum dado de compras para {mes}/{ano}")
        return

    os.makedirs(output_dir, exist_ok=True)
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, f'Relatório - {mes}/{ano}', ln=True)

    total_mes = dados_mes['TOTAL'].sum()
    total_custos = custos_mes['VALOR'].sum()
    total_geral = total_mes + total_custos

    pdf.set_font("Arial", '', 12)
    pdf.multi_cell(0, 8, 
        f"Total Comprado: {formatar_valor(total_mes)}\n"
        f"Total Custos: {formatar_valor(total_custos)}\n"
        f"Total Geral: {formatar_valor(total_geral)}"
    )

    semanas = dados_mes['Semana'].unique()
    for semana in sorted(semanas):
        pdf.add_page()
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, f"Semana {semana} (Domingo a Sábado)", ln=True)

        dados_semana = dados_mes[dados_mes['Semana'] == semana]
        resumo_tipo = dados_semana.groupby('tipo').agg({'TOTAL': 'sum'}).reset_index()
        resumo_tipo['PERCENTUAL'] = resumo_tipo['TOTAL'] / resumo_tipo['TOTAL'].sum() * 100
        resumo_tipo = resumo_tipo.sort_values(by='PERCENTUAL', ascending=False)

        grafico_file = f"{output_dir}/grafico_compras_{ano}_{mes}_semana{semana}.png"
        salvar_grafico_barras(resumo_tipo, f'Compras - Semana {semana}', 'tipo', 'TOTAL', grafico_file)
        pdf.image(grafico_file, x=10, y=None, w=180)
        os.remove(grafico_file)

        pdf.ln(5)
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 8, 'Resumo:', ln=True)

        for _, row in resumo_tipo.iterrows():
            pdf.set_font("Arial", '', 12)
            valor_formatado = formatar_valor(row['TOTAL'])
            texto = f"{row['tipo']}: {valor_formatado} ("
            pdf.write(5, texto)
            pdf.set_text_color(0, 0, 255)
            pdf.set_font("Arial", 'B', 12)
            pdf.write(5, f"{row['PERCENTUAL']:.1f}%")
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Arial", '', 12)
            pdf.write(5, ")\n")

    if not custos_mes.empty:
        pdf.add_page()
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, 'Custos - Consolidado do Mês', ln=True)

        resumo_custos = custos_mes.groupby('TIPO').agg({'VALOR': 'sum'}).reset_index()
        resumo_custos['PERCENTUAL'] = resumo_custos['VALOR'] / resumo_custos['VALOR'].sum() * 100
        resumo_custos = resumo_custos.sort_values(by='PERCENTUAL', ascending=False)

        grafico_file = f"{output_dir}/grafico_custos_{ano}_{mes}.png"
        salvar_grafico_pizza(resumo_custos, f'Custos - {mes}/{ano}', 'VALOR', 'TIPO', grafico_file)
        pdf.image(grafico_file, x=30, y=None, w=150)
        os.remove(grafico_file)

        pdf.ln(5)
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 8, 'Resumo:', ln=True)

        for _, row in resumo_custos.iterrows():
            pdf.set_font("Arial", '', 12)
            valor_formatado = formatar_valor(row['VALOR'])
            texto = f"{row['TIPO']}: {valor_formatado} ("
            pdf.write(5, texto)
            pdf.set_text_color(0, 0, 255)
            pdf.set_font("Arial", 'B', 12)
            pdf.write(5, f"{row['PERCENTUAL']:.1f}%")
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Arial", '', 12)
            pdf.write(5, ")\n")

    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, 'Fechamento do Mês', ln=True)
    pdf.set_font("Arial", '', 12)
    pdf.multi_cell(0, 8, 
        f"Total Comprado: {formatar_valor(total_mes)}\n"
        f"Total Custos: {formatar_valor(total_custos)}\n"
        f"Total Geral: {formatar_valor(total_geral)}"
    )

    pdf_file = f"{output_dir}/relatorio_{ano}_{mes}.pdf"
    pdf.output(pdf_file)
    print(f"Relatório gerado: {pdf_file}")

# === MAIN ===
compras_df = pd.read_excel('itens.xlsx', sheet_name='compras da semana')
compras_df.columns = compras_df.columns.str.strip()
compras_df['Data'] = pd.to_datetime(compras_df['Data'])
compras_df['Semana'] = compras_df['Data'].dt.to_period('W-SUN').apply(lambda r: r.start_time.isocalendar()[1])
compras_df['Ano'] = compras_df['Data'].dt.year
compras_df['Mes'] = compras_df['Data'].dt.month

custos_df = pd.read_excel('itens.xlsx', sheet_name='CUSTOS ')
custos_df.columns = custos_df.columns.str.strip()
custos_df['DATA'] = pd.to_datetime(custos_df['DATA'])
custos_df['Ano'] = custos_df['DATA'].dt.year
custos_df['Mes'] = custos_df['DATA'].dt.month

for (ano, mes) in compras_df[['Ano', 'Mes']].drop_duplicates().sort_values(['Ano', 'Mes']).values:
    gerar_relatorio_mensal(compras_df, custos_df, ano, mes)

input("Pressione Enter para sair...")
