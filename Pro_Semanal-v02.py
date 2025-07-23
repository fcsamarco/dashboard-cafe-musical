import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from fpdf import FPDF
import os
import traceback

# Cores padrão
cores = ['#1f77b4', '#2ca02c', '#d62728']

# Dicionário para tradução dos meses
meses = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
}

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

class PDF(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', align='R')

def log_sucesso(mensagem):
    print(mensagem)
    with open("relatorio.log", "a", encoding="utf-8") as log:
        log.write(mensagem + "\n")

def log_erro(erro):
    print(f"Erro: {erro}")
    with open("relatorio_erro.log", "a", encoding="utf-8") as log:
        log.write(erro + "\n")

def gerar_relatorio_mensal(compras_df, custos_df, receb_df, ano, mes, output_dir='relatorios'):
    dados_mes = compras_df[(compras_df['Ano'] == ano) & (compras_df['Mes'] == mes)]
    custos_mes = custos_df[(custos_df['Ano'] == ano) & (custos_df['Mes'] == mes)]
    receb_mes = receb_df[(receb_df['Ano'] == ano) & (receb_df['Mes'] == mes)]

    if dados_mes.empty:
        print(f"Nenhum dado de compras para {mes}/{ano}")
        return

    os.makedirs(output_dir, exist_ok=True)
    pdf = PDF()
    pdf.add_page()

    nome_mes = meses.get(mes, f"Mês {mes}")

    # Título centralizado
    pdf.set_font("Arial", 'B', 20)
    pdf.cell(0, 10, f'Relatório - {nome_mes}', ln=True, align='C')

    # Logo centralizada abaixo
    logo_path = "logo_cafe_musical.png"
    if os.path.exists(logo_path):
        pdf.image(logo_path, x=(210 - 60) / 2, y=pdf.get_y() + 5, w=60)
        pdf.ln(50)  # espaço após a logo
    else:
        pdf.ln(20)

    total_mes = dados_mes['TOTAL'].sum()
    total_custos = custos_mes['VALOR'].sum()
    total_receb = receb_mes['VALOR'].sum()
    saldo = total_receb - (total_mes + total_custos)

    semanas = sorted(dados_mes['Semana'].unique())

    for idx, semana in enumerate(semanas, 1):
        pdf.add_page()
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, f"Semana {idx} de {nome_mes}", ln=True)

        dados_semana = dados_mes[dados_mes['Semana'] == semana]
        resumo_tipo = dados_semana.groupby('tipo').agg({'TOTAL': 'sum'}).reset_index()
        resumo_tipo['PERCENTUAL'] = resumo_tipo['TOTAL'] / resumo_tipo['TOTAL'].sum() * 100
        resumo_tipo = resumo_tipo.sort_values(by='PERCENTUAL', ascending=False)

        grafico_file = f"{output_dir}/grafico_compras_{ano}_{mes}_semana{idx}.png"
        salvar_grafico_barras(resumo_tipo, f'Compras - Semana {idx}', 'tipo', 'TOTAL', grafico_file)
        pdf.image(grafico_file, x=10, y=None, w=180)
        os.remove(grafico_file)

        pdf.ln(5)
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 8, 'Resumo:', ln=True)

        for _, row in resumo_tipo.iterrows():
            pdf.set_font("Arial", '', 12)
            pdf.write(5, f"{row['tipo']}: ")
            pdf.set_font("Arial", 'B', 12)
            pdf.write(5, f"{formatar_valor(row['TOTAL'])} (")
            pdf.set_text_color(0, 0, 255)
            pdf.write(5, f"{row['PERCENTUAL']:.1f}%")
            pdf.set_text_color(0, 0, 0)
            pdf.write(5, ")\n")

    for titulo, df, col_tipo, col_valor in [
        ('Custos', custos_mes, 'TIPO', 'VALOR'),
        ('Recebimentos', receb_mes, 'Fonte', 'VALOR')
    ]:
        if not df.empty:
            pdf.add_page()
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(0, 10, f'{titulo} - Consolidado do Mês', ln=True)

            resumo = df.groupby(col_tipo).agg({col_valor: 'sum'}).reset_index()
            resumo['PERCENTUAL'] = resumo[col_valor] / resumo[col_valor].sum() * 100
            resumo = resumo.sort_values(by='PERCENTUAL', ascending=False)

            grafico_file = f"{output_dir}/grafico_{titulo.lower()}_{ano}_{mes}.png"
            salvar_grafico_pizza(resumo, f'{titulo} - {nome_mes}', col_valor, col_tipo, grafico_file)
            pdf.image(grafico_file, x=30, y=None, w=150)
            os.remove(grafico_file)

            pdf.ln(5)
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(0, 8, 'Resumo:', ln=True)

            for _, row in resumo.iterrows():
                pdf.set_font("Arial", '', 12)
                pdf.write(5, f"{row[col_tipo]}: {formatar_valor(row[col_valor])} (")
                pdf.set_text_color(0, 0, 255)
                pdf.set_font("Arial", 'B', 12)
                pdf.write(5, f"{row['PERCENTUAL']:.1f}%")
                pdf.set_text_color(0, 0, 0)
                pdf.write(5, ")\n")

    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, 'Fechamento do Mês', ln=True)
    pdf.set_font("Arial", '', 12)

    for label, value, color in [
        ("Total Comprado: ", total_mes, (255, 0, 0)),
        ("Total Custos: ", total_custos, (255, 0, 0)),
        ("Total Recebido: ", total_receb, (0, 0, 255))
    ]:
        pdf.set_text_color(0, 0, 0)
        pdf.write(5, label)
        pdf.set_text_color(*color)
        pdf.set_font("Arial", 'B', 12)
        pdf.write(5, f"{formatar_valor(value)}\n")

    pdf.set_text_color(0, 0, 0)
    pdf.write(5, "Saldo do Fechamento: ")
    if saldo >= 0:
        pdf.set_text_color(0, 0, 255)
    else:
        pdf.set_text_color(255, 0, 0)
    pdf.set_font("Arial", 'B', 12)
    pdf.write(5, f"{formatar_valor(saldo)}\n")

    pdf_file = f"{output_dir}/relatorio_{ano}_{mes}.pdf"
    pdf.output(pdf_file)
    log_sucesso(f"Relatório gerado: {pdf_file}")

try:
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

    receb_df = pd.read_excel('itens.xlsx', sheet_name='Recebimentos')
    receb_df.columns = receb_df.columns.str.strip()
    receb_df['Data'] = pd.to_datetime(receb_df['Data'])
    receb_df['Ano'] = receb_df['Data'].dt.year
    receb_df['Mes'] = receb_df['Data'].dt.month

    for (ano, mes) in compras_df[['Ano', 'Mes']].drop_duplicates().sort_values(['Ano', 'Mes']).values:
        gerar_relatorio_mensal(compras_df, custos_df, receb_df, ano, mes)

except Exception as e:
    log_erro(traceback.format_exc())

input("Pressione Enter para sair...")
