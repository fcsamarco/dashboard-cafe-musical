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
    plt.pie(dados[valor_coluna], labels=dados[nome_coluna], autopct='%1.1f%%', startangle=90)
    plt.title(titulo)
    plt.tight_layout()
    plt.savefig(filename, facecolor='white')
    plt.close()

class PDF(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(0, 0, 0)
        self.cell(0, 10, f'Página {self.page_no()}', align='R')

def log_sucesso(mensagem):
    print(mensagem)
    with open("relatorio.log", "a", encoding="utf-8") as log:
        log.write(mensagem + "\n")

def log_erro(erro):
    print(f"Erro: {erro}")
    with open("relatorio_erro.log", "a", encoding="utf-8") as log:
        log.write(erro + "\n")

def gerar_relatorio_periodo(compras_df, custos_df, receb_df, data_inicial, data_final, output_dir='relatorios'):
    compras_df['Semana'] = compras_df['Data'].dt.to_period('W-SUN')
    custos_df['Semana'] = custos_df['DATA'].dt.to_period('W-SUN')
    receb_df['Semana'] = receb_df['Data'].dt.to_period('W-SUN')

    dados_periodo = compras_df[(compras_df['Data'] >= data_inicial) & (compras_df['Data'] <= data_final)]
    custos_periodo = custos_df[(custos_df['DATA'] >= data_inicial) & (custos_df['DATA'] <= data_final)]
    receb_periodo = receb_df[(receb_df['Data'] >= data_inicial) & (receb_df['Data'] <= data_final)]

    if dados_periodo.empty:
        print("Nenhum dado de compras no período informado.")
        return

    os.makedirs(output_dir, exist_ok=True)
    pdf = PDF()
    pdf.add_page()

    pdf.set_font("Arial", 'B', 20)
    pdf.cell(0, 10, f'Relatório - {data_inicial.strftime("%d/%m/%Y")} a {data_final.strftime("%d/%m/%Y")}', ln=True, align='C')

    logo_path = "logo_cafe_musical.png"
    if os.path.exists(logo_path):
        pdf.image(logo_path, x=(210 - 60) / 2, y=pdf.get_y() + 5, w=60)
        pdf.ln(50)
    else:
        pdf.ln(20)

    total_compras = dados_periodo['TOTAL'].sum()
    total_custos = custos_periodo['VALOR'].sum()
    total_receb = receb_periodo['VALOR'].sum()
    saldo = total_receb - (total_compras + total_custos)

    semanas = sorted(dados_periodo['Semana'].unique())

    for idx, semana_period in enumerate(semanas, 1):
        data_inicio_semana = semana_period.start_time
        data_fim_semana = semana_period.end_time

        semana_data = dados_periodo[(dados_periodo['Data'] >= data_inicio_semana) & (dados_periodo['Data'] <= data_fim_semana)]
        custos_semana = custos_df[(custos_df['DATA'] >= data_inicio_semana) & (custos_df['DATA'] <= data_fim_semana)]
        receb_semana = receb_df[(receb_df['Data'] >= data_inicio_semana) & (receb_df['Data'] <= data_fim_semana)]

        pdf.add_page()
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, f"Semana {idx} ({data_inicio_semana.strftime('%d/%m/%Y')} a {data_fim_semana.strftime('%d/%m/%Y')})", ln=True)

        resumo_tipo = semana_data.groupby('tipo').agg({'TOTAL': 'sum'}).reset_index()
        if not resumo_tipo.empty:
            resumo_tipo['PERCENTUAL'] = resumo_tipo['TOTAL'] / resumo_tipo['TOTAL'].sum() * 100
            resumo_tipo = resumo_tipo.sort_values(by='PERCENTUAL', ascending=False)

            grafico_file = f"{output_dir}/grafico_compras_semana{idx}.png"
            salvar_grafico_barras(resumo_tipo, f'Compras - Semana {idx}', 'tipo', 'TOTAL', grafico_file)
            pdf.image(grafico_file, x=10, y=None, w=180)
            os.remove(grafico_file)

            total_semana = semana_data['TOTAL'].sum()

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

            pdf.set_font("Arial", 'B', 12)
            pdf.ln(3)
            pdf.cell(0, 8, f"Total da Semana: {formatar_valor(total_semana)}", ln=True)

        if not custos_semana.empty:
            resumo_custos = custos_semana.groupby('TIPO').agg({'VALOR': 'sum'}).reset_index()
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(0, 8, 'Custos da Semana:', ln=True)
            for _, row in resumo_custos.iterrows():
                pdf.set_font("Arial", '', 12)
                pdf.cell(0, 8, f"- {row['TIPO']}: {formatar_valor(row['VALOR'])}", ln=True)

        if not receb_semana.empty:
            resumo_receb = receb_semana.groupby('Fonte').agg({'VALOR': 'sum'}).reset_index()
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(0, 8, 'Recebimentos da Semana:', ln=True)
            for _, row in resumo_receb.iterrows():
                pdf.set_font("Arial", '', 12)
                pdf.cell(0, 8, f"- {row['Fonte']}: {formatar_valor(row['VALOR'])}", ln=True)

    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, 'Fechamento do Período', ln=True)
    pdf.set_font("Arial", '', 12)

    for label, value, color in [
        ("Total Comprado: ", total_compras, (255, 0, 0)),
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

    resumo_final = pd.DataFrame({
        'Categoria': ['Compras', 'Custos', 'Recebido', 'Saldo'],
        'Valor': [total_compras, total_custos, total_receb, saldo]
    })

    grafico_final = f"{output_dir}/grafico_fechamento.png"
    salvar_grafico_pizza(resumo_final, 'Fechamento do Período', 'Valor', 'Categoria', grafico_final)
    pdf.image(grafico_final, x=30, y=None, w=150)
    os.remove(grafico_final)

    pdf_file = f"{output_dir}/relatorio_{data_inicial.strftime('%d%m')}_{data_final.strftime('%d%m')}.pdf"
    pdf.output(pdf_file)
    log_sucesso(f"Relatório gerado: {pdf_file}")

# === EXECUÇÃO ===
try:
    compras_df = pd.read_excel('itens.xlsx', sheet_name='compras da semana')
    compras_df.columns = compras_df.columns.str.strip()
    compras_df['Data'] = pd.to_datetime(compras_df['Data'])

    custos_df = pd.read_excel('itens.xlsx', sheet_name='CUSTOS ')
    custos_df.columns = custos_df.columns.str.strip()
    custos_df['DATA'] = pd.to_datetime(custos_df['DATA'])

    receb_df = pd.read_excel('itens.xlsx', sheet_name='Recebimentos')
    receb_df.columns = receb_df.columns.str.strip()
    receb_df['Data'] = pd.to_datetime(receb_df['Data'])

    data_inicio = pd.to_datetime('2025-04-24')
    data_fim = pd.to_datetime('2025-05-17')

    gerar_relatorio_periodo(compras_df, custos_df, receb_df, data_inicio, data_fim)

except Exception as e:
    log_erro(traceback.format_exc())

input("Pressione Enter para sair...")
