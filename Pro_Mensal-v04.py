import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from fpdf import FPDF
import os
import traceback

# Cores atualizadas para gráficos de barras
cor_comprado = '#F4A261'  # laranja suave
cor_custos = '#E9C46A'    # amarelo queimado
cor_recebido = '#2A9D8F'  # verde esmeralda
cor_saldo_pos = '#264653' # azul petróleo
cor_saldo_neg = '#E76F51' # vermelho terracota

# Dicionário meses
meses = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
}

def formatar_valor(valor):
    return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

def salvar_grafico_barras_resumo(valores, labels, cores, titulo, filename):
    plt.figure(figsize=(8, 5))
    bars = plt.bar(labels, valores, color=cores)
    plt.title(titulo)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.tight_layout()
    for bar, val in zip(bars, valores):
        plt.text(bar.get_x() + bar.get_width()/2, bar.get_height(), formatar_valor(val),
                 ha='center', va='bottom', fontsize=9, fontweight='bold')
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
    with open("relatorio_anual.log", "a", encoding="utf-8") as log:
        log.write(mensagem + "\n")

def log_erro(erro):
    print(f"Erro: {erro}")
    with open("relatorio_anual_erro.log", "a", encoding="utf-8") as log:
        log.write(erro + "\n")

def gerar_relatorio_anual(compras_df, custos_df, receb_df, ano, output_dir='relatorios'):
    os.makedirs(output_dir, exist_ok=True)
    pdf = PDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 20)
    pdf.cell(0, 10, f'Relatório Anual - {ano}', ln=True, align='C')

    logo_path = "logo_cafe_musical.png"
    if os.path.exists(logo_path):
        pdf.image(logo_path, x=(210 - 60) / 2, y=pdf.get_y() + 5, w=60)
        pdf.ln(50)
    else:
        pdf.ln(20)

    total_compras, total_custos, total_receb = 0, 0, 0
    resumo_meses = []
    comparativo_mensal = []

    for mes in range(1, 13):
        dados_mes = compras_df[(compras_df['Ano'] == ano) & (compras_df['Mes'] == mes)]
        custos_mes = custos_df[(custos_df['Ano'] == ano) & (custos_df['Mes'] == mes)]
        receb_mes = receb_df[(receb_df['Ano'] == ano) & (receb_df['Mes'] == mes)]

        if dados_mes.empty and custos_mes.empty and receb_mes.empty:
            continue

        nome_mes = meses[mes]
        total_mes = dados_mes['TOTAL'].sum()
        total_cust = custos_mes['VALOR'].sum()
        total_rec = receb_mes['VALOR'].sum()
        saldo = total_rec - (total_mes + total_cust)

        resumo_meses.append({
            'Mês': nome_mes,
            'Compras': total_mes,
            'Custos': total_cust,
            'Recebimentos': total_rec,
            'Saldo': saldo
        })

        total_compras += total_mes
        total_custos += total_cust
        total_receb += total_rec

        pdf.add_page()
        pdf.set_fill_color(240, 240, 240)
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 8, f"{nome_mes}", ln=True, fill=True)

        for label, value, color in [
            ("Total Comprado: ", total_mes, (244, 162, 97)),
            ("Total Custos: ", total_cust, (233, 196, 106)),
            ("Total Recebido: ", total_rec, (42, 157, 143))
        ]:
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Arial", '', 12)
            pdf.write(5, label)
            pdf.set_text_color(*color)
            pdf.set_font("Arial", 'B', 12)
            pdf.write(5, f"{formatar_valor(value)}\n")

        pdf.set_text_color(0, 0, 0)
        pdf.write(5, "Saldo do Mês: ")
        cor_saldo = (38, 70, 83) if saldo >= 0 else (231, 111, 81)
        pdf.set_text_color(*cor_saldo)
        pdf.set_font("Arial", 'B', 12)
        pdf.write(5, f"{formatar_valor(saldo)}\n")

        valores = [total_mes, total_cust, total_rec, abs(saldo)]
        cores = [cor_comprado, cor_custos, cor_recebido, cor_saldo_pos if saldo >=0 else cor_saldo_neg]
        labels = ['Comprado', 'Custos', 'Recebido', 'Saldo']
        grafico_file = f"{output_dir}/grafico_barras_{ano}_{mes}.png"
        salvar_grafico_barras_resumo(valores, labels, cores, f"{nome_mes} - Resumo", grafico_file)
        if os.path.exists(grafico_file):
            pdf.image(grafico_file, x=15, y=None, w=180)
            os.remove(grafico_file)

        comentario = f"{nome_mes}: Saldo {'positivo' if saldo >=0 else 'negativo'} de {formatar_valor(saldo)}."
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", 'I', 11)
        pdf.multi_cell(0, 6, comentario)

    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, f'Resumo Anual - {ano}', ln=True)

    for i, item in enumerate(resumo_meses):
        pdf.set_font("Arial", 'B', 13)
        pdf.cell(0, 8, item['Mês'], ln=True)
        pdf.set_font("Arial", '', 12)
        pdf.write(5, f"Total Comprado: {formatar_valor(item['Compras'])}\n")
        pdf.write(5, f"Total Custos: {formatar_valor(item['Custos'])}\n")
        pdf.write(5, f"Total Recebido: {formatar_valor(item['Recebimentos'])}\n")
        pdf.write(5, f"Saldo: {formatar_valor(item['Saldo'])}\n")

        if i > 0:
            delta = item['Saldo'] - resumo_meses[i - 1]['Saldo']
            perc = (delta / abs(resumo_meses[i - 1]['Saldo'])) * 100 if resumo_meses[i - 1]['Saldo'] != 0 else 0
            pdf.set_text_color(100, 100, 100)
            texto = f"Crescimento mensal: {formatar_valor(delta)} ({perc:.1f}%)\n"
            pdf.set_font("Arial", 'I', 11)
            pdf.write(5, texto)
            pdf.set_text_color(0, 0, 0)

    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, f'Fechamento do Ano - {ano}', ln=True)

    saldo_anual = total_receb - (total_compras + total_custos)

    for label, value, color in [
        ("Total Comprado: ", total_compras, (244, 162, 97)),
        ("Total Custos: ", total_custos, (233, 196, 106)),
        ("Total Recebido: ", total_receb, (42, 157, 143))
    ]:
        pdf.set_text_color(0, 0, 0)
        pdf.write(5, label)
        pdf.set_text_color(*color)
        pdf.set_font("Arial", 'B', 12)
        pdf.write(5, f"{formatar_valor(value)}\n")

    pdf.set_text_color(0, 0, 0)
    pdf.write(5, "Saldo do Ano: ")
    pdf.set_text_color(*(38, 70, 83) if saldo_anual >= 0 else (231, 111, 81))
    pdf.set_font("Arial", 'B', 12)
    pdf.write(5, f"{formatar_valor(saldo_anual)}\n")

    valores = [total_compras, total_custos, total_receb, abs(saldo_anual)]
    cores = [cor_comprado, cor_custos, cor_recebido, cor_saldo_pos if saldo_anual >=0 else cor_saldo_neg]
    labels = ['Comprado', 'Custos', 'Recebido', 'Saldo']
    grafico_file = f"{output_dir}/grafico_barras_anual_{ano}.png"
    salvar_grafico_barras_resumo(valores, labels, cores, f"Resumo Anual {ano}", grafico_file)
    if os.path.exists(grafico_file):
        pdf.image(grafico_file, x=15, y=None, w=180)
        os.remove(grafico_file)

    pdf.output(f"{output_dir}/relatorio_anual_{ano}.pdf")
    log_sucesso(f"Relatório Anual gerado: relatorio_anual_{ano}.pdf")

try:
    compras_df = pd.read_excel('itens.xlsx', sheet_name='compras da semana')
    compras_df.columns = compras_df.columns.str.strip()
    compras_df['Data'] = pd.to_datetime(compras_df['Data'])
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

    anos = compras_df['Ano'].unique()
    for ano in sorted(anos):
        gerar_relatorio_anual(compras_df, custos_df, receb_df, ano)

except Exception as e:
    log_erro(traceback.format_exc())

input("Pressione Enter para sair...")
