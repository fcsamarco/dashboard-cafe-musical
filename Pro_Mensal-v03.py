import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from fpdf import FPDF
import os
import traceback

# Cores padrão para gráficos de pizza
cor_comprado = '#FFA500'  # laranja
cor_custos = '#FFFF00'    # amarelo
cor_recebido = '#008000'  # verde
cor_saldo_pos = '#0000FF' # azul
cor_saldo_neg = '#FF0000' # vermelho

# Dicionário meses
meses = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
}

def formatar_valor(valor):
    return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

def salvar_grafico_pizza_resumo(valores, labels, cores, titulo, filename):
    if not any(pd.notna(valores)) or sum(valores) == 0:
        print(f"Aviso: gráfico '{titulo}' não gerado pois todos os valores são nulos ou zero.")
        return  # Pula a geração do gráfico

    plt.figure(figsize=(6,6))
    plt.pie(valores, labels=labels, autopct='%1.1f%%', startangle=90, colors=cores)
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

    # Capa
    pdf.set_font("Arial", 'B', 20)
    pdf.cell(0, 10, f'Relatório Anual - {ano}', ln=True, align='C')

    # Logo
    logo_path = "logo_cafe_musical.png"
    if os.path.exists(logo_path):
        pdf.image(logo_path, x=(210 - 60) / 2, y=pdf.get_y() + 5, w=60)
        pdf.ln(50)
    else:
        pdf.ln(20)

    total_compras = 0
    total_custos = 0
    total_receb = 0

    resumo_meses = []

    for mes in range(1, 13):
        dados_mes = compras_df[(compras_df['Ano'] == ano) & (compras_df['Mes'] == mes)]
        custos_mes = custos_df[(custos_df['Ano'] == ano) & (custos_df['Mes'] == mes)]
        receb_mes = receb_df[(receb_df['Ano'] == ano) & (receb_df['Mes'] == mes)]

        if dados_mes.empty and custos_mes.empty and receb_mes.empty:
            continue

        nome_mes = meses.get(mes, f"Mês {mes}")

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

        # Página do mês
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, f"{nome_mes}", ln=True)

        for label, value, color in [
            ("Total Comprado: ", total_mes, (255, 0, 0)),
            ("Total Custos: ", total_cust, (255, 0, 0)),
            ("Total Recebido: ", total_rec, (0, 0, 255))
        ]:
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Arial", '', 12)
            pdf.write(5, label)
            pdf.set_text_color(*color)
            pdf.set_font("Arial", 'B', 12)
            pdf.write(5, f"{formatar_valor(value)}\n")

        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", '', 12)
        pdf.write(5, "Saldo do Mês: ")
        if saldo >= 0:
            pdf.set_text_color(0, 0, 255)
        else:
            pdf.set_text_color(255, 0, 0)
        pdf.set_font("Arial", 'B', 12)
        pdf.write(5, f"{formatar_valor(saldo)}\n")

        # Gráfico de pizza do mês
        valores = [total_mes, total_cust, total_rec, abs(saldo)]
        cores = [cor_comprado, cor_custos, cor_recebido, cor_saldo_pos if saldo >=0 else cor_saldo_neg]
        labels = ['Comprado', 'Custos', 'Recebido', 'Saldo']
        grafico_file = f"{output_dir}/grafico_pizza_{ano}_{mes}.png"
        salvar_grafico_pizza_resumo(valores, labels, cores, f"{nome_mes} - Resumo", grafico_file)
        if os.path.exists(grafico_file):
            pdf.image(grafico_file, x=30, y=None, w=150)
            os.remove(grafico_file)

    # Resumo anual
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, f'Resumo Anual - {ano}', ln=True)

    for item in resumo_meses:
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 8, item['Mês'], ln=True)
        pdf.set_font("Arial", '', 12)

        for label, value, color in [
            ("Total Comprado: ", item['Compras'], (255, 0, 0)),
            ("Total Custos: ", item['Custos'], (255, 0, 0)),
            ("Total Recebido: ", item['Recebimentos'], (0, 0, 255)),
            ("Saldo: ", item['Saldo'], (0, 0, 255) if item['Saldo'] >=0 else (255, 0, 0))
        ]:
            pdf.set_text_color(0, 0, 0)
            pdf.write(5, label)
            pdf.set_text_color(*color)
            pdf.set_font("Arial", 'B', 12)
            pdf.write(5, f"{formatar_valor(value)}\n")

    # Fechamento anual
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, f'Fechamento do Ano - {ano}', ln=True)

    saldo_anual = total_receb - (total_compras + total_custos)

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
    pdf.write(5, "Saldo do Ano: ")
    if saldo_anual >= 0:
        pdf.set_text_color(0, 0, 255)
    else:
        pdf.set_text_color(255, 0, 0)
    pdf.set_font("Arial", 'B', 12)
    pdf.write(5, f"{formatar_valor(saldo_anual)}\n")

    # Gráfico de pizza anual
    valores = [total_compras, total_custos, total_receb, abs(saldo_anual)]
    cores = [cor_comprado, cor_custos, cor_recebido, cor_saldo_pos if saldo_anual >=0 else cor_saldo_neg]
    labels = ['Comprado', 'Custos', 'Recebido', 'Saldo']
    grafico_file = f"{output_dir}/grafico_pizza_anual_{ano}.png"
    salvar_grafico_pizza_resumo(valores, labels, cores, f"Resumo Anual {ano}", grafico_file)
    if os.path.exists(grafico_file):
        pdf.image(grafico_file, x=30, y=None, w=150)
        os.remove(grafico_file)

    pdf_file = f"{output_dir}/relatorio_anual_{ano}.pdf"
    pdf.output(pdf_file)
    log_sucesso(f"Relatório Anual gerado: {pdf_file}")

# Execução
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
