import pandas as pd
from fpdf import FPDF
import os

def gerar_relatorio_anual(analises, output_file='relatorios/relatorio_anual.pdf'):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, 'Relatório Anual', ln=True)

    total_lucro = 0

    for analise in analises:
        pdf.add_page()
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, f"{analise['mes']}/{analise['ano']}", ln=True)

        pdf.set_font("Arial", '', 12)
        pdf.multi_cell(0, 8, f"Compras: R$ {analise['compras']:.2f}\n"
                             f"Vendas: R$ {analise['vendas']:.2f}\n"
                             f"Custos: R$ {analise['custos']:.2f}\n"
                             f"Lucro: R$ {analise['lucro']:.2f}")

        total_lucro += analise['lucro']

    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, 'Fechamento Anual', ln=True)
    pdf.set_font("Arial", '', 12)
    pdf.cell(0, 10, f"Lucro Total: R$ {total_lucro:.2f}", ln=True)

    pdf.output(output_file)
    print(f"Relatório anual gerado: {output_file}")

compras = pd.read_excel('itens.xlsx', sheet_name='compras da semana')
custos = pd.read_excel('itens.xlsx', sheet_name='CUSTOS ')
vendas = pd.read_excel('itens.xlsx', sheet_name='VENDAS ')

for df in [compras, custos, vendas]:
    df.columns = df.columns.str.strip()

compras['Data'] = pd.to_datetime(compras['Data'])
compras['Ano'] = compras['Data'].dt.year
compras['Mes'] = compras['Data'].dt.month

custos['DATA'] = pd.to_datetime(custos['DATA'])
custos['Ano'] = custos['DATA'].dt.year
custos['Mes'] = custos['DATA'].dt.month
custos['VALOR'] = pd.to_numeric(custos['VALOR'], errors='coerce')

if 'Data' in vendas.columns:
    vendas['Data'] = pd.to_datetime(vendas['Data'])
    vendas['Ano'] = vendas['Data'].dt.year
    vendas['Mes'] = vendas['Data'].dt.month
    vendas['TOTAL'] = pd.to_numeric(vendas['TOTAL'], errors='coerce')

analises = []
for (ano, mes) in compras[['Ano', 'Mes']].drop_duplicates().sort_values(['Ano', 'Mes']).values:
    total_compras = compras[(compras['Ano'] == ano) & (compras['Mes'] == mes)]['TOTAL'].sum()
    total_vendas = vendas[(vendas['Ano'] == ano) & (vendas['Mes'] == mes)]['TOTAL'].sum()
    total_custos = custos[(custos['Ano'] == ano) & (custos['Mes'] == mes)]['VALOR'].sum()
    lucro = total_vendas - (total_compras + total_custos)

    analises.append({
        'ano': ano,
        'mes': mes,
        'compras': total_compras,
        'vendas': total_vendas,
        'custos': total_custos,
        'lucro': lucro
    })

gerar_relatorio_anual(analises)

input("Pressione Enter para sair...")