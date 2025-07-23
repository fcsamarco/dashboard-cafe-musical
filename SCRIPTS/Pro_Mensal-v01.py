import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import os

# Função para formatar valores no padrão BR
def formatar_valor(valor):
    return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

def salvar_grafico_barras(analises, filename):
    meses = [f"{int(row['mes']):02d}" for row in analises]
    categorias = ['compras', 'vendas', 'custos', 'lucro']
    valores = {cat: [row[cat] for row in analises] for cat in categorias}
    
    x = range(len(meses))
    width = 0.2

    plt.figure(figsize=(10, 6))
    
    for idx, cat in enumerate(categorias):
        plt.bar([p + idx * width for p in x], valores[cat], width=width, label=cat.capitalize())

    plt.xlabel('Mês')
    plt.ylabel('Valor (R$)')
    plt.title('Compras, Vendas, Custos e Lucro por Mês')
    plt.xticks([p + 1.5 * width for p in x], meses)
    plt.legend()
    plt.grid(axis='y', linestyle='--', alpha=0.6)
    plt.tight_layout()
    plt.savefig(filename, facecolor='white')
    plt.close()

def gerar_relatorio_anual(analises, output_file='relatorios/relatorio_anual.pdf'):
    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    total_lucro = sum(a['lucro'] for a in analises)
    total_compras = sum(a['compras'] for a in analises)
    total_vendas = sum(a['vendas'] for a in analises)
    total_custos = sum(a['custos'] for a in analises)

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, 'Relatório Anual', ln=True)

    # Gráfico
    grafico_file = 'relatorios/grafico_anual.png'
    salvar_grafico_barras(analises, grafico_file)
    pdf.image(grafico_file, x=10, y=None, w=180)
    os.remove(grafico_file)

    pdf.ln(5)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, 'Resumo:', ln=True)

    # Percentuais e resumo
    for a in analises:
        percentual = (a['compras'] + a['custos']) / (total_compras + total_custos) * 100
        mes_str = f"{int(a['mes']):02d}/{a['ano']}"
        valor = formatar_valor(a['compras'] + a['custos'])
        
        pdf.set_font("Arial", '', 12)
        pdf.write(5, f"{mes_str}: {valor} (")
        pdf.set_text_color(0, 0, 255)
        pdf.set_font("Arial", 'B', 12)
        pdf.write(5, f"{percentual:.1f}%")
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", '', 12)
        pdf.write(5, ")\n")

    # Observações automáticas
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, 'Observações e Apontamentos:', ln=True)
    pdf.set_font("Arial", '', 12)

    maior_gasto = max(analises, key=lambda x: x['compras'] + x['custos'])
    maior_lucro = max(analises, key=lambda x: x['lucro'])
    maior_venda = max(analises, key=lambda x: x['vendas'])
    maior_custo = max(analises, key=lambda x: x['custos'])

    pdf.multi_cell(0, 8,
        f"Mês com maior gasto: {int(maior_gasto['mes']):02d}/{maior_gasto['ano']} - {formatar_valor(maior_gasto['compras'] + maior_gasto['custos'])}\n"
        f"Mês com maior lucro: {int(maior_lucro['mes']):02d}/{maior_lucro['ano']} - {formatar_valor(maior_lucro['lucro'])}\n"
        f"Mês com maior venda: {int(maior_venda['mes']):02d}/{maior_venda['ano']} - {formatar_valor(maior_venda['vendas'])}\n"
        f"Mês com maior custo: {int(maior_custo['mes']):02d}/{maior_custo['ano']} - {formatar_valor(maior_custo['custos'])}\n"
    )

    # Fechamento
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, 'Fechamento Anual', ln=True)
    pdf.set_font("Arial", '', 12)
    pdf.multi_cell(0, 8,
        f"Total Compras: {formatar_valor(total_compras)}\n"
        f"Total Vendas: {formatar_valor(total_vendas)}\n"
        f"Total Custos: {formatar_valor(total_custos)}\n"
        f"Lucro Total: {formatar_valor(total_lucro)}"
    )

    pdf.output(output_file)
    print(f"Relatório anual gerado: {output_file}")

# === MAIN ===
compras = pd.read_excel('itens.xlsx', sheet_name='compras da semana')
custos = pd.read_excel('itens.xlsx', sheet_name='CUSTOS ')

try:
    vendas = pd.read_excel('itens.xlsx', sheet_name='VENDAS ')
    vendas.columns = vendas.columns.str.strip()
    vendas['Data'] = pd.to_datetime(vendas['Data'])
    vendas['Ano'] = vendas['Data'].dt.year
    vendas['Mes'] = vendas['Data'].dt.month
    vendas['TOTAL'] = pd.to_numeric(vendas['TOTAL'], errors='coerce')
    vendas_disponivel = True
except:
    vendas = pd.DataFrame()
    vendas_disponivel = False

for df in [compras, custos]:
    df.columns = df.columns.str.strip()

compras['Data'] = pd.to_datetime(compras['Data'])
compras['Ano'] = compras['Data'].dt.year
compras['Mes'] = compras['Data'].dt.month

custos['DATA'] = pd.to_datetime(custos['DATA'])
custos['Ano'] = custos['DATA'].dt.year
custos['Mes'] = custos['DATA'].dt.month
custos['VALOR'] = pd.to_numeric(custos['VALOR'], errors='coerce')

analises = []
for (ano, mes) in compras[['Ano', 'Mes']].drop_duplicates().sort_values(['Ano', 'Mes']).values:
    total_compras = compras[(compras['Ano'] == ano) & (compras['Mes'] == mes)]['TOTAL'].sum()
    total_custos = custos[(custos['Ano'] == ano) & (custos['Mes'] == mes)]['VALOR'].sum()
    if vendas_disponivel:
        total_vendas = vendas[(vendas['Ano'] == ano) & (vendas['Mes'] == mes)]['TOTAL'].sum()
    else:
        total_vendas = 0.0
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
