import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF

# Carregar os dados
df = pd.read_excel("itens.xlsx")
df.columns = df.columns.str.strip()

# Converter coluna de data
df["Data"] = pd.to_datetime(df["Data"])
df["Itens"] = df["Itens"].str.strip()

# Adicionar coluna de semana (domingo a sábado)
df["Semana"] = df["Data"] - pd.to_timedelta(df["Data"].dt.weekday + 1 % 7, unit='D')
df["Semana"] = df["Semana"].dt.date

# Agrupar os dados por semana
semanas = df["Semana"].sort_values().unique()

# Classe do PDF
class PDFRelatorio(FPDF):
    def header(self):
        self.set_font("Arial", "B", 14)
        self.cell(0, 10, "Relatório de Gastos Semanais", ln=True, align="C")
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Página {self.page_no()}", align="C")

# Criar PDF
pdf = PDFRelatorio()

# Processar cada semana
for semana in semanas:
    df_semana = df[df["Semana"] == semana]
    total_geral = df_semana["TOTAL"].sum()
    por_tipo = df_semana.groupby("tipo")["TOTAL"].sum().sort_values(ascending=False)
    por_item = df_semana.groupby("Itens")["TOTAL"].sum().sort_values(ascending=False)

    # Gráfico por tipo
    plt.figure(figsize=(10, 5))
    por_tipo.plot(kind='bar', color='coral')
    plt.title(f"Gastos por Tipo - Semana de {semana}")
    plt.xlabel("Tipo")
    plt.ylabel("Total (R$)")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    chart_path = f"grafico_semana_{semana}.png"
    plt.savefig(chart_path)
    plt.close()

    # Página do relatório
    pdf.add_page()
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, f"Semana de: {semana}", ln=True)

    # Gráfico
    pdf.image(chart_path, x=30, w=150)
    pdf.ln(5)

    # Total geral
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Total da Semana: R$ {total_geral:,.2f}", ln=True)

    # Resumo por tipo
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Resumo por Tipo:", ln=True)
    pdf.set_font("Arial", "", 11)
    for tipo, valor in por_tipo.items():
        percentual = (valor / total_geral) * 100 if total_geral else 0
        pdf.cell(0, 8, f"{tipo}: R$ {valor:,.2f} ({percentual:.1f}%)", ln=True)

    # Resumo por item
    pdf.ln(3)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Totais por Item:", ln=True)
    pdf.set_font("Arial", "", 10)
    for item, valor in por_item.items():
        pdf.multi_cell(0, 6, f"{item}: R$ {valor:,.2f}")

# Exportar PDF
pdf.output("relatorio_semanal.pdf")

print("✅ Relatório semanal gerado: relatorio_semanal.pdf")
