import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF

# Carregar a planilha
df = pd.read_excel("itens.xlsx")
df.columns = df.columns.str.strip()

# Corrigir nomes das colunas
df["Itens"] = df["Itens"].str.strip()

# Agrupar dados
gasto_por_tipo = df.groupby("tipo")["TOTAL"].sum().sort_values(ascending=False)
gasto_por_item = df.groupby("Itens")["TOTAL"].sum().sort_values(ascending=False)
total_geral = df["TOTAL"].sum()

# Gerar gráfico de pizza
fig, ax = plt.subplots(figsize=(8, 8))
colors = plt.cm.Pastel1.colors
ax.pie(
    gasto_por_tipo,
    labels=gasto_por_tipo.index,
    autopct="%1.1f%%",
    startangle=90,
    colors=colors
)
ax.set_title("Distribuição de Gastos por Tipo", fontsize=14)
plt.savefig("grafico_gastos.png", bbox_inches="tight")
plt.close()

# Criar PDF
class PDFRelatorio(FPDF):
    def header(self):
        self.set_font("Arial", "B", 16)
        self.cell(0, 10, "Relatório de Gastos por Tipo", ln=True, align="C")
        self.ln(10)
    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Página {self.page_no()}", align="C")

pdf = PDFRelatorio()
pdf.add_page()

# Gráfico
pdf.image("grafico_gastos.png", x=30, w=150)
pdf.ln(10)

# Total geral
pdf.set_font("Arial", "", 12)
pdf.cell(0, 10, f"Total Geral de Gastos: R$ {total_geral:,.2f}", ln=True)

# Resumo por tipo
pdf.set_font("Arial", "B", 12)
pdf.cell(0, 10, "Resumo por Tipo:", ln=True)
pdf.set_font("Arial", "", 11)
for tipo, valor in gasto_por_tipo.items():
    pdf.cell(0, 8, f"{tipo}: R$ {valor:,.2f}", ln=True)

# Resumo por item
pdf.ln(5)
pdf.set_font("Arial", "B", 12)
pdf.cell(0, 10, "Totais por Item:", ln=True)
pdf.set_font("Arial", "", 10)
for item, valor in gasto_por_item.items():
    pdf.multi_cell(0, 6, f"{item}: R$ {valor:,.2f}")

# Salvar PDF
pdf.output("relatorio_gastos.pdf")

print("✅ PDF gerado com sucesso: relatorio_gastos.pdf")
