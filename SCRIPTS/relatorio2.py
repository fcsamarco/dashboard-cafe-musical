import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF

# Carregar a planilha
df = pd.read_excel("itens.xlsx")
df.columns = df.columns.str.strip()
df["Itens"] = df["Itens"].str.strip()

# Agrupar dados
gasto_por_tipo = df.groupby("tipo")["TOTAL"].sum().sort_values(ascending=False)
gasto_por_item = df.groupby("Itens")["TOTAL"].sum().sort_values(ascending=False)
total_geral = df["TOTAL"].sum()

# Gerar gráfico de barras verticais
plt.figure(figsize=(10, 6))
gasto_por_tipo.plot(kind='bar', color='skyblue')
plt.title("Gastos por Tipo")
plt.xlabel("Tipo")
plt.ylabel("Total (R$)")
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.savefig("grafico_gastos.png")
plt.close()

# Classe do PDF
class PDFRelatorio(FPDF):
    def header(self):
        self.set_font("Arial", "B", 16)
        self.cell(0, 10, "Relatório de Gastos por Tipo", ln=True, align="C")
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Página {self.page_no()}", align="C")

# Criar o PDF
pdf = PDFRelatorio()
pdf.add_page()

# Inserir gráfico
pdf.image("grafico_gastos.png", x=30, w=150)
pdf.ln(10)

# Total geral
pdf.set_font("Arial", "", 12)
pdf.cell(0, 10, f"Total Geral de Gastos: R$ {total_geral:,.2f}", ln=True)

# Resumo por tipo com porcentagem
pdf.set_font("Arial", "B", 12)
pdf.cell(0, 10, "Resumo por Tipo (com %):", ln=True)
pdf.set_font("Arial", "", 11)
for tipo, valor in gasto_por_tipo.items():
    percentual = (valor / total_geral) * 100
    pdf.cell(0, 8, f"{tipo}: R$ {valor:,.2f} ({percentual:.1f}%)", ln=True)

# Totais por item
pdf.ln(5)
pdf.set_font("Arial", "B", 12)
pdf.cell(0, 10, "Totais por Item:", ln=True)
pdf.set_font("Arial", "", 10)
for item, valor in gasto_por_item.items():
    pdf.multi_cell(0, 6, f"{item}: R$ {valor:,.2f}")

# Salvar PDF
pdf.output("relatorio_gastos.pdf")

print("✅ Relatório gerado com sucesso: relatorio_gastos.pdf")
