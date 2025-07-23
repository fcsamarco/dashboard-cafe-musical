import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF

# Configurar fontes do matplotlib para UTF-8
plt.rcParams['axes.unicode_minus'] = False

# Carregar dados
df = pd.read_excel("itens.xlsx")
df.columns = df.columns.str.strip()
df["Itens"] = df["Itens"].str.strip()
df["Data"] = pd.to_datetime(df["Data"])

# Criar coluna de semana (domingo a sábado)
df["Semana"] = df["Data"] - pd.to_timedelta((df["Data"].dt.weekday + 1) % 7, unit="D")
df["Semana"] = df["Semana"].dt.date
semanas = df["Semana"].sort_values().unique()

# Classe para relatório principal
class PDFRelatorio(FPDF):
    def header(self):
        self.set_font("Arial", "B", 14)
        self.cell(0, 10, "Relatório de Gastos Semanais", ln=True, align="C")
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Página {self.page_no()}", align="C")

# Classe para totais por item
class PDFItens(FPDF):
    def header(self):
        self.set_font("Arial", "B", 14)
        self.cell(0, 10, "Totais por Item - Semanal", ln=True, align="C")
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Página {self.page_no()}", align="C")

# Criar PDFs
pdf = PDFRelatorio()
pdf_itens = PDFItens()

# Processar cada semana
for semana in semanas:
    df_semana = df[df["Semana"] == semana]
    total_geral = df_semana["TOTAL"].sum()
    por_tipo = df_semana.groupby("tipo")["TOTAL"].sum().sort_values(ascending=False)
    por_item = df_semana.groupby("Itens")["TOTAL"].sum().sort_values(ascending=False)

    # Gráfico com porcentagem sobre cada barra
    plt.figure(figsize=(10, 6))
    bars = por_tipo.plot(kind="bar", color="skyblue")
    plt.title(f"Gastos por Tipo - Semana de {semana}")
    plt.xlabel("Tipo")
    plt.ylabel("Total (R$)")
    plt.xticks(rotation=45, ha="right")

    # Adicionar rótulos com valores e porcentagens
    for idx, valor in enumerate(por_tipo):
        percentual = (valor / total_geral) * 100 if total_geral else 0
        plt.text(idx, valor + (total_geral * 0.01), f"R$ {valor:,.0f}\n({percentual:.1f}%)", 
                 ha='center', va='bottom', fontsize=8, color='black')

    plt.tight_layout()
    chart_path = f"grafico_semana_{semana}.png"
    plt.savefig(chart_path)
    plt.close()

    # Adicionar página ao PDF principal
    pdf.add_page()
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, f"Semana de: {semana}", ln=True)

    pdf.image(chart_path, x=30, w=150)
    pdf.ln(5)
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Total da Semana: R$ {total_geral:,.2f}", ln=True)

    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Resumo por Tipo:", ln=True)
    pdf.set_font("Arial", "", 11)

    # Resumo com porcentagem azul e negrito
    for tipo, valor in por_tipo.items():
        percentual = (valor / total_geral) * 100 if total_geral else 0
        texto = f"{tipo}: R$ {valor:,.2f} ("
        pdf.set_text_color(0, 0, 0)
        pdf.write(8, texto)
        pdf.set_text_color(0, 0, 255)
        pdf.set_font("Arial", "B", 11)
        pdf.write(8, f"{percentual:.1f}%")
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", "", 11)
        pdf.write(8, ")\n")

    # Criar página no relatório de itens
    pdf_itens.add_page()
    pdf_itens.set_font("Arial", "B", 12)
    pdf_itens.cell(0, 10, f"Semana de: {semana}", ln=True)
    pdf_itens.set_font("Arial", "", 10)
    for item, valor in por_item.items():
        pdf_itens.multi_cell(0, 6, f"{item}: R$ {valor:,.2f}")

# Salvar PDFs
pdf.output("relatorio_semanal.pdf")
pdf_itens.output("totais_por_item_semanal.pdf")

print("✅ Relatórios gerados:")
print("📄 relatorio_semanal.pdf")
print("📄 totais_por_item_semanal.pdf")
