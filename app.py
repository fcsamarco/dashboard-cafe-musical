
import pandas as pd
import dash
from dash import html, dcc
import plotly.express as px

df_compras = pd.read_excel('itens.xlsx', sheet_name='compras da semana')
df_compras['Data'] = pd.to_datetime(df_compras['Data'])
df_compras['Ano'] = df_compras['Data'].dt.year
df_compras['Mes'] = df_compras['Data'].dt.month

df_custos = pd.read_excel('itens.xlsx', sheet_name='CUSTOS ')
df_custos['DATA'] = pd.to_datetime(df_custos['DATA'])
df_custos['Ano'] = df_custos['DATA'].dt.year
df_custos['Mes'] = df_custos['DATA'].dt.month

df_receb = pd.read_excel('itens.xlsx', sheet_name='Recebimentos')
df_receb['Data'] = pd.to_datetime(df_receb['Data'])
df_receb['Ano'] = df_receb['Data'].dt.year
df_receb['Mes'] = df_receb['Data'].dt.month

app = dash.Dash(__name__)

app.layout = html.Div([
    html.H1("Dashboard Financeiro - Café Musical"),

    dcc.Dropdown(
        id='ano_dropdown',
        options=[{'label': str(ano), 'value': ano} for ano in sorted(df_compras['Ano'].unique())],
        value=sorted(df_compras['Ano'].unique())[-1],
        clearable=False
    ),

    html.Div(id='graficos')
])

@app.callback(
    dash.dependencies.Output('graficos', 'children'),
    [dash.dependencies.Input('ano_dropdown', 'value')]
)
def atualizar_graficos(ano):
    dados = df_compras[df_compras['Ano'] == ano]
    custos = df_custos[df_custos['Ano'] == ano]
    receb = df_receb[df_receb['Ano'] == ano]

    total_compras = dados.groupby('Mes')['TOTAL'].sum().reset_index()
    total_custos = custos.groupby('Mes')['VALOR'].sum().reset_index()
    total_receb = receb.groupby('Mes')['VALOR'].sum().reset_index()

    fig_compras = px.bar(total_compras, x='Mes', y='TOTAL', title='Total Comprado por Mês')
    fig_custos = px.bar(total_custos, x='Mes', y='VALOR', title='Total Custos por Mês', color_discrete_sequence=['#E9C46A'])
    fig_receb = px.bar(total_receb, x='Mes', y='VALOR', title='Total Recebido por Mês', color_discrete_sequence=['#2A9D8F'])

    return [
        dcc.Graph(figure=fig_compras),
        dcc.Graph(figure=fig_custos),
        dcc.Graph(figure=fig_receb)
    ]

if __name__ == '__main__':
    app.run(debug=True)
