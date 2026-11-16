import os
import pandas as pd
import plotly.graph_objects as go
from dash import Dash, dcc, html, dash_table, Input, Output
import dash_bootstrap_components as dbc

# 1. LER O FICHEIRO EXCEL DA MESMA PASTA DO app.py
pasta_atual = os.path.dirname(os.path.abspath(__file__))

possiveis_ficheiros = [
    os.path.join(pasta_atual, 'Ramen.xlsx'),
    os.path.join(pasta_atual, 'Ramen.xslx'),
    os.path.join(pasta_atual, '11212025The-Ramen-Rater.xlsx')
]

caminho_final = next((p for p in possiveis_ficheiros if os.path.exists(p)), None)
if not caminho_final:
    raise FileNotFoundError(
        f"Ficheiro Excel não encontrado em {pasta_atual}. "
        "Garante que o ficheiro Excel está na raiz do repositório junto ao app.py."
    )

xls = pd.ExcelFile(caminho_final, engine='openpyxl')
sheet = 'The Big List' if 'The Big List' in xls.sheet_names else xls.sheet_names[0]
df = pd.read_excel(caminho_final, sheet_name=sheet, engine='openpyxl')

# 2. LIMPEZA DE DADOS
if 'T' in df.columns and df['T'].isna().all():
    df = df.drop(columns=['T'])

df['Stars'] = pd.to_numeric(df['Stars'], errors='coerce')
df = df.dropna(subset=['Stars', 'Country']).copy()

df['Country'] = df['Country'].astype(str).str.strip()
if 'Style' in df.columns:
    df['Style'] = df['Style'].fillna('Desconhecido').astype(str).str.strip()

paises_disponiveis = sorted(df['Country'].unique())

# 3. INICIALIZAR APLICAÇÃO DASH E EXPOR SERVER PARA O GUNICORN
app = Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])
server = app.server

# 4. LAYOUT
app.layout = dbc.Container([

    dbc.Row([
        dbc.Col([
            html.H1("The Ramen Rater — Dashboard", className="text-center text-primary mt-3 mb-1"),
            html.P(f"Base de dados: {len(df)} avaliações válidas", className="text-center text-muted mb-4")
        ], width=12)
    ]),

    dbc.Row([
        dbc.Col([
            html.Label("Filtrar por País (Country):", style={"fontWeight": "bold"}),
            dcc.Dropdown(
                id="filter-country",
                options=[{"label": c, "value": c} for c in paises_disponiveis],
                value=[],
                multi=True,
                placeholder="Selecione um ou mais países (vazio = todos)..."
            )
        ], md=6, className="mb-3"),

        dbc.Col([
            html.Label("Filtrar por Classificação (Stars):", style={"fontWeight": "bold"}),
            dcc.RangeSlider(
                id="filter-stars",
                min=0.0,
                max=5.0,
                step=0.25,
                value=[0.0, 5.0],
                marks={i: f"{i} ★" for i in range(6)},
                tooltip={"placement": "bottom", "always_visible": True}
            )
        ], md=6, className="mb-3")
    ], style={"backgroundColor": "#f8f9fa", "padding": "20px", "borderRadius": "10px", "marginBottom": "20px"}),

    dbc.Row([
        dbc.Col([
            dcc.Tabs(id="tabs", value="tab-graphics", children=[
                dcc.Tab(label="📊 Graphics", value="tab-graphics"),
                dcc.Tab(label="📋 Table", value="tab-table")
            ]),

            # TAB 1: GRAPHICS
            html.Div(id="container-graphics", className="mt-4", children=[
                dbc.Row([
                    dbc.Col([
                        dcc.Graph(id="graph-country", style={"height": "450px"})
                    ], md=8),
                    dbc.Col([
                        dcc.Graph(id="graph-style", style={"height": "450px"})
                    ], md=4)
                ]),
                dbc.Row([
                    dbc.Col([
                        dcc.Graph(id="graph-brands", style={"height": "400px"})
                    ], width=12)
                ], className="mt-3")
            ]),

            # TAB 2: TABLE
            html.Div(id="container-table", className="mt-4", style={"display": "none"}, children=[
                html.H5(id="table-count-title", className="mb-3"),
                dash_table.DataTable(
                    id="ramen-table",
                    columns=[{"name": i, "id": i} for i in df.columns],
                    page_size=15,
                    style_table={'overflowX': 'auto'},
                    style_header={'backgroundColor': '#2c3e50', 'color': 'white', 'fontWeight': 'bold'},
                    style_cell={'textAlign': 'left', 'padding': '10px'},
                    filter_action="native",
                    sort_action="native"
                )
            ])
        ], width=12)
    ])

], fluid=True)


# 5. CALLBACK 1: ALTERNAR VISIBILIDADE ENTRE TABS
@app.callback(
    [
        Output("container-graphics", "style"),
        Output("container-table", "style")
    ],
    Input("tabs", "value")
)
def toggle_tabs(active_tab):
    if active_tab == "tab-graphics":
        return {"display": "block"}, {"display": "none"}
    return {"display": "none"}, {"display": "block"}


# 6. CALLBACK 2: ATUALIZAR GRÁFICOS E TABELA COM OS FILTROS
@app.callback(
    [
        Output("graph-country", "figure"),
        Output("graph-style", "figure"),
        Output("graph-brands", "figure"),
        Output("ramen-table", "data"),
        Output("table-count-title", "children")
    ],
    [
        Input("filter-country", "value"),
        Input("filter-stars", "value")
    ]
)
def update_dashboard(selected_countries, stars_range):
    dff = df.copy()

    if selected_countries:
        dff = dff[dff['Country'].isin(selected_countries)]

    dff = dff[(dff['Stars'] >= stars_range[0]) & (dff['Stars'] <= stars_range[1])]

    if dff.empty:
        empty_fig = go.Figure()
        empty_fig.update_layout(title="Sem dados para os filtros selecionados")
        return empty_fig, empty_fig, empty_fig, [], "A mostrar 0 registos"

    # 1. Gráfico de Barras por País
    counts = dff['Country'].value_counts()
    fig_bar = go.Figure(data=[
        go.Bar(
            x=counts.index.tolist(),
            y=counts.values.tolist(),
            marker_color='#2c3e50'
        )
    ])
    fig_bar.update_layout(
        title="Contagem de Ramen por País",
        xaxis_title="País",
        yaxis_title="Quantidade",
        template="plotly_white",
        margin=dict(l=40, r=20, t=50, b=80)
    )

    # 2. Gráfico Circular por Estilo (Style)
    style_counts = dff['Style'].value_counts()
    fig_pie = go.Figure(data=[
        go.Pie(
            labels=style_counts.index.tolist(),
            values=style_counts.values.tolist(),
            hole=0.4
        )
    ])
    fig_pie.update_layout(
        title="Distribuição por Embalagem (Style)",
        template="plotly_white",
        margin=dict(l=20, r=20, t=50, b=20)
    )

    # 3. Top 10 Marcas Mais Avaliadas
    top_brands = dff['Brand'].value_counts().head(10).sort_values(ascending=True)
    fig_brands = go.Figure(data=[
        go.Bar(
            x=top_brands.values.tolist(),
            y=top_brands.index.tolist(),
            orientation='h',
            marker_color='#18bc9c'
        )
    ])
    fig_brands.update_layout(
        title="Top 10 Marcas Mais Avaliadas",
        xaxis_title="Nº de Reviews",
        yaxis_title="Marca",
        template="plotly_white",
        margin=dict(l=150, r=20, t=50, b=40)
    )

    return (
        fig_bar,
        fig_pie,
        fig_brands,
        dff.to_dict('records'),
        f"A mostrar {len(dff)} registos filtrados"
    )


# 7. EXECUTAR LOCALMENTE OU VIA PORTA DO RENDER
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8050))
    app.run(host="0.0.0.0", port=port, debug=False)