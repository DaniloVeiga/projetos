import dash
import subprocess
import threading
import webbrowser

from dash import html, Input, Output, dcc
from dash.exceptions import PreventUpdate

import subprocess
import os


CARD_STYLE = {
    "border": "1px solid #D5D8DC",
    "borderRadius": "10px",
    "backgroundColor": "white",
    "padding": "20px",
    "display": "flex",
    "flexDirection": "column",
    "justifyContent": "center",
    "alignItems": "center",
    "fontSize": "18px",
    "fontWeight": "bold",
}


DASHBOARDS = [
    {
        "titulo": "Visualização dos Projetos",
        "descricao": "Projeção do andamento dos Projetos.",
        "imagem": "003.png",
        "id": "btn-projetos",
        "arquivo": r"\\VWTBJBSHFSCO.sa.vwg\Pecas_e_Acessorios_Vinhedo\Dashboard P&A\dashboard_v7\Projetos\Projetos.exe",
        "mensagem": "Dashboard de Projetos iniciado com sucesso.",
    },
    {
        "titulo": "Status Ilustrações",
        "descricao": "Demonstração dos Status das Ilustrações.",
        "imagem": "003.png",
        "id": "btn-ilustracoes",
        "arquivo": r"\\VWTBJBSHFSCO.sa.vwg\Pecas_e_Acessorios_Vinhedo\Dashboard P&A\dashboard_v7\Ilustracoes\Ilustracoes.exe",
        "mensagem": "Dashboard de Ilustrações iniciado com sucesso.",
    },
    {
        "titulo": "Lições aprendidas",
        "descricao": "Acervo Documental dos Procedimentos.",
        "imagem": "002.png",
        "id": "btn-licoes-aprendidas",
        "arquivo": "licoes_aprendidas.py",
        "mensagem": "Dashboard Em Desenvolvimento.",
    },
    {
        "titulo": "Status Documentações",
        "descricao": "Demonstração dos Status das Documentações.",
        "imagem": "002.png",
        "id": "btn-documentacoes",
        "arquivo": "documentacoes.py",
        "mensagem": "Dashboard Em Desenvolvimento.",
    },
    {
        "titulo": "Tickets",
        "descricao": "Demonstração dos Status dos Tickets.",
        "imagem": "003.png",
        "id": "btn-tickets",
        "arquivo": r"\\VWTBJBSHFSCO.sa.vwg\Pecas_e_Acessorios_Vinhedo\Dashboard P&A\dashboard_v7\Tickets\Tickets.exe",
        "mensagem": "Dashboard de Tickets iniciado com sucesso.",
    },
]

#---------------------------
CARD_STYLE2 = {
    "border": "1px solid #D5D8DC",
    "borderRadius": "10px",
    "backgroundColor": "white",
    "padding": "20px",
    "display": "flex",
    "flexDirection": "column",
    "justifyContent": "center",
    "alignItems": "center",
    "fontSize": "18px",
    "fontWeight": "bold",
}


DASHBOARDS2 = [
    {
        "titulo": "Desenvolvido por.:",
        "descricao": "Colaboradores.:",
        "imagem": "004.png",
        "Colaborador 1":"Breno Trautwein Neto.: breno.neto@volkswagen-groupservices.com",
        "Colaborador 2":"Danilo Barroso da Veiga.: dveiga@zeentech.com.br",
        "Colaborador 3":"Kayky Pires de Paula.:  extern.kayky.pires@volkswagen.com.br", 
        "Colaborador 4":"Leandro Molonha.: leandro.molonha@volkswagen.com.br",       
        "Colaborador 5":"Marcelo Henrique Costa.: marcelo.costa5@volkswagen.com.br",
        "imagem_roda_pe":"006.png",
    },
    {
        "titulo": "Atualizado por.:",
                "descricao": "Colaboradores.:",
                "imagem": "005.png",
                "Colaborador 1":"Breno Trautwein Neto.: breno.neto@volkswagen-groupservices.com",
                "Colaborador 2":"Danilo Barroso da Veiga.: dveiga@zeentech.com.br",
                "Colaborador 3":"Kayky Pires de Paula.:  extern.kayky.pires@volkswagen.com.br", 
                "Colaborador 4":"Leandro Molonha.: leandro.molonha@volkswagen.com.br",       
                "Colaborador 5":"Marcelo Henrique Costa.: marcelo.costa5@volkswagen.com.br",
                "imagem_roda_pe":"006.png",
    },
]
#------------------------------------------

def criar_card(info):
    textos = {"btn-licoes-aprendidas":"Em desenvolvimento",
                              "btn-documentacoes":"Em desenvolvimento"}

    return html.Div(
        [

            html.Div([
            html.H2(info["titulo"]),
            html.Img(src=fr"./assets/{info['imagem']}",style={"width": "50px",
                                                            "height": "50px","objectFit": "contain",
                                                            "marginBottom": "10px"}),

            ],style={"display": "flex",
                     "alignItems": "center",
                     "justifyContent": "center",
                     "gap": "10px"},),

            html.Hr(style={"width": "80%"}),

            html.P(info["descricao"]),

            html.Button(
                textos.get(info["id"], "Acessar"),
                id=info["id"],
                n_clicks=0,
                style={
                    "width": "200px",
                    "padding": "10px",
                    "backgroundColor": "#F00A0A"
                    #if info["id"] in ['btn-licoes-aprendidas','btn-documentacoes','btn-tickets']
                    if info["id"] in ['btn-licoes-aprendidas','btn-documentacoes']
                    #else "#F00A0A",
                    else "#2C3E50",
                    "color": "white",
                    "border": "none",
                    "borderRadius": "5px",
                    "cursor": "pointer",
                },
            ),

            html.Div(
                id=f"resultado-{info['id']}",
                children="",
                style={
                    "marginTop": "15px",
                    "textAlign": "center",
                    "fontWeight": "bold",
                    "color": "#27AE60",
                    "minHeight": "25px",
                },
            ),
        ],
        style=CARD_STYLE,
    )

def criar_card2(info2):
 
    return html.Div(
        [

            html.Div([
            html.H2(info2["titulo"]),
            html.Img(src=fr"./assets/{info2['imagem']}",style={"width": "50px",
                                                            "height": "50px","objectFit": "contain",
                                                            "marginBottom": "10px"}),

            ],style={"display": "flex",
                     "alignItems": "center",
                     "justifyContent": "center",
                     "gap": "10px"},),

            html.Hr(style={"width": "80%"}),

            html.P(info2["descricao"]),
            html.H3(info2["Colaborador 1"]),        
            html.H3(info2["Colaborador 2"]),        
            html.H3(info2["Colaborador 3"]),        
            html.H3(info2["Colaborador 4"]),
            html.H3(info2["Colaborador 5"]),
            html.Hr(style={"width": "80%"}),
            html.Img(src=fr"./assets/{info2["imagem_roda_pe"]}",style={"width": "300px",
                                                                        "height": "350px","objectFit": "contain",
                                                                        "marginBottom": "10px"}),        

        ],
        style=CARD_STYLE,
    )


def dashboard_tecnica():

    app = dash.Dash(__name__)

    app.layout = html.Div(
        [
            ####INICIO DAS GUIAS.
            dcc.Tabs([
                dcc.Tab(label='Página Inicial', children=[
                    html.Div([html.H1(
                                            "DASHBOARD TÉCNICA - PEÇAS",
                                            style={"margin": "0"}),
                                        html.Img(src=r"./assets/001.png",
                                            style={"height": "50px","marginLeft": "15px",},),],
                                    style={"display": "flex","justifyContent": "center","alignItems": "center","padding": "15px",}),
                    html.Hr(),
                    
                                html.Div(
                                    [criar_card(item) for item in DASHBOARDS],
                                    style={
                                        "display": "grid",
                                        "gridTemplateColumns": "repeat(2, 1fr)",
                                        "gap": "15px",
                                        "padding": "20px",
                                    },
                                ),
                ]),
            dcc.Tab(label='Créditos', children=[
                html.Div([html.H1("INFORMAÇÕES SOBRE O DESENVOLVIMENTO",
                                                            style={"margin": "0"}),
                                                        html.Img(src=r"./assets/001.png",
                                                            style={"height": "50px","marginLeft": "15px",},),],
                                                    style={"display": "flex","justifyContent": "center","alignItems": "center","padding": "15px",}),
                                    html.Hr(),
                        html.Div(
                                                            [criar_card2(item) for item in DASHBOARDS2],
                                                            style={
                                                                "display": "grid",
                                                                "gridTemplateColumns": "repeat(2, 1fr)",
                                                                "gap": "15px",
                                                                "padding": "20px",
                                                            },
                                                        ),
                                    
            ])]),],
        style={
            
            "backgroundColor": "#DBD9EC",
            "minHeight": "100vh",
        },
    )

    # -----------------------------------
    # CALLBACKS DOS BOTÕES
    # -----------------------------------

    for dashboard in DASHBOARDS:

        def criar_callback(arquivo, mensagem):

            def abrir_dashboard(n_clicks):

                if not n_clicks:
                    raise PreventUpdate

                subprocess.Popen([os.path.abspath(arquivo)],cwd=os.path.dirname(os.path.abspath(arquivo)),shell=False)

                return mensagem

            return abrir_dashboard

        app.callback(
            Output(
                f"resultado-{dashboard['id']}",
                "children"
            ),
            Input(
                dashboard["id"],
                "n_clicks"
            ),
            prevent_initial_call=True,
        )(
            criar_callback(
                dashboard["arquivo"],
                dashboard["mensagem"]
            )
        )

    app.run(
        debug=True,
        use_reloader=False,
        port=8050
    )


def abrir_navegador():

    webbrowser.open_new(
        "http://127.0.0.1:8050"
    )


if __name__ == "__main__":

    threading.Timer(
        1,
        abrir_navegador
    ).start()

    dashboard_tecnica()