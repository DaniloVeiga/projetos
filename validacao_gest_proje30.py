#--------------------------------------
#1. IMPORTS
#--------------------------------------
import dash #Importa o framework Dash para criar aplicações web interativas em Python.
from dash import html, dcc, dash_table, Input, Output, State, ctx, no_update
#Importa componentes do Dash:
#html → elementos HTML (div, h1, etc.)
#dcc → componentes interativos (dropdown, graphs)
#dash_table → tabelas
#Input, Output, State → callbacks
#ctx → contexto do callback
#no_update → evita atualização de saída
import pandas as pd #Biblioteca para manipulação de dados (DataFrames).
import plotly.express as px #Criação simplificada de gráficos (ex: Gantt).
import plotly.io as pio #Controle de renderização de gráficos Plotly.
from pathlib import Path #Manipulação de caminhos de arquivos.
import webbrowser #Abre navegador automaticamente.
import threading #Executa tarefas paralelas (ex: abrir navegador com delay).
import tempfile #Criação de arquivos temporários.
from datetime import datetime #Manipulação de datas.
from datetime import date #Manipulação de datas.
from datetime import timedelta #Manipulação de datas.
from dash.exceptions import PreventUpdate #Evita execução de callback quando necessário.
import sqlite3 #Banco de dados SQLite.
#--------------------------------------
#2. CONFIGURAÇÕES
DB_SQLITE = r".\base_sql\devolutivas.db" #Nome do arquivo do banco SQLite.
COLUNAS_DEVOLUTIVA = [
    'KOM_DEVOLUTIVA',
    'Inicio_da_Liberacao_DEVOLUTIVA',
    'Início de Compras - DEVOLUTIVA',
    'Fim de Liberação - DEVOLUTIVA',
    'Fim de Compras - DEVOLUTIVA',
    'Início de Programação/Disposição - DEVOLUTIVA',
    'Fim da Programação/Disposição - DEVOLUTIVA'
] + [f'Inicio_da_Liberacao_G{i}' for i in range(10)]+ [f'Fim_de_Liberacao_DEVOLUTIVA_G{i}' for i in range(10)]

#Lista de colunas relacionadas às devolutivas.
ARQUIVO_EXCEL = Path('./base_validada/base_compilada_padrao.xlsx') #Caminho do arquivo Excel base.
COLUNAS_DATA = [
    'KOM',
    'inicio_da_liberacao',
    'Fim de Liberação',
    'Início de Compras',
    'Fim de Compras',
    'Início de Programação/Disposição',
    'Fim da Programação/Disposição',
    'SOP',
    'ME'
] #Colunas que representam datas no sistema.
ORDEM_FASES = ['Liberação', 'Compras', 'Programação / Disposição'] #Define a ordem lógica das fases do projeto.
ALTURA_GANTT = 650
LARGURA_GANTT = 1500
#Dimensões do gráfico de Gantt.
COLUNAS_DROPDOWN_SIM_NAO = ['NOVO CATÁLOGO','TEVON','KOM_DEVOLUTIVA','Início de Compras - DEVOLUTIVA','Fim de Liberação - DEVOLUTIVA','Fim de Compras - DEVOLUTIVA','Início de Programação/Disposição - DEVOLUTIVA','Fim da Programação/Disposição - DEVOLUTIVA','Inicio_da_Liberacao_G0',
                            'Inicio_da_Liberacao_G1','Inicio_da_Liberacao_G2','Inicio_da_Liberacao_G3','Inicio_da_Liberacao_G4','Inicio_da_Liberacao_G5','Inicio_da_Liberacao_G6','Inicio_da_Liberacao_G7','Inicio_da_Liberacao_G8','Inicio_da_Liberacao_G9','Fim_de_Liberacao_DEVOLUTIVA_G0','Fim_de_Liberacao_DEVOLUTIVA_G1',
                            'Fim_de_Liberacao_DEVOLUTIVA_G2','Fim_de_Liberacao_DEVOLUTIVA_G3',
                            'Fim_de_Liberacao_DEVOLUTIVA_G4','Fim_de_Liberacao_DEVOLUTIVA_G5',
                            'Fim_de_Liberacao_DEVOLUTIVA_G6','Fim_de_Liberacao_DEVOLUTIVA_G7',
                            'Fim_de_Liberacao_DEVOLUTIVA_G8','Fim_de_Liberacao_DEVOLUTIVA_G9'] #Colunas com valores tipo SIM/NÃO.

def contar_usuarios_por_setor():
    conn = sqlite3.connect(r".\base_sql\usuarios.db")

    df = pd.read_sql_query("""
        SELECT setor, COUNT(*) as total
        FROM cadastro
        GROUP BY setor
    """, conn)

    conn.close()

    return dict(zip(df['setor'], df['total']))


HOJE_ISO = date.today().isoformat() #Data atual no formato ISO.
HOJE_ISO = date.today().isoformat()#vermelho bordô
HOJE_MAIS_30 = (date.today() + timedelta(days=30)).isoformat() #vermelho
HOJE_MAIS_60 = (date.today() + timedelta(days=60)).isoformat() #vermelho
HOJE_MAIS_90 = (date.today() + timedelta(days=90)).isoformat() #amarelo
HOJE_MAIS_180 = (date.today() + timedelta(days=180)).isoformat() #azul
HOJE_MAIS_180_PLUS = (date.today() + timedelta(days=190)).isoformat() #verde
SOP_HOJE_MAIS_180 = (date.today() + timedelta(days=180)).isoformat() # hoje menos 1 ano.
SOP_HOJE_MAIS_365 = (date.today() + timedelta(days=365)).isoformat() # hoje menos 1 ano.
ME_HOJE_MAIS_180 = (date.today() + timedelta(days=180)).isoformat() # hoje menos 1 ano.
ME_HOJE_MAIS_365 = (date.today() + timedelta(days=365)).isoformat() # hoje menos 1 ano.


def criar_banco_sqlite():
    conn = sqlite3.connect(DB_SQLITE)
    cur = conn.cursor()

    
    cur.execute("""
            CREATE TABLE IF NOT EXISTS sessao_ativa (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email TEXT,
                nome TEXT,
                data_login TEXT
            )
        """)


    cur.execute("""
        CREATE TABLE IF NOT EXISTS devolutivas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_hora TEXT,
            email TEXT,
            nome TEXT,
            setor TEXT,
            projeto TEXT,
            coluna TEXT,
            valor TEXT,
            Inicio_da_Liberacao_DEVOLUTIVA TEXT,
            Fim_de_Liberacao_DEVOLUTIVA TEXT,
            status_projeto TEXT
        )
    """)

    conn.commit()
    conn.close()
#Cria conexão com banco e cursor.

def buscar_usuario_cadastro(email):
    conn = sqlite3.connect(r".\\base_sql\\usuarios.db")
    cur = conn.cursor()

    cur.execute("""
        SELECT 
            nome, 
            setor, 
            evento,
            Grupo0, Grupo1, Grupo2, Grupo3, Grupo4,
            Grupo5, Grupo6, Grupo7, Grupo8, Grupo9, Grupo_UNLOCKED, Grupo_Unlocked_Compras, Grupo_Unlocked_Prog_Disp
        FROM cadastro
        WHERE lower(email) = lower(?)
    """, (email,))

    resultado = cur.fetchone()
    conn.close()

    if resultado:
        nome = resultado[0]
        setor = resultado[1]
        evento = resultado[2]

        
        nomes_grupos = [
            'Grupo0', 'Grupo1', 'Grupo2', 'Grupo3', 'Grupo4',
            'Grupo5', 'Grupo6', 'Grupo7', 'Grupo8', 'Grupo9','Grupo_UNLOCKED','Grupo_Unlocked_Compras',
            'Grupo_Unlocked_Prog_Disp'
        ]

        grupos = [
            nomes_grupos[i]
            for i, valor in enumerate(resultado[3:])
            if str(valor) == '1'
        ]


        return nome, setor, evento, grupos

    return None


#--------------------------------------
#3. FUNÇÕES
#Carrega uma planilha do Excel.
def carregar_df(sheet):
    df = pd.read_excel(ARQUIVO_EXCEL, sheet_name=sheet)

    # ===== COLUNAS ISO (PARA COMPARAÇÃO) =====
    if 'inicio_da_liberacao' in df.columns:
        df['Inicio_Liberacao_ISO'] = (
            pd.to_datetime(df['inicio_da_liberacao'], dayfirst=True, errors='coerce')
            .dt.strftime('%Y-%m-%d')
        )

    if 'KOM' in df.columns:
        df['KOM_ISO'] = (
            pd.to_datetime(df['KOM'], dayfirst=True, errors='coerce')
            .dt.strftime('%Y-%m-%d')
        )
    if 'Início de Compras' in df.columns:
        df['Inicio_compras_ISO'] = (
            pd.to_datetime(df['Início de Compras'], dayfirst=True, errors='coerce')
            .dt.strftime('%Y-%m-%d')
        )
    if 'Fim de Compras' in df.columns:
        df['Fim_compras_ISO'] = (
            pd.to_datetime(df['Fim de Compras'], dayfirst=True, errors='coerce')
            .dt.strftime('%Y-%m-%d')
        )
    if 'Fim de Liberação' in df.columns:
        df['Fim_Liberacao_ISO'] = (
            pd.to_datetime(df['Fim de Liberação'], dayfirst=True, errors='coerce')
            .dt.strftime('%Y-%m-%d')
        )
    if 'Início de Programação/Disposição' in df.columns:
        df['Início_Prog_Dis_ISO'] = (
            pd.to_datetime(df['Início de Programação/Disposição'], dayfirst=True, errors='coerce')
            .dt.strftime('%Y-%m-%d')
        )
    if 'Fim da Programação/Disposição' in df.columns:
        df['Fim_Prog_Dis_ISO'] = (
            pd.to_datetime(df['Fim da Programação/Disposição'], dayfirst=True, errors='coerce')
            .dt.strftime('%Y-%m-%d')
        )
    if 'SOP' in df.columns:
        df['SOP_ISO'] = (
            pd.to_datetime(df['SOP'], dayfirst=True, errors='coerce')
            .dt.strftime('%Y-%m-%d')
        )
    if 'ME' in df.columns:
        df['ME_ISO'] = (
            pd.to_datetime(df['ME'], dayfirst=True, errors='coerce')
            .dt.strftime('%Y-%m-%d')
        )
        
    # ===== FORMATAÇÃO VISUAL (TABELA) =====
    for col in COLUNAS_DATA:
        if col in df.columns:
            df[col] = (
                pd.to_datetime(df[col], dayfirst=True, errors='coerce')
                .dt.strftime('%d/%m/%Y')
            )

    return df

#Função para gerar gráfico de Gantt.
def criar_gantt(df, fases=None):
    tarefas = [] #Lista de tarefas para montar gráfico.

    for _, row in df.iterrows():
        projeto = row.get('PROJECT')

        def add_fase(ini, fim, fase):
            if fases and fase not in fases:
                return
            if pd.notna(ini) and pd.notna(fim):
                tarefas.append({
                    'Linha': f'{projeto} - {fase}',
                    'Fase': fase,
                    'Início': ini,
                    'Fim': fim
                })

        add_fase(row['inicio_da_liberacao'], row['Fim de Liberação'], 'Liberação')
        add_fase(row['Início de Compras'], row['Fim de Compras'], 'Compras')
        add_fase(
            row['Início de Programação/Disposição'],
            row['Fim da Programação/Disposição'],
            'Programação / Disposição'
        )

    if not tarefas:
        return px.timeline(title='gestão de Projetos – Cronograma')

    fig = px.timeline(
        pd.DataFrame(tarefas),
        x_start='Início',
        x_end='Fim',
        y='Linha',
        color='Fase',
        category_orders={'Fase': ORDEM_FASES}
    )

    fig.update_yaxes(autorange='reversed')
    fig.update_layout(
        height=ALTURA_GANTT,
        width=LARGURA_GANTT,
        autosize=False,
        plot_bgcolor='#FFFFFF',   # área do gráfico
        #paper_bgcolor='#F4F6F7'   # fundo geral

    )
    
    fig.update_xaxes(
        showgrid=True,
        gridcolor='#1C1C1C',
        gridwidth=1,
        #autorange='reversed',
        autorange=True,  # ✅ padrão
        #----------------
        showline=True,        # ✅ mostra a linha do eixo
        linecolor='black',    # cor da linha
        linewidth=1,          # espessura
        mirror=True           # opcional: espelha em cima
        )
    fig.update_yaxes(
        showline=True,        # ✅ mostra a linha do eixo
        linecolor='black',    # cor da linha
        linewidth=1,          # espessura
        mirror=True           # opcional: espelha em cima
        )

    return fig

#Abre automaticamente o dashboard.
def abrir_navegador():
    webbrowser.open_new('http://127.0.0.1:8050/')

#Cria um card visual com título, número e cor.
def card_resumo(titulo, valor, cor):
    #Retorna um componente visual estilizado.
    return html.Div(
        style={
            'backgroundColor': cor,
            'padding': '20px',
            'borderRadius': '10px',
            'boxShadow': '0 4px 10px rgba(0,0,0,0.15)',
            'color': 'white',
            'textAlign': 'center'
        },
        children=[
            html.H4(titulo),
            html.H1(valor, style={'margin': '10px 0'})
        ]
    )
#--------------------------------------
#📥 Dados do usuário
#Busca dados do banco baseado no usuário.
def carregar_devolutivas_usuario(email):
    
    conn = sqlite3.connect(DB_SQLITE)

    df = pd.read_sql_query("""
        SELECT projeto, coluna, valor, Inicio_da_Liberacao_DEVOLUTIVA, Fim_de_Liberacao_DEVOLUTIVA
        FROM devolutivas d1
        WHERE data_hora = (
            SELECT MAX(d2.data_hora)
            FROM devolutivas d2
            WHERE d2.projeto = d1.projeto
              AND d2.coluna = d1.coluna
        )
    """, conn)

    conn.close()
    return df

#🔄 Aplicar devolutivas
#Aplica alterações no DataFrame original.

def aplicar_devolutivas(df_excel, df_devolutivas):

    for _, row in df_devolutivas.iterrows():
        mask = df_excel["PROJECT"] == row["projeto"]

        # ✅ devolutiva normal
        df_excel.loc[mask, row["coluna"]] = row["valor"]

        # ✅ regra CORRIGIDO
        if row["valor"] == "CORRIGIDO":
            df_excel.loc[mask, "Inicio_da_Liberacao_DEVOLUTIVA"] = "PENDENTE"

        # 🔥 PADRONIZAÇÃO AQUI (ESSENCIAL)
        valor = row.get("Inicio_da_Liberacao_DEVOLUTIVA")

        if valor not in [None, '', 'NaT']:
            valor = str(valor).strip().upper()   # ✅ CORREÇÃO
            df_excel.loc[mask, "Inicio_da_Liberacao_DEVOLUTIVA"] = valor

        valor_fim = row.get("Fim_de_Liberacao_DEVOLUTIVA")

        if valor_fim not in [None, '', 'NaT']:
            valor_fim = str(valor_fim).strip().upper()  # ✅ CORREÇÃO
            df_excel.loc[mask, "Fim_de_Liberacao_DEVOLUTIVA"] = valor_fim

    # ✅ 🔥 GARANTE PADRÃO GLOBAL (resolve 90%)
    df_excel["Inicio_da_Liberacao_DEVOLUTIVA"] = (
        df_excel["Inicio_da_Liberacao_DEVOLUTIVA"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    df_excel.loc[
        ~df_excel["Inicio_da_Liberacao_DEVOLUTIVA"].isin(["FEITO", "PENDENTE"]),
        "Inicio_da_Liberacao_DEVOLUTIVA"
    ] = "PENDENTE"

    return df_excel

    

#🌐 Status público
#Regra:
#✅ Todos SIM → DONE
#❌ Caso contrário → vazio
def carregar_status_publico():
    """
    Retorna o status consolidado para o usuário público.
    Regra:
    - SIM (ou DONE) somente se TODOS os usuários do setor responderam SIM
    - qualquer ausência ou NÃO → vazio
    """

    conn = sqlite3.connect(DB_SQLITE)

    query = """
    SELECT
        projeto,
        coluna,
        setor,
        COUNT(DISTINCT email) AS total_respostas,
        SUM(CASE WHEN valor = 'SIM' OR valor = 'DONE' THEN 1 ELSE 0 END) AS total_sim
    FROM devolutivas
    GROUP BY projeto, setor, coluna
    """

    df = pd.read_sql_query(query, conn)
    conn.close()
# aplica a regra final em Python (mais legível e seguro)
    def calcular_status(row):
        
        totais = contar_usuarios_por_setor()
        total_esperado = totais.get(row["setor"], 0)

        if row["total_sim"] == total_esperado:
            return "SIM"
        return ""

    df["status"] = df.apply(calcular_status, axis=1)

    return df[["projeto", "coluna", "setor", "status"]]


def carregar_df_sqlite():
    conn = sqlite3.connect(r".\base_sql\demandas.db")

    df = pd.read_sql_query("""
        SELECT DATA_PROJ, PROGRAMA, PROJECT, PRODUCT, NOVO_CAT, NUM_CAT, FRONT_PAGE, TEVON,KOM, KOM_DEVOLUTIVA,
                           inicio_da_liberacao, Inicio_da_Liberacao_DEVOLUTIVA,
                           Inicio_da_Liberacao_G0, Inicio_da_Liberacao_G1, Inicio_da_Liberacao_G2  
                           , Inicio_da_Liberacao_G3, Inicio_da_Liberacao_G4
                           , Inicio_da_Liberacao_G5, Inicio_da_Liberacao_G6, Inicio_da_Liberacao_G7
                           , Inicio_da_Liberacao_G8, Inicio_da_Liberacao_G9, 
                           Fim_de_Liberacao_DEVOLUTIVA_G0, Fim_de_Liberacao_DEVOLUTIVA_G1,
                           Fim_de_Liberacao_DEVOLUTIVA_G2, Fim_de_Liberacao_DEVOLUTIVA_G3
                           , Fim_de_Liberacao_DEVOLUTIVA_G4, Fim_de_Liberacao_DEVOLUTIVA_G5
                           , Fim_de_Liberacao_DEVOLUTIVA_G6, Fim_de_Liberacao_DEVOLUTIVA_G7
                           , Fim_de_Liberacao_DEVOLUTIVA_G8, Fim_de_Liberacao_DEVOLUTIVA_G9,
                           Fim_de_Liberacao_DEVOLUTIVA FROM tarefas
    """, conn)

    conn.close()

    # 🔄 RENOMEAR colunas para ficar igual ao Excel
    df.rename(columns={
        'DATA_PROJ': 'DATA',
        'NOVO_CAT': 'NOVO CATÁLOGO',
        'NUM_CAT': 'NÚMERO CATÁLOGO',
        'inicio_da_liberacao': 'inicio_da_liberacao',
        'Inicio_da_Liberacao_DEVOLUTIVA':'Inicio_da_Liberacao_DEVOLUTIVA',
        'Inicio_da_Liberacao_G0':'Inicio_da_Liberacao_G0',
        'Inicio_da_Liberacao_G1':'Inicio_da_Liberacao_G1',
        'Inicio_da_Liberacao_G2':'Inicio_da_Liberacao_G2',
        'Inicio_da_Liberacao_G3':'Inicio_da_Liberacao_G3',
        'Inicio_da_Liberacao_G4':'Inicio_da_Liberacao_G4',
        'Inicio_da_Liberacao_G5':'Inicio_da_Liberacao_G5',
        'Inicio_da_Liberacao_G6':'Inicio_da_Liberacao_G6',
        'Inicio_da_Liberacao_G7':'Inicio_da_Liberacao_G7',
        'Inicio_da_Liberacao_G8':'Inicio_da_Liberacao_G8',
        'Inicio_da_Liberacao_G9':'Inicio_da_Liberacao_G9',
        'Fim_de_Liberacao_DEVOLUTIVA_G0':'Fim_de_Liberacao_DEVOLUTIVA_G0',
        'Fim_de_Liberacao_DEVOLUTIVA_G1':'Fim_de_Liberacao_DEVOLUTIVA_G1',
        'Fim_de_Liberacao_DEVOLUTIVA_G2':'Fim_de_Liberacao_DEVOLUTIVA_G2',
        'Fim_de_Liberacao_DEVOLUTIVA_G3':'Fim_de_Liberacao_DEVOLUTIVA_G3',
        'Fim_de_Liberacao_DEVOLUTIVA_G4':'Fim_de_Liberacao_DEVOLUTIVA_G4',
        'Fim_de_Liberacao_DEVOLUTIVA_G5':'Fim_de_Liberacao_DEVOLUTIVA_G5',
        'Fim_de_Liberacao_DEVOLUTIVA_G6':'Fim_de_Liberacao_DEVOLUTIVA_G6',
        'Fim_de_Liberacao_DEVOLUTIVA_G7':'Fim_de_Liberacao_DEVOLUTIVA_G7',
        'Fim_de_Liberacao_DEVOLUTIVA_G8':'Fim_de_Liberacao_DEVOLUTIVA_G8',
        'Fim_de_Liberacao_DEVOLUTIVA_G9':'Fim_de_Liberacao_DEVOLUTIVA_G9',
        
        'inicio_da_liberacao': 'inicio_da_liberacao',
        'fim_da_liberacao': 'Fim de Liberação',
        'inicio_de_compras': 'Início de Compras',
        'fim_de_compras': 'Fim de Compras',

        'COMMENTS': 'COMMENTS'
    }, inplace=True)

    
    # ✅ CRIAR COLUNA ISO PARA INICIO DA LIBERAÇÃO
   
    if 'inicio_da_liberacao' in df.columns:

        
        # ✅ base datetime única
        data_dt = pd.to_datetime(df['inicio_da_liberacao'], errors='coerce')


        # ISO (comparação)
        df['Inicio_Liberacao_ISO'] = data_dt.dt.strftime('%Y-%m-%d')

        # display (tela)
        df['inicio_da_liberacao'] = data_dt.dt.strftime('%d/%m/%Y')

    
    if 'KOM' in df.columns:
        df['KOM_ISO'] = (
            pd.to_datetime(df['KOM'], dayfirst=True, errors='coerce')
            .dt.strftime('%Y-%m-%d')
        )




    # ✅ Ajusta datas (igual você já faz no Excel)
    
    for col in COLUNAS_DATA:
        if col in df.columns and col != 'inicio_da_liberacao':  # ✅ ESSENCIAL
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d/%m/%Y')


    return df


def verificar_sessao_ativa():
    conn = sqlite3.connect(DB_SQLITE)
    cur = conn.cursor()

    cur.execute("SELECT nome, email FROM sessao_ativa LIMIT 1")
    resultado = cur.fetchone()

    conn.close()

    return resultado  # None ou (nome, email)

def criar_sessao(usuario):
    conn = sqlite3.connect(DB_SQLITE)
    cur = conn.cursor()

    cur.execute("DELETE FROM sessao_ativa")  # garante 1 só sessão

    cur.execute("""
        INSERT INTO sessao_ativa (email, nome, data_login)
        VALUES (?, ?, ?)
    """, (
        usuario["email"],
        usuario["usuario"],
        datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    ))

    conn.commit()
    conn.close()


def encerrar_sessao():
    conn = sqlite3.connect(DB_SQLITE)
    cur = conn.cursor()

    cur.execute("DELETE FROM sessao_ativa")

    conn.commit()
    conn.close()

def resetar_banco():
    conn = sqlite3.connect(DB_SQLITE)
    cur = conn.cursor()

    # 🔥 apaga todas as devolutivas (LOG)
    cur.execute("DELETE FROM devolutivas")

    # 🔥 reseta o autoincremento (opcional, mas recomendado)
    cur.execute("DELETE FROM sqlite_sequence WHERE name='devolutivas'")

    conn.commit()
    conn.close()
#--------------------------------------
#3.1 LEGENDAS
#Cria item de legenda com cor e texto.
def legenda_item(cor, texto):
    return html.Div(
        style={
            'display': 'flex',
            'alignItems': 'center',
            'gap': '6px'
        },
        children=[
            html.Div(
                style={
                    'width': '16px',
                    'height': '16px',
                    'backgroundColor': cor,
                    'border': '1px solid #333'
                }
            ),
            html.Span(texto, style={'fontSize': '13px'})
        ]
    )

#Legenda da tabela com cores de prazo.
def legenda_cores_tabela():
    return html.Div(
        style={
            'display': 'flex',
            'gap': '20px',
            'alignItems': 'center',
            'flexWrap': 'wrap',
            'marginBottom': '10px'
        },
        children=[
            legenda_item('#800020', 'Atrasado'),
            legenda_item('#FF0000', 'Até 30 dias'),
            legenda_item('#FFFF00', 'Até 90 dias'),
            legenda_item('#1E90FF', 'Até 180 dias'),
            legenda_item('#00FF2AC3', 'Acima de 180 dias'),
        ]
    )

#Legenda específica para SOP / ME.
def legenda_cores_tabela_SOPME():
    return html.Div(
        style={
            'display': 'flex',
            'gap': '20px',
            'alignItems': 'center',
            'flexWrap': 'wrap',
            'marginBottom': '10px'
        },
        children=[
            legenda_item('#FF0000', 'Até 180 dias'),
            legenda_item('#FFFF00', 'Acima de 180 dias'),
            legenda_item('#00FF2AC3', 'Acima de 365 dias'),
        ]
    )

#--------------------------------------
#4. APP
app = dash.Dash(__name__, suppress_callback_exceptions=True)
#Inicializa aplicação Dash.
app.title = 'gestão de Projetos – Dashboard'
#Define título do navegador.
df_controle = carregar_df_sqlite()
#Carrega dados principais.
PROGRAMAS = sorted(df_controle['PROGRAMA'].dropna().unique())
PROJETOS = sorted(df_controle['PROJECT'].dropna().unique())
DATAS_TEXTO = df_controle['DATA'].fillna('').astype(str).unique().tolist()
#Listas para filtros.
#--------------------------------------
# 5. LAYOUT (ORIGINAL)
app.layout = html.Div([

    dcc.Store(id='dados-originais-usuario',storage_type='session'),
    dcc.Store(id='usuario-store'),

    # 🔐 Store isolado por usuário
    dcc.Store(
        id='dados-usuario',
        storage_type='session'
    ),
html.Div(
    #Define estrutura visual inicial.
    id='login-container',
    #Container da tela de login.
    style={
        'display': 'flex',
        'justifyContent': 'center',
        'alignItems': 'center',
        'height': '100vh',
        'backgroundColor': '#F4F6F7',
        'gap': '30px'
    },
    children=[

        # ===== BOX 1 – PERFIS =====
        html.Div(
            style={
                'backgroundColor': 'white',
                'padding': '30px 40px',
                'borderRadius': '8px',
                'boxShadow': '0 4px 10px rgba(0,0,0,0.15)',
                'minWidth': '360px',
                'minHeight': '500px',  # ✅ ADICIONADO
                'textAlign': 'center'
            },
            children=[
                html.H2('Gerenciamento de Perfis'),
                html.P(
                    'Controle edição de projetos',
                    style={'color': '#7F8C8D'}
                ),
                dcc.Input(
                    id='input-email',
                    type='text',
                    placeholder='Digite seu e-mail',
                    style={'width': '100%', 'padding': '8px'}
                ),
                
                html.Br(), html.Br(),
                html.Button(
                    'Acessar Perfis',
                    id='btn-login',
                    style={
                        'width': '100%',
                        'padding': '10px',
                        'backgroundColor': '#2C3E50',
                        'color': 'white',
                        'border': 'none'
                    }
                ),

                
                html.P(
                    id='setor-preview',
                    style={
                        'marginTop': '10px',
                        'fontWeight': 'bold',
                        'color': '#16A085'
                    }
                ),


                html.Div(
                    id='msg-login',
                    style={'color': 'red', 'marginTop': '10px'}
                )
            ]
        ),

        # ===== BOX 2 – Dashboard =====
        html.Div(
            style={
                'backgroundColor': 'white',
                'padding': '30px 40px',
                'borderRadius': '8px',
                'boxShadow': '0 4px 10px rgba(0,0,0,0.15)',
                'minWidth': '360px',
                'minHeight': '500px',  # ✅ ADICIONADO
                'textAlign': 'center'
            },
            children=[
                html.H2('Acesso ao Dashboard'),
                html.P(
                    'Consulta de projetos',
                    style={'color': '#7F8C8D'}
                ),
                html.Br(),
                html.Button(
                    'Acessar Dashboard',
                    id='btn-perfis',
                    style={
                        'width': '100%',
                        'padding': '10px',
                        'backgroundColor': '#16A085',
                        'color': 'white',
                        'border': 'none'
                    }
                )
            ]
        )

    ]
),


    # ---------- DASHBOARD ----------
    html.Div(
        id='dashboard-container',
        #Container dO DASHBOARD
        style={'display': 'none', 'height': '100vh', 'overflow': 'hidden'},
        children=[

            html.Div(
                style={'minWidth': '1500px', 'overflowX': 'auto', 'padding': '15px'},
                children=[

                    html.Div(
                        style={'display': 'flex', 'justifyContent': 'space-between'},
                        children=[
                            html.H3('gestão de Projetos – Dashboard'), html.Br(),
                       
                            html.Div([
                                html.Span(id='usuario-logado'), html.Br(),
                                html.Span(id='setor-usuario'), html.Br(),
                                html.Span(id='evento-usuario'), html.Br(),  # 👈 NOVO
                                html.Span(id='grupos-usuario'), html.Br(), # ✅ NOVO
                                html.Span(id='data-acesso'), html.Br(), html.Br(),
                                html.Button('Sair', id='btn-logout',
                                            style={'backgroundColor': '#E74C3C', 'color': 'white'})
                            ])
                        ]
                    ),

                    html.Hr(),

                    dcc.Tabs([

                        # ===================== TABELA =====================
                        dcc.Tab(label='Tabela', children=[
                            #
                            html.Div(
                            style={'display': 'flex', 'alignItems': 'center', 'gap': '10px'},
                            children=[
                                html.Button(
                                    'Salvar Alterações - Técnica',
                                    id='btn-salvar',
                                    style={
                                        'backgroundColor': '#27AE60',
                                        'color': 'white',
                                        'padding': '8px 15px',
                                        'border': 'none',
                                        'fontWeight': 'bold'
                                    }
                                ),
                                html.Span(
                                    id='msg-salvar',
                                    style={'fontSize': '13px'}
                                )
                            ]
                        ),
                        html.Br(),

                            #

                            #INÍCIO DA TABELA DE LEGENDAS                            
                            # ===== LEGENDAS EM GRID =====
                            html.Div(
                                style={
                                    'display': 'grid',
                                    'gridTemplateColumns': '1fr 1fr',
                                    'gap': '30px',
                                    'marginBottom': '15px'
                                },
                                children=[

                                    # --- COLUNA 1 ---
                                    html.Div([
                                        html.H4('Legenda de Cores - Escopo'),
                                        legenda_cores_tabela(),
                                    ]),

                                    # --- COLUNA 2 ---
                                    html.Div([
                                        html.H4('Legenda de Cores - SOP & ME'),
                                        legenda_cores_tabela_SOPME(),
                                    ])
                                ]
                            ),

                            html.Br(),
                            #FIM DA TABELA DE LEGENDAS
                            
                            html.Div(
                                style={'display': 'flex', 'gap': '15px', 'flexWrap': 'wrap'},
                                children=[
                                    dcc.Dropdown(id='tabela-filtro-programa',
                                                 options=[{'label': p, 'value': p} for p in PROGRAMAS],
                                                 multi=True, placeholder='Programa'),
                                    dcc.Dropdown(id='tabela-filtro-projeto',
                                                 options=[{'label': p, 'value': p} for p in PROJETOS],
                                                 multi=True, placeholder='Project'),
                                    dcc.Dropdown(id='tabela-filtro-data',
                                                 options=[{'label': d if d else '(Vazio)', 'value': d}
                                                          for d in DATAS_TEXTO],
                                                 multi=True, placeholder='Data'),
                                    html.Button('Limpar filtros', id='btn-limpar-filtros-tabela')
                                ]
                            ),

                            html.Br(),

                            dash_table.DataTable(
                                id='tabela-editavel',
                                #data=df_controle.to_dict('records'),
                                data=[],
                                # ===== INÍCIO DO BLOCO DE PRIVILÉGIO =====
                                #columns=[],  # 🔐 será controlado pelo privilégio
                              
                                columns=[],

                                # ===== FIM DO BLOCO DE PRIVILÉGIO =====

                                
                                editable=True,           # 🔒 bloqueio total
                                cell_selectable=True,    # 🔒 evita foco/edição

                                page_size=15,
                                style_table={'overflowX': 'auto'},
                                style_cell={'fontSize': 11},
                                dropdown={
                                # SIM / NÃO
                                **{
                                    c: {
                                        'options': [
                                            {'label': 'SIM', 'value': 'SIM'},
                                            {'label': 'NÃO', 'value': 'NÃO'},
                                        ]
                                    }
                                    for c in COLUNAS_DROPDOWN_SIM_NAO
                                },
                                    },
                                # ✅ ÚNICA ALTERAÇÃO DO PROJETO
                                style_data_conditional=[
                                    {
                                        'if': {
                                            'column_id': 'NOVO CATÁLOGO',
                                            'filter_query': '{NOVO CATÁLOGO} = "NÃO"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    #---------------KOM - FORMATAÇÃO CONDICIONAL
                                    {
                                        'if': {
                                            'column_id': 'KOM',
                                            'filter_query': (
                                                f'{{KOM_ISO}} < "{HOJE_ISO}"'
                                            )
                                        },
                                        'backgroundColor': '#CD2626',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'KOM',
                                            'filter_query': (
                                                f'{{KOM_ISO}} > "{HOJE_ISO}"'
                                            )
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'KOM',
                                            'filter_query': (
                                                f'{{KOM_ISO}} > "{HOJE_MAIS_30}"'

                                            )
                                        },
                                        'backgroundColor': '#FFFF00',
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'KOM',
                                            'filter_query': (
                                                f'{{KOM_ISO}} > "{HOJE_MAIS_90}"'

                                            )
                                        },
                                        'backgroundColor': "#006AFF",
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'KOM',
                                            'filter_query': (
                                                f'{{KOM_ISO}} > "{HOJE_MAIS_180}"'
                                            )
                                        },
                                        'backgroundColor': "#00FF2AC3",
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'inicio_da_liberacao',
                                            'filter_query': (
                                                f'{{Inicio_Liberacao_ISO}} < "{HOJE_ISO}"'
                                            )
                                        },
                                        'backgroundColor': '#CD2626',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'inicio_da_liberacao',
                                            'filter_query': (
                                                f'{{Inicio_Liberacao_ISO}} > "{HOJE_ISO}"'
                                            )
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'inicio_da_liberacao',
                                            'filter_query': (
                                                f'{{Inicio_Liberacao_ISO}} > "{HOJE_MAIS_30}"'

                                            )
                                        },
                                        'backgroundColor': '#FFFF00',
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'inicio_da_liberacao',
                                            'filter_query': (
                                                f'{{Inicio_Liberacao_ISO}} > "{HOJE_MAIS_90}"'

                                            )
                                        },
                                        'backgroundColor': "#006AFF",
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'inicio_da_liberacao',
                                            'filter_query': (
                                                f'{{Inicio_Liberacao_ISO}} > "{HOJE_MAIS_180}"'
                                            )
                                        },
                                        'backgroundColor': "#00FF2AC3",
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    #---------------Início da Liberação - FORMATAÇÃO CONDICIONAL
                                    {
                                        'if': {
                                            'column_id': 'Início de Compras',
                                            'filter_query': (
                                                f'{{Inicio_compras_ISO}} < "{HOJE_ISO}"'
                                            )
                                        },
                                        'backgroundColor': '#CD2626',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Início de Compras',
                                            'filter_query': (
                                                f'{{Inicio_compras_ISO}} > "{HOJE_ISO}"'
                                            )
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Início de Compras',
                                            'filter_query': (
                                                f'{{Inicio_compras_ISO}} > "{HOJE_MAIS_30}"'

                                            )
                                        },
                                        'backgroundColor': '#FFFF00',
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Início de Compras',
                                            'filter_query': (
                                                f'{{Inicio_compras_ISO}} > "{HOJE_MAIS_90}"'

                                            )
                                        },
                                        'backgroundColor': "#006AFF",
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Início de Compras',
                                            'filter_query': (
                                                f'{{Inicio_compras_ISO}} > "{HOJE_MAIS_180}"'
                                            )
                                        },
                                        'backgroundColor': "#00FF2AC3",
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    #
                                    {
                                        'if': {
                                            'column_id': 'Fim de Compras',
                                            'filter_query': (
                                                f'{{Fim_compras_ISO}} < "{HOJE_ISO}"'
                                            )
                                        },
                                        'backgroundColor': '#CD2626',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim de Compras',
                                            'filter_query': (
                                                f'{{Fim_compras_ISO}} > "{HOJE_ISO}"'
                                            )
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim de Compras',
                                            'filter_query': (
                                                f'{{Fim_compras_ISO}} > "{HOJE_MAIS_30}"'

                                            )
                                        },
                                        'backgroundColor': '#FFFF00',
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim de Compras',
                                            'filter_query': (
                                                f'{{Fim_compras_ISO}} > "{HOJE_MAIS_90}"'

                                            )
                                        },
                                        'backgroundColor': "#006AFF",
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim de Compras',
                                            'filter_query': (
                                                f'{{Fim_compras_ISO}} > "{HOJE_MAIS_180}"'
                                            )
                                        },
                                        'backgroundColor': "#00FF2AC3",
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                   #
                                    {
                                        'if': {
                                            'column_id': 'Fim de Liberação',
                                            'filter_query': (
                                                f'{{Fim_Liberacao_ISO}} < "{HOJE_ISO}"'
                                            )
                                        },
                                        'backgroundColor': '#CD2626',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim de Liberação',
                                            'filter_query': (
                                                f'{{Fim_Liberacao_ISO}} > "{HOJE_ISO}"'
                                            )
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim de Liberação',
                                            'filter_query': (
                                                f'{{Fim_Liberacao_ISO}} > "{HOJE_MAIS_30}"'

                                            )
                                        },
                                        'backgroundColor': '#FFFF00',
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim de Liberação',
                                            'filter_query': (
                                                f'{{Fim_Liberacao_ISO}} > "{HOJE_MAIS_90}"'

                                            )
                                        },
                                        'backgroundColor': "#006AFF",
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim de Liberação',
                                            'filter_query': (
                                                f'{{Fim_Liberacao_ISO}} > "{HOJE_MAIS_180}"'
                                            )
                                        },
                                        'backgroundColor': "#00FF2AC3",
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                   #
                                   #
                                    {
                                        'if': {
                                            'column_id': 'Início de Programação/Disposição',
                                            'filter_query': (
                                                f'{{Início_Prog_Dis_ISO}} < "{HOJE_ISO}"'
                                            )
                                        },
                                        'backgroundColor': '#CD2626',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Início de Programação/Disposição',
                                            'filter_query': (
                                                f'{{Início_Prog_Dis_ISO}} > "{HOJE_ISO}"'
                                            )
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Início de Programação/Disposição',
                                            'filter_query': (
                                                f'{{Início_Prog_Dis_ISO}} > "{HOJE_MAIS_30}"'

                                            )
                                        },
                                        'backgroundColor': '#FFFF00',
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Início de Programação/Disposição',
                                            'filter_query': (
                                                f'{{Início_Prog_Dis_ISO}} > "{HOJE_MAIS_90}"'

                                            )
                                        },
                                        'backgroundColor': "#006AFF",
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Início de Programação/Disposição',
                                            'filter_query': (
                                                f'{{Início_Prog_Dis_ISO}} > "{HOJE_MAIS_180}"'
                                            )
                                        },
                                        'backgroundColor': "#00FF2AC3",
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    #
                                   #
                                    {
                                        'if': {
                                            'column_id': 'Fim da Programação/Disposição',
                                            'filter_query': (
                                                f'{{Fim_Prog_Dis_ISO}} < "{HOJE_ISO}"'
                                            )
                                        },
                                        'backgroundColor': '#CD2626',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim da Programação/Disposição',
                                            'filter_query': (
                                                f'{{Fim_Prog_Dis_ISO}} > "{HOJE_ISO}"'
                                            )
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim da Programação/Disposição',
                                            'filter_query': (
                                                f'{{Fim_Prog_Dis_ISO}} > "{HOJE_MAIS_30}"'

                                            )
                                        },
                                        'backgroundColor': '#FFFF00',
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim da Programação/Disposição',
                                            'filter_query': (
                                                f'{{Fim_Prog_Dis_ISO}} > "{HOJE_MAIS_90}"'

                                            )
                                        },
                                        'backgroundColor': "#006AFF",
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim da Programação/Disposição',
                                            'filter_query': (
                                                f'{{Fim_Prog_Dis_ISO}} > "{HOJE_MAIS_180}"'
                                            )
                                        },
                                        'backgroundColor': "#00FF2AC3",
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    #
                                    {
                                        'if': {
                                            'column_id': 'SOP',
                                            'filter_query': f'{{SOP_ISO}} < "{SOP_HOJE_MAIS_180}"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'SOP',
                                            'filter_query': f'{{SOP_ISO}} > "{SOP_HOJE_MAIS_180}"'
                                        },
                                        'backgroundColor': '#FFFF00',
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'SOP',
                                            'filter_query': f'{{SOP_ISO}} > "{SOP_HOJE_MAIS_365}"'
                                        },
                                        'backgroundColor': '#00FF2AC3',
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    #
                                    {
                                        'if': {
                                            'column_id': 'ME',
                                            'filter_query': f'{{ME_ISO}} < "{ME_HOJE_MAIS_180}"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'ME',
                                            'filter_query': f'{{ME_ISO}} > "{ME_HOJE_MAIS_180}"'
                                        },
                                        'backgroundColor': '#FFFF00',
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'ME',
                                            'filter_query': f'{{ME_ISO}} > "{ME_HOJE_MAIS_365}"'
                                        },
                                        'backgroundColor': '#00FF2AC3',
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'NOVO CATÁLOGO',
                                            'filter_query': '{NOVO CATÁLOGO} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'NOVO CATÁLOGO',
                                            'filter_query': '{NOVO CATÁLOGO} = "DEFINIR"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'NOVO CATÁLOGO',
                                            'filter_query': '{NOVO CATÁLOGO} = "DONE"'
                                        },
                                        'backgroundColor': '#BFBFBF',
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'NÚMERO CATÁLOGO',
                                            'filter_query': '{NÚMERO CATÁLOGO} = "DEFINIR"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'NÚMERO CATÁLOGO',
                                            'filter_query': '{NÚMERO CATÁLOGO} = "DONE"'
                                        },
                                        'backgroundColor': '#BFBFBF',
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'FRONT_PAGE',
                                            'filter_query': '{FRONT_PAGE} = "OK"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'FRONT_PAGE',
                                            'filter_query': '{FRONT_PAGE} = "DONE"'
                                        },
                                        'backgroundColor': '#BFBFBF',
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'TEVON',
                                            'filter_query': '{TEVON} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'TEVON',
                                            'filter_query': '{FRONT PAGE} = "DONE"'
                                        },
                                        'backgroundColor': '#BFBFBF',
                                        'color': 'black',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'KOM_DEVOLUTIVA',
                                            'filter_query': '{KOM_DEVOLUTIVA} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'KOM_DEVOLUTIVA',
                                            'filter_query': '{KOM_DEVOLUTIVA} = "NÃO"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_DEVOLUTIVA',
                                            'filter_query': '{Inicio_da_Liberacao_DEVOLUTIVA} = "FEITO"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_DEVOLUTIVA',
                                            'filter_query': '{Inicio_da_Liberacao_DEVOLUTIVA} = "PENDENTE"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Início de Compras - DEVOLUTIVA',
                                            'filter_query': '{Início de Compras - DEVOLUTIVA} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Início de Compras - DEVOLUTIVA',
                                            'filter_query': '{Início de Compras - DEVOLUTIVA} = "NÃO"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim_de_Liberacao_DEVOLUTIVA',
                                            'filter_query': '{Fim_de_Liberacao_DEVOLUTIVA} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim_de_Liberacao_DEVOLUTIVA',
                                            'filter_query': '{Fim_de_Liberacao_DEVOLUTIVA} = "NÃO"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim de Compras - DEVOLUTIVA',
                                            'filter_query': '{Fim de Compras - DEVOLUTIVA} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim de Compras - DEVOLUTIVA',
                                            'filter_query': '{Fim de Compras - DEVOLUTIVA} = "NÃO"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Início de Programação/Disposição - DEVOLUTIVA',
                                            'filter_query': '{Início de Programação/Disposição - DEVOLUTIVA} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Início de Programação/Disposição - DEVOLUTIVA',
                                            'filter_query': '{Início de Programação/Disposição - DEVOLUTIVA} = "NÃO"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim da Programação/Disposição - DEVOLUTIVA',
                                            'filter_query': '{Fim da Programação/Disposição - DEVOLUTIVA} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Fim da Programação/Disposição - DEVOLUTIVA',
                                            'filter_query': '{Fim da Programação/Disposição - DEVOLUTIVA} = "NÃO"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G0',
                                            'filter_query': '{Inicio_da_Liberacao_G0} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G0',
                                            'filter_query': '{Inicio_da_Liberacao_G0} = "NÃO"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G1',
                                            'filter_query': '{Inicio_da_Liberacao_G1} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G1',
                                            'filter_query': '{Inicio_da_Liberacao_G1} = "NÃO"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G2',
                                            'filter_query': '{Inicio_da_Liberacao_G2} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G2',
                                            'filter_query': '{Inicio_da_Liberacao_G2} = "NÃO"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G3',
                                            'filter_query': '{Inicio_da_Liberacao_G3} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G3',
                                            'filter_query': '{Inicio_da_Liberacao_G3} = "NÃO"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G4',
                                            'filter_query': '{Inicio_da_Liberacao_G4} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G4',
                                            'filter_query': '{Inicio_da_Liberacao_G4} = "NÃO"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G5',
                                            'filter_query': '{Inicio_da_Liberacao_G5} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G5',
                                            'filter_query': '{Inicio_da_Liberacao_G5} = "NÃO"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G6',
                                            'filter_query': '{Inicio_da_Liberacao_G6} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G6',
                                            'filter_query': '{Inicio_da_Liberacao_G6} = "NÃO"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G7',
                                            'filter_query': '{Inicio_da_Liberacao_G7} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G7',
                                            'filter_query': '{Inicio_da_Liberacao_G7} = "NÃO"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G8',
                                            'filter_query': '{Inicio_da_Liberacao_G8} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G8',
                                            'filter_query': '{Inicio_da_Liberacao_G8} = "NÃO"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G9',
                                            'filter_query': '{Inicio_da_Liberacao_G9} = "SIM"'
                                        },
                                        'backgroundColor': '#00B050',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    },
                                    {
                                        'if': {
                                            'column_id': 'Inicio_da_Liberacao_G9',
                                            'filter_query': '{Inicio_da_Liberacao_G9} = "NÃO"'
                                        },
                                        'backgroundColor': '#FF0000',
                                        'color': 'white',
                                        'fontWeight': 'bold'
                                    }
                                ]
                            )
                        ]),

                        # ===================== CRONOGRAMA =====================
                        dcc.Tab(label='Cronograma', children=[
                            html.Br(),

                            html.Div(
                                style={'display': 'flex', 'gap': '15px', 'flexWrap': 'wrap'},
                                children=[
                                    dcc.Dropdown(id='filtro-programa',
                                                 options=[{'label': p, 'value': p} for p in PROGRAMAS],
                                                 multi=True, placeholder='Programa'),
                                    dcc.Dropdown(id='filtro-projeto',
                                                 options=[{'label': p, 'value': p} for p in PROJETOS],
                                                 multi=True, placeholder='Projeto'),
                                    dcc.Dropdown(id='filtro-fase',
                                                 options=[{'label': f, 'value': f} for f in ORDEM_FASES],
                                                 multi=True, placeholder='Fase'),
                                    dcc.DatePickerRange(id='filtro-data',
                                                        display_format='DD/MM/YYYY')
                                ]
                            ),

                            html.Br(),

                            html.Div(
                                style={'display': 'flex', 'justifyContent': 'space-between'},
                                children=[
                                    html.Button('Limpar filtros', id='btn-limpar-filtros'),
                                    html.Div([
                                        html.Button('Exportar PNG', id='btn-exportar-png'),
                                        html.Button('Exportar PDF', id='btn-exportar-pdf'),
                                        dcc.Download(id='download-png'),
                                        dcc.Download(id='download-pdf')
                                    ])
                                ]
                            ),

                            dcc.Graph(
                                id='grafico-gantt',
                                config={'displaylogo': False,
                                        'modeBarButtonsToRemove': ['toImage']}
                            )
                        ]),
                        
                        # ===================== OVERVIEW =====================
                            dcc.Tab(label='Overview', children=[
                                html.Br(),

                                html.Div(
                                    id='overview-cards',
                                    style={
                                        'display': 'grid',
                                        'gridTemplateColumns': 'repeat(3, 1fr)',
                                        'gap': '20px'
                                    }
                                )
                            ]),
                        # ===================== LOG =====================
                            dcc.Tab(label='LOG', children=[
                            html.Br(),

                            html.Div(
                                style={'display': 'flex', 'justifyContent': 'space-between'},
                                children=[
                                    html.H4('LOG – Histórico de Devolutivas'),
                                    html.Button(
                                        'Atualizar',
                                        id='btn-atualizar-log',
                                        style={
                                            'backgroundColor': '#34495E',
                                            'color': 'white',
                                            'border': 'none',
                                            'padding': '6px 12px'
                                        }
                                    )
                                ]
                            ),

                            html.Br(),

                            dash_table.DataTable(
                                id='tabela-log',
                                page_size=15,
                                sort_action='native',
                                filter_action='native',
                                style_table={'overflowX': 'auto'},
                                style_cell={'fontSize': 11, 'textAlign': 'left'},
                            )
                        ])
                    ])
                ]
            )
        ]
    )
])

#--------------------------------------
#6. LOGIN

@app.callback(
    Output('setor-preview', 'children'),
    Input('input-email', 'value')
)
def mostrar_setor_por_email(email):

    
    if not email:
       return ''

    if '@' not in email:
        return '⌛ Digite um e-mail válido'


    dados = buscar_usuario_cadastro(email)

    if not dados:
        return '❌ E-mail não cadastrado'

    nome, setor, evento, grupos = dados

    # ✅ NÃO É TÉCNICA
    ##############AQUI
    if setor.lower() != 'técnica':
        options = [
        {"label": grupo, "value": grupo}
        for grupo in grupos
        ]

        #return f'Setor identificado: {setor}'
        return html.Div([
        html.Span(f'Setor: {setor}'),
        html.Br(),

        dcc.Dropdown(
            id='dropdown-grupos-preview',   # ✅ novo id
            options=options,
            placeholder='Selecione seu grupo',
            disabled=True,  # ✅ BLOQUEIA
            style={'marginTop': '10px'}
        ), 
    ])

    # ✅ É TÉCNICA → monta options dinâmico
    options = [
        {"label": grupo, "value": grupo}
        for grupo in grupos
    ]

    return html.Div([
        html.Span(f'Setor: {setor}'),
        html.Br(),

        dcc.Dropdown(
            id='dropdown-grupos-preview',   # ✅ novo id
            options=options,
            placeholder='Selecione seu grupo',
            style={'marginTop': '10px'}
        )
    ])

@app.callback(
    Output('usuario-store', 'data'),
    Output('msg-login', 'children'),
    Input('btn-login', 'n_clicks'),
    Input('btn-perfis', 'n_clicks'),
    Input('btn-logout', 'n_clicks'),
    State('input-email', 'value'),
    State('dropdown-grupos-preview', 'value'),
    prevent_initial_call=True
)
#Controla login/logout.
def login_logout(n_login, n_dashboard, n_logout, email, grupo_selecionado):

    grupo_selecionado = grupo_selecionado or None
    trigger = ctx.triggered_id

    # ✅ LOGOUT
    if trigger == 'btn-logout':
        encerrar_sessao()  # ✅ ADICIONAR
        return None, ''

    # ✅ ACESSO PÚBLICO (sem login)
    if trigger == 'btn-perfis':

        # 🔴 BLOQUEIO DE SESSÃO ATIVA
        sessao = verificar_sessao_ativa()

        if sessao:
            nome_ativo, email_ativo = sessao

            return no_update, html.Div([
                "Acesso negado, sistema com login ativo.",
                html.Br(),
                f"Usuário.: {nome_ativo}",
                html.Br(),
                f"Email.: {email_ativo}"
            ])

        # ✅ LIBERA ACESSO PÚBLICO
        return {
            'email': 'publico',
            'usuario': 'Público',
            'setor': 'visualizacao',
            'evento': 'UNIFICADO',
            'acesso': datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        }, ''

    # ✅ A PARTIR DAQUI É LOGIN NORMAL
    if not email or '@' not in email:
        return no_update, 'Informe um e-mail válido'

    email = email.lower().strip()

    # ✅ consulta já com email limpo
    dados_usuario = buscar_usuario_cadastro(email)

    # 🚫 BLOQUEIA SE NÃO ESTIVER CADASTRADO
    if not dados_usuario:
        return no_update, 'E-mail não cadastrado'

    usuario, setor, evento, grupos = dados_usuario

    # 🚫 BLOQUEIO PARA TÉCNICA
    if setor.lower() == 'técnica' and not grupo_selecionado:
        return no_update, '⚠️ Selecione um grupo antes de acessar'

    # ==========================================
    # 🔴 🔴 AQUI ENTRA O BLOQUEIO DE SESSÃO 🔴 🔴
    # ==========================================
    sessao = verificar_sessao_ativa()

    if sessao:
        nome_ativo, email_ativo = sessao

        return no_update, (
            "Acesso negado, sistema com login ativo.", html.Br(),
            f"Usuário.: {nome_ativo}", html.Br(),
            f"Email.: {email_ativo}"
        )

    # ==========================================
    # ✅ CRIA A SESSÃO (LOGIN LIBERADO)
    # ==========================================
    criar_sessao({
        "email": email,
        "usuario": usuario
    })

    # ✅ LOGIN NORMAL CONTINUA
    return {
        'email': email,
        'usuario': usuario,
        'setor': setor,
        'evento': evento,
        'grupos': grupos,
        'grupo_selecionado': grupo_selecionado,
        'acesso': datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    }, ''
@app.callback(
    Output('dados-originais-usuario', 'data'),
    Output('dados-usuario', 'data'),
    Input('usuario-store', 'data')
)
def inicializar_dados_usuario(usuario):
    #Carrega dados do usuário ao logar.
    if not usuario:
        raise PreventUpdate
    #Evita execução sem usuário.

    df_excel = carregar_df_sqlite()

    df_dev = carregar_devolutivas_usuario(
        usuario["email"]
    )

    df_view = aplicar_devolutivas(df_excel, df_dev)

    return (
        df_view.to_dict("records"),   # dados-originais-usuario
        df_view.to_dict("records")    # dados-usuario
    )

@app.callback(
    Output('tabela-editavel', 'data'),
    Input('dados-usuario', 'data')
)
def atualizar_tabela(dados):
    #Atualiza tabela na tela.
    if not dados:
        return []
    return dados

# ==================================================
# 🔐 CONTROLE DE PRIVILÉGIO POR SETOR
# ==================================================
@app.callback(
    Output('tabela-editavel', 'columns'),
    Input('usuario-store', 'data')
)
def controlar_colunas_por_setor(usuario):

    if not usuario:
        #raise dash.exceptions.PreventUpdate
        
        return [
                {
                    'name': c,
                    'id': c,
                    'presentation': 'input',
                    'editable': False
                }
                for c in df_controle.columns
            ]


    setor = usuario.get('setor', '').lower()

    COLUNAS_BLOQUEADAS = ['DATA', 'PROGRAMA', 'PRODUCT', 'PROJECT','NOVO CATÁLOGO','NÚMERO CATÁLOGO','FRONT_PAGE','TEVON','KOM','KOM - QUARTIL','KOM_DEVOLUTIVA','SOP','ME','COMMENTS']
    
   # ✅ VISUALIZAÇÃO TOTAL (visitante ou público)
    if setor == 'visualizacao':
        return [{
            'name': 'DATA',
            'id': 'DATA',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'PROGRAMA',
            'id': 'PROGRAMA',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'PROJECT',
            'id': 'PROJECT',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'PRODUCT',
            'id': 'PRODUCT',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'NOVO CATÁLOGO',
            'id': 'NOVO CATÁLOGO',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'NÚMERO CATÁLOGO',
            'id': 'NÚMERO CATÁLOGO',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'FRONT_PAGE',
            'id': 'FRONT_PAGE',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'TEVON',
            'id': 'TEVON',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'KOM',
            'id': 'KOM',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'KOM_DEVOLUTIVA',
            'id': 'KOM_DEVOLUTIVA',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'inicio_da_liberacao',
            'id': 'inicio_da_liberacao',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'Inicio_da_Liberacao_DEVOLUTIVA',
            'id': 'Inicio_da_Liberacao_DEVOLUTIVA',
            'presentation': 'dropdown',
            'editable': False
        },
        {
            'name': 'Início de Compras',
            'id': 'Início de Compras',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'Início de Compras - DEVOLUTIVA',
            'id': 'Início de Compras - DEVOLUTIVA',
            'presentation': 'dropdown',
            'editable': False
        },
        {
            'name': 'Fim de Compras',
            'id': 'Fim de Compras',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'Fim de Compras - DEVOLUTIVA',
            'id': 'Fim de Compras - DEVOLUTIVA',
            'presentation': 'dropdown',
            'editable': False
        },
        {
            'name': 'Fim de Liberação',
            'id': 'Fim de Liberação',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'Fim_de_Liberacao_DEVOLUTIVA',
            'id': 'Fim_de_Liberacao_DEVOLUTIVA',
            'presentation': 'dropdown',
            'editable': False
        },
        #------------------------
        #programação 
        {
            'name': 'Início de Programação/Disposição',
            'id': 'Início de Programação/Disposição',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'Início de Programação/Disposição - DEVOLUTIVA',
            'id': 'Início de Programação/Disposição - DEVOLUTIVA',
            'presentation': 'dropdown',
            'editable': False
        },
        {
            'name': 'Fim da Programação/Disposição',
            'id': 'Fim da Programação/Disposição',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'Fim da Programação/Disposição - DEVOLUTIVA',
            'id': 'Fim da Programação/Disposição - DEVOLUTIVA',
            'presentation': 'dropdown',
            'editable': False
        },
        #------------------------
        {
            'name': 'SOP',
            'id': 'SOP',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'ME',
            'id': 'ME',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'COMMENTS',
            'id': 'COMMENTS',
            'presentation': 'input',
            'editable': False
        }]
    
# ✅ CASO ESPECIAL: compras → apenas coluna DATA
    if setor == 'compras':
        return [{
            'name': 'DATA',
            'id': 'DATA',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'PROGRAMA',
            'id': 'PROGRAMA',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'PROJECT',
            'id': 'PROJECT',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'PRODUCT',
            'id': 'PRODUCT',
            'presentation': 'input',
            'editable': False
        },
         {
            'name': 'NOVO CATÁLOGO',
            'id': 'NOVO CATÁLOGO',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'NÚMERO CATÁLOGO',
            'id': 'NÚMERO CATÁLOGO',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'Início de Compras',
            'id': 'Início de Compras',
            'presentation': 'input',
            'editable': True
        },
        {
            'name': 'Início de Compras - DEVOLUTIVA',
            'id': 'Início de Compras - DEVOLUTIVA',
            'presentation': 'dropdown',
            'editable': True
        },
        {
            'name': 'Fim de Compras',
            'id': 'Fim de Compras',
            'presentation': 'input',
            'editable': True
        },
        {
            'name': 'Fim de Compras - DEVOLUTIVA',
            'id': 'Fim de Compras - DEVOLUTIVA',
            'presentation': 'dropdown',
            'editable': True
        },
        {
            'name': 'SOP',
            'id': 'SOP',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'ME',
            'id': 'ME',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'COMMENTS',
            'id': 'COMMENTS',
            'presentation': 'input',
            'editable': False
        }]
# ✅ CASO ESPECIAL: TÉCNICA → apenas coluna DATA
    if setor == 'técnica':

        grupo = usuario.get('grupo_selecionado')
        grupo_idx = grupo.replace("Grupo", "") if grupo else None

        colunas = [
            {
                'name': 'DATA',
                'id': 'DATA',
                'editable': False
            },
            {
                'name': 'PROGRAMA',
                'id': 'PROGRAMA',
                'editable': False
            },
            {
                'name': 'PROJECT',
                'id': 'PROJECT',
                'editable': False
            },
            {
                'name': 'PRODUCT',
                'id': 'PRODUCT',
                'editable': False
            },
            {
                'name': 'inicio_da_liberacao',
                'id': 'inicio_da_liberacao',
                'editable': True
            },
            #{
            #    'name': 'Inicio_da_Liberacao_DEVOLUTIVA',
            #    'id': 'Inicio_da_Liberacao_DEVOLUTIVA',
            #    'presentation': 'dropdown',
            #    'editable': True
            #}
        ]

        # ✅ adiciona SOMENTE o grupo selecionado
        if grupo_idx is not None:

            # 🔹 INÍCIO
            colunas.append({
                'name': f'Inicio_da_Liberacao_G{grupo_idx}',
                'id': f'Inicio_da_Liberacao_G{grupo_idx}',
                'presentation': 'dropdown',
                'editable': True
            })

            # 🔹 FIM (NOVO)
            colunas.append({
                'name': f'Fim_de_Liberacao_DEVOLUTIVA_G{grupo_idx}',
                'id': f'Fim_de_Liberacao_DEVOLUTIVA_G{grupo_idx}',
                'presentation': 'dropdown',
                'editable': True
            })

            return colunas
    # ✅ CASO ESPECIAL: TÉCNICA → apenas coluna DATA
    if setor == 'programação/disposição':
        return [{
            'name': 'DATA',
            'id': 'DATA',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'PROGRAMA',
            'id': 'PROGRAMA',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'PROJECT',
            'id': 'PROJECT',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'PRODUCT',
            'id': 'PRODUCT',
            'presentation': 'input',
            'editable': False
        },
         {
            'name': 'NOVO CATÁLOGO',
            'id': 'NOVO CATÁLOGO',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'NÚMERO CATÁLOGO',
            'id': 'NÚMERO CATÁLOGO',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'Início de Programação/Disposição',
            'id': 'Início de Programação/Disposição',
            'presentation': 'input',
            'editable': True
        },
        {
            'name': 'Início de Programação/Disposição - DEVOLUTIVA',
            'id': 'Início de Programação/Disposição - DEVOLUTIVA',
            'presentation': 'dropdown',
            'editable': True
        },
        {
            'name': 'Fim da Programação/Disposição',
            'id': 'Fim da Programação/Disposição',
            'presentation': 'input',
            'editable': True
        },
        {
            'name': 'Fim da Programação/Disposição - DEVOLUTIVA',
            'id': 'Fim da Programação/Disposição - DEVOLUTIVA',
            'presentation': 'dropdown',
            'editable': True
        },
        {
            'name': 'SOP',
            'id': 'SOP',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'ME',
            'id': 'ME',
            'presentation': 'input',
            'editable': False
        },
        {
            'name': 'COMMENTS',
            'id': 'COMMENTS',
            'presentation': 'input',
            'editable': False
        }]
    
    if setor == 'gestão':
        return [
            {
                'name': c,
                'id': c,
                'presentation': (
                    'dropdown' if c in COLUNAS_DROPDOWN_SIM_NAO else 'input'
                ),
                'editable': True  # ✅ gestão edita tudo
            }
            for c in df_controle.columns
        ]

    colunas = []
    for c in df_controle.columns:
        colunas.append({
            'name': c,
            'id': c,
            'presentation': (
                'dropdown'
                if c in COLUNAS_DROPDOWN_SIM_NAO
                else 'input'
            ),
            # ===== INÍCIO DO BLOCO DE PRIVILÉGIO =====
            'editable': not (
                setor != 'gestão' and c in COLUNAS_BLOQUEADAS
            )
            # ===== FIM DO BLOCO DE PRIVILÉGIO =====
        })
    
    return colunas

#--------------------------------------
#7. Troca de tela
@app.callback(
    Output('login-container', 'style'),
    Output('dashboard-container', 'style'),
    Output('usuario-logado', 'children'),
    Output('setor-usuario', 'children'),
    Output('evento-usuario', 'children'),  # 👈 NOVO
    Output('grupos-usuario', 'children'),
    Output('data-acesso', 'children'),
    Input('usuario-store', 'data'),
    State('login-container', 'style')
)
def trocar_tela(usuario, login_style):

    if not usuario:
        login_style['display'] = 'flex'
        return login_style, {'display': 'none'}, '', '', '', '', ''

    login_style['display'] = 'none'

    # ✅ AQUI
   
    grupo_sel = usuario.get('grupo_selecionado')
    grupos_txt = grupo_sel if grupo_sel else "Sem grupo selecionado"


    return (
        login_style,
        {'display': 'block'},
        f"Usuário: {usuario['usuario']}",
        f"Setor: {usuario['setor']}",
        f"Evento: {usuario.get('evento', '')}",
        f"Grupos: {grupos_txt}",   # ✅ aparece no ticket
        f"Acesso em: {usuario['acesso']}"
    )
#Filtra dados com base nos inputs.
def filtrar_tabela(programas, projetos, datas, limpar, dados_originais):

    if not dados_originais:
        raise dash.exceptions.PreventUpdate

    df = pd.DataFrame(dados_originais)

    # ✅ LIMPAR FILTROS (funciona para TODOS!)
    if ctx.triggered_id == 'btn-limpar-filtros-tabela':
        return df.to_dict('records'), None, None, None

    # ✅ aplicar filtros normalmente
    if programas:
        df = df[df['PROGRAMA'].isin(programas)]

    if projetos:
        df = df[df['PROJECT'].isin(projetos)]

    if datas and 'DATA' in df.columns:
        df = df[df['DATA'].fillna('').astype(str).isin(datas)]

    return (
        df.to_dict('records'),
        programas,
        projetos,
        datas
    )

#--------------------------------------
# 8.1.  GRAVAÇÃO
@app.callback(
    Output('msg-salvar', 'children'),
    Output('dados-usuario', 'data'),
    Output('dados-originais-usuario', 'data'),
    Input('btn-salvar', 'n_clicks'),
    State('tabela-editavel', 'data'),
    State('dados-originais-usuario', 'data'),  # ✅ ADICIONE ISSO
    #State('tabela-editavel', 'data_previous'),
    State('usuario-store', 'data'),
    prevent_initial_call=True
)

#def salvar_no_excel(n_clicks, data_atual, data_anterior, usuario):
def salvar_alteracoes(n_clicks, data_atual, dados_anteriores_usuario, usuario):
    if not usuario or not data_atual:
        return "❌ Nada para salvar", no_update

    
    # ✅ base neutra SEM devolutivas
    #base_neutra = carregar_df('Controle do Projeto').to_dict('records')

    #devolutivas = detectar_devolutivas(data_atual,base_neutra)
    
    # ✅ compara com o último estado exibido ao usuário
    devolutivas = detectar_devolutivas(
        data_atual,
        dados_anteriores_usuario
    )


    # 2️⃣ Salva devolutivas no SQLite (fonte permanente)
    if devolutivas:
        salvar_devolutivas_sqlite(
            devolutivas,
            usuario={
                "email": usuario["email"],
                "nome": usuario["usuario"],
                "setor": usuario["setor"]
            },
            data_ref=data_atual
        )

    # 4️⃣ 🔁 Reconstrói a visão do usuário A PARTIR DO SQLITE
    df_base = carregar_df_sqlite()
    df_dev = carregar_devolutivas_usuario(usuario["email"])
    df_view = aplicar_devolutivas(df_base, df_dev)

    # 5️⃣ Retorna mensagem + dados isolados do usuário
    return (
        f"✅ Alterações salvas por {usuario['usuario']} em {datetime.now().strftime('%H:%M:%S')}",
        df_view.to_dict("records"), df_view.to_dict("records")
    )

def detectar_alteracoes(atual, anterior):
    alteracoes = []

    for i, (linha_nova, linha_antiga) in enumerate(zip(atual, anterior)):
        for coluna in linha_nova:
            if linha_nova[coluna] != linha_antiga.get(coluna):
                alteracoes.append({
                    "linha": i,
                    "coluna": coluna,
                    "antes": linha_antiga.get(coluna),
                    "depois": linha_nova[coluna]
                })
    return alteracoes
def registrar_log(alteracoes, usuario, data_ref):
    
    log_df = pd.read_excel(ARQUIVO_EXCEL, sheet_name="LOG", engine="openpyxl")

    novos_logs = []
    for alt in alteracoes:
        novos_logs.append({
            "data_hora": datetime.now(),
            "email": usuario["email"],
            "nome": usuario["nome"],
            "setor": usuario["setor"],
            "projeto": data_ref[alt["linha"]]["PROJECT"],
            "coluna": alt["coluna"],
            "valor_anterior": alt["antes"],
            "valor_novo": alt["depois"],
        })

    log_df = pd.concat([log_df, pd.DataFrame(novos_logs)], ignore_index=True)

    with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        log_df.to_excel(writer, sheet_name="LOG", index=False)

def detectar_devolutivas(atual, anterior):
    alteracoes = []

    if not anterior:
        return alteracoes

    for i, (linha_nova, linha_antiga) in enumerate(zip(atual, anterior)):
        for coluna in COLUNAS_DEVOLUTIVA:
            if coluna in linha_nova:

                novo = linha_nova.get(coluna)
                antigo = linha_antiga.get(coluna)

                # ✅ regra geral
                if novo != antigo:

                    # ✅ mantém regra original
                    if novo in ['SIM', 'NÃO']:
                        alteracoes.append({
                            "linha": i,
                            "coluna": coluna,
                            "valor": novo
                        })

                    # ✅ NOVO BLOCO (ESSENCIAL)
                    if coluna == "Inicio_da_Liberacao_DEVOLUTIVA" and novo in ['FEITO', 'PENDENTE']:
                        alteracoes.append({
                            "linha": i,
                            "coluna": coluna,
                            "valor": novo
                        })                    
                    
                    # ✅ NOVO BLOCO PARA FIM (ESPELHO DO INICIO)
                    if coluna == "Fim_de_Liberacao_DEVOLUTIVA" and novo in ['FEITO', 'PENDENTE']:
                        alteracoes.append({
                            "linha": i,
                            "coluna": coluna,
                            "valor": novo
                        })
                #----------------------
                # ✅ NOVA REGRA: gestão alterou G para NÃO
                #if (
                #    coluna.startswith("Inicio_da_Liberacao_G") or
                #    coluna.startswith("Fim_de_Liberacao_DEVOLUTIVA_G")
                #):
                #    if novo == "NÃO":
                #        alteracoes.append({
                #            "linha": i,
                #            "coluna": "Inicio_da_Liberacao_DEVOLUTIVA",
                #            "valor": "PENDENTE"
                #        })
                #        alteracoes.append({
                #            "linha": i,
                #            "coluna": "Fim_de_Liberacao_DEVOLUTIVA",
                #            "valor": "PENDENTE"
                #        })



    return alteracoes

def salvar_devolutivas_sqlite(devolutivas, usuario, data_ref):
    conn = sqlite3.connect(DB_SQLITE)
    cur = conn.cursor()

    projetos_afetados = set()
    colunas_afetadas = set()

    for d in devolutivas:
        projeto = data_ref[d["linha"]]["PROJECT"]
        coluna = d["coluna"]
        valor = d["valor"]

        #--------------------
        # 🔥 REGRA: gestão marcou NÃO em coluna G
        if (
            usuario["setor"].lower() == "gestão"
            and (
                coluna.startswith("Inicio_da_Liberacao_G") or
                coluna.startswith("Fim_de_Liberacao_DEVOLUTIVA_G")
            )
            and valor == "NÃO"
        ):

            valor_inicio = "PENDENTE"
            valor_fim = "PENDENTE"

            # ✅ INSERT log
            cur.execute("""
                INSERT INTO devolutivas
                (data_hora, email, nome, setor, projeto, coluna, valor,
                Inicio_da_Liberacao_DEVOLUTIVA, Fim_de_Liberacao_DEVOLUTIVA, status_projeto)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                usuario["email"],
                usuario["nome"],
                usuario["setor"],
                projeto,
                coluna,
                valor,
                valor_inicio,
                valor_fim,
                "PENDENTE"
            ))

            projeto = str(projeto).strip()
            # 🔥 Atualiza tabela principal (tarefas)
            conn_tarefas = sqlite3.connect(r".\\base_sql\\demandas.db")
            cur_tarefas = conn_tarefas.cursor()

            coluna_map = {
                "KOM - DEVOLUTIVA": "KOM_DEVOLUTIVA"
            }

            coluna_db = coluna_map.get(coluna, coluna)

            cur_tarefas.execute(f"""
                UPDATE tarefas
                SET {coluna_db} = ?
                WHERE PROJECT = ?
            """, (valor, projeto))

            conn_tarefas.commit()
            conn_tarefas.close()

            # ✅ UPDATE no registro real
            cur.execute("""
                UPDATE devolutivas
                SET Inicio_da_Liberacao_DEVOLUTIVA = ?,
                    Fim_de_Liberacao_DEVOLUTIVA = ?
                WHERE projeto = ?
            """, (valor_inicio, valor_fim, projeto))

            # ✅ Atualiza estrutura em memória
            data_ref[d["linha"]]["Inicio_da_Liberacao_DEVOLUTIVA"] = valor_inicio
            data_ref[d["linha"]]["Fim_de_Liberacao_DEVOLUTIVA"] = valor_fim

            # ✅ Atualiza o objeto atual também
            d["Inicio_da_Liberacao_DEVOLUTIVA"] = valor_inicio
            d["Fim_de_Liberacao_DEVOLUTIVA"] = valor_fim

            
        # 🔥 ESSENCIAL: impedir o INSERT normal
            continue

        #--------------------

        projets_test = projeto
        # ✅ 1. captura o valor da linha
        valor_inicio = data_ref[d["linha"]].get("Inicio_da_Liberacao_DEVOLUTIVA")

        # ✅ DEBUG (muito importante agora)
        print("VALOR CAPTURADO:", valor_inicio)

        # ✅ 2. tratamento
        if valor_inicio in [None, '', 'NaT']:
            valor_inicio = None
        
        valor_fim = data_ref[d["linha"]].get("Fim_de_Liberacao_DEVOLUTIVA")
        
        if valor_fim in [None, '', 'NaT']:
            valor_fim = None

        # ✅ 3. INSERT usando valor tratado
        cur.execute("""
            INSERT INTO devolutivas
            (data_hora, email, nome, setor, projeto, coluna, valor, Inicio_da_Liberacao_DEVOLUTIVA,
                    Fim_de_Liberacao_DEVOLUTIVA)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            usuario["email"],
            usuario["nome"],
            usuario["setor"],
            projeto,
            coluna,
            d["valor"],
            valor_inicio,   # ✅ aqui usa a variável
            valor_fim   # ✅ NOVO
        ))

        # Guarda para atualizar somente o necessário
        projetos_afetados.add(projeto)
        colunas_afetadas.add(coluna)

    # 🔁 UPDATE do status_projeto (sua query)
    # 🔁 UPDATE do status_projeto (regra por setor)
    for projeto in projetos_afetados:
        for coluna in colunas_afetadas:

            # ✅ REGRA PARA COMPRAS
            if usuario["setor"] == "compras":
                cur.execute("""
                    UPDATE devolutivas
                    SET status_projeto = (
                        SELECT
                            CASE
                                WHEN COUNT(DISTINCT f2.nome) >= 1 THEN 'FEITO'
                                ELSE 'PENDENTE'
                            END
                        FROM devolutivas f2
                        WHERE f2.projeto = ?
                        AND f2.coluna = ?
                        AND f2.setor == 'Compras'
                        AND f2.valor = 'SIM'
                    )
                    WHERE projeto = ?
                    AND coluna = ?
                """, (projeto, coluna, projeto, coluna))
            elif usuario["setor"] == "Programação/Disposição":
                cur.execute("""
                    UPDATE devolutivas
                    SET status_projeto = (
                        SELECT
                            CASE
                                WHEN COUNT(DISTINCT f2.nome) >= 1 THEN 'FEITO'
                                ELSE 'PENDENTE'
                            END
                        FROM devolutivas f2
                        WHERE f2.projeto = ?
                        AND f2.coluna = ?
                        AND f2.setor == 'Programação/Disposição'
                        AND f2.valor = 'SIM'
                    )
                    WHERE projeto = ?
                    AND coluna = ?
                """, (projeto, coluna, projeto, coluna))
#Programação/Disposição
            elif usuario["setor"] == "gestão":
                cur.execute("""
                    UPDATE devolutivas
                    SET status_projeto = (
                        SELECT
                            CASE
                                WHEN COUNT(DISTINCT f2.nome) >= 1 THEN 'FEITO'
                                ELSE 'PENDENTE'
                            END
                        FROM devolutivas f2
                        WHERE f2.projeto = ?
                        AND f2.coluna = ?
                        AND f2.setor == 'gestão'
                        AND f2.valor = 'SIM'
                    )
                    WHERE projeto = ?
                    AND coluna = ?
                """, (projeto, coluna, projeto, coluna))
#Programação/Disposição
            # ✅ REGRA PADRÃO (ex: Técnica / demais)
        else:
                    # 🔹 1. ATUALIZA INICIO E FIM PRIMEIRO
                    cur.execute("""
                        UPDATE devolutivas
                        SET 
                            Inicio_da_Liberacao_DEVOLUTIVA = (
                                SELECT
                                    CASE
                                        WHEN COUNT(DISTINCT d2.email) >= 4
                                        AND COUNT(DISTINCT d2.coluna) >= 10
                                        THEN 'FEITO'
                                        ELSE 'PENDENTE'
                                    END
                                FROM devolutivas d2
                                WHERE d2.projeto = devolutivas.projeto
                            ),

                            Fim_de_Liberacao_DEVOLUTIVA = (
                                SELECT
                                    CASE
                                        WHEN (
                                            SELECT COUNT(DISTINCT d3.coluna)
                                            FROM devolutivas d3
                                            WHERE d3.projeto = devolutivas.projeto
                                            AND d3.coluna LIKE 'Fim_de_Liberacao_DEVOLUTIVA_G%'
                                            AND d3.valor = 'SIM'
                                        ) = 10
                                        THEN 'FEITO'
                                        ELSE 'PENDENTE'
                                    END
                            )
                        WHERE projeto = ?
                    """, (projeto,))

                    # 🔹 2. AGORA CALCULA status_projeto CORRETAMENTE
                    cur.execute("""
                        UPDATE devolutivas
                        SET status_projeto = 
                            CASE
                                WHEN Inicio_da_Liberacao_DEVOLUTIVA = 'FEITO'
                                AND Fim_de_Liberacao_DEVOLUTIVA = 'FEITO'
                                THEN 'FEITO'
                                ELSE 'PENDENTE'
                            END
                        WHERE projeto = ?
                    """, (projeto,))
    conn.commit()
    conn.close()
#--------------------------------------
# ==================================================
# 9. GANTT (INALTERADO)
# ==================================================
@app.callback(
    Output('grafico-gantt', 'figure'),
    Input('filtro-programa', 'value'),
    Input('filtro-projeto', 'value'),
    Input('filtro-fase', 'value'),
    Input('filtro-data', 'start_date'),
    Input('filtro-data', 'end_date')
)
def atualizar_gantt(programas, projetos, fases, ini, fim):

    # ✅ 1. Carrega direto do banco
    df = carregar_df_sqlite()

    if df.empty:
        return px.timeline()

    # ✅ 2. Converter datas
    for col in COLUNAS_DATA:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')

    # ✅ 3. Aplicar filtros
    if programas:
        df = df[df['PROGRAMA'].isin(programas)]

    if projetos:
        df = df[df['PROJECT'].isin(projetos)]

    if ini and fim:
        ini = pd.to_datetime(ini)
        fim = pd.to_datetime(fim)

        df = df[
            (df['inicio_da_liberacao'] <= fim) &
            (df['Fim da Programação/Disposição'] >= ini)
        ]

    # ✅ 4. Gerar Gantt
    return criar_gantt(df, fases)

@app.callback(
    Output('filtro-programa', 'value'),
    Output('filtro-projeto', 'value'),
    Output('filtro-fase', 'value'),
    Output('filtro-data', 'start_date'),
    Output('filtro-data', 'end_date'),
    Input('btn-limpar-filtros', 'n_clicks'),
    prevent_initial_call=True
)
def limpar_filtros_cronograma(n_clicks):
    return None, None, None, None, None

def limpar_filtros_cronograma(n_clicks):
    return None, None, None, None, None

#--------------------------------------
# ✅ 9.1. OVERVIEW – TOTAL SOP 2026  ← COLE AQUI

@app.callback(
    Output('overview-cards', 'children'),
    Input('tabela-editavel', 'data')
)
def atualizar_overview(dados):

    if not dados:
        return []

    df = pd.DataFrame(dados)

    # garante datetime
    df['SOP_DT'] = pd.to_datetime(
        df['SOP'],
        dayfirst=True,
        errors='coerce'
    )

    # total SOP 2026
    total_sop_2026 = df[df['SOP_DT'].dt.year == 2026].shape[0]
    total_sop_2027 = df[df['SOP_DT'].dt.year == 2027].shape[0]
    total_sop_2028 = df[df['SOP_DT'].dt.year == 2028].shape[0]

    return [
        card_resumo(
            titulo='Total de SOP – 2026',
            valor=total_sop_2026,
            cor="#18AF68"
        ),
        card_resumo(
            titulo='Total de SOP – 2027',
            valor=total_sop_2027,
            cor="#4084C8"
        ),
        card_resumo(
            titulo='Total de SOP – 2028',
            valor=total_sop_2028,
            cor="#C62B2B"
        )
    ]

#--------------------------------------
# ✅ 9.2. LOG
@app.callback(
    Output('tabela-log', 'data'),
    Output('tabela-log', 'columns'),
    Input('btn-atualizar-log', 'n_clicks'),
    Input('usuario-store', 'data'),
    prevent_initial_call=True
)
def carregar_log_sqlite(n_clicks, usuario):

    if not usuario:
        raise PreventUpdate

    conn = sqlite3.connect(DB_SQLITE)

    query = """
        SELECT
            id,
            data_hora,
            email,
            nome,
            setor,
            projeto,
            coluna,
            valor,
            Inicio_da_Liberacao_DEVOLUTIVA,  -- ✅ NOVA COLUNA
            Fim_de_Liberacao_DEVOLUTIVA,  -- ✅ NOVA COLUNA
            status_projeto
        FROM devolutivas
        ORDER BY data_hora DESC
    """

    df_log = pd.read_sql_query(query, conn)
    conn.close()

    if df_log.empty:
        return [], []

    columns = [
        {'name': c.replace('_', ' ').upper(), 'id': c}
        for c in df_log.columns
    ]
 

    return df_log.to_dict('records'), columns


def ajustar_tabela_devolutivas():
    conn = sqlite3.connect(DB_SQLITE)
    cur = conn.cursor()

    colunas_novas = [
        "status_projeto TEXT",
        "Inicio_da_Liberacao_DEVOLUTIVA TEXT",   # ✅ AQUI
        "Fim_de_Liberacao_DEVOLUTIVA TEXT"   # ✅ AQUI
    ]

    for col in colunas_novas:
        try:
            cur.execute(f"ALTER TABLE devolutivas ADD COLUMN {col}")
            conn.commit()
        except sqlite3.OperationalError:
            # coluna já existe
            pass

    conn.close()


#--------------------------------------
# ✅ 9.3. BOTÕES
@app.callback(
    Output('btn-salvar', 'children'),
    Input('usuario-store', 'data')
)
def atualizar_nome_botao(usuario):
    if not usuario:
        raise PreventUpdate

    setor = usuario.get('setor', '').lower()

    if setor == 'técnica':
        return 'Salvar Alterações - Técnica'

    if setor == 'compras':
        return 'Salvar Alterações - Compras'

    if setor == 'programação/disposição':
        return 'Salvar Alterações - Programação'

    if setor == 'gestão':
        return 'Salvar Alterações - gestão'

    return 'Salvar Alterações'

#--------------------------------------
# 10. RUN

#print("Abrindo BD:", Path(r".\base_sql\usuarios.db").resolve())

if __name__ == "__main__":
    #resetar_banco()
    criar_banco_sqlite()
    encerrar_sessao()
    ajustar_tabela_devolutivas()  # ✅ ESSENCIAL
    threading.Timer(1, abrir_navegador).start()
    app.run(debug=False)
