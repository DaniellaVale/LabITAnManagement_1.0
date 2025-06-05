'''Descrição da aplicação: este aplicativo faz o gerenciamento do almoxarifado do LabITAn-DQA-IQ-UFJ e mostra lista de professores e alunos
Autores: Daniella Lopez Vale e Maiara Oliveira Salles.
O Aplicativo foi desenvolvido por Daniella L. Vale usando o auxilio de ferramentas de Inteligência Artificial generativa (CHAT-GPT e DeepSeek) e Maiara O. Salles contribuiu com ideias sobre a arquitetura da aplicação.'''
import flet as ft
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import webbrowser
import pandas as pd
import sys
import time

'''Função para alterar entre as páginas'''
def ir_para_pagina(page, funcao_pagina):
    page.clean()  # Limpa os controles da página atual
    funcao_pagina(page)  # Carrega a nova página
    page.update()  # Atualiza a interface

# Variável global para armazenar as permissões do usuário
user_permissions = {
    'logged_in': False,
    'email': '',
    'can_edit_controlled': False
}

# Configurações diretas no código (substituindo o .env)
SHEET_ID_LOGIN = "1hJKxE_0HkYVCfpX8Nzvitub84b9ynU5KlWtf5xeNGJI"  # Substitua pelo ID real
SHEET_ID_ALUNOS = "1mNqkB6Vnh33ByGNBRqQzOFV5sfaBwd10exIycGrfdmU"  # Substitua pelo ID real
SHEET_ID_PROFESSORES = "1LuAd9-0jCbds1OmY5t-p4S9zAdk_eHm2PV_Mg5xt8EQ"  # Substitua pelo ID real
SHEET_ID_N_CONT = "1jGjRi5g3GCWck68fTIOLdwQsKdy_lfSqxKMgI0tx9dI"  # Substitua pelo ID real
SHEET_ID_CONT = "1uwFx0n4buf2fVmsyfgZo4zEofLbRiDIuNyb0rwhg4to"  # Substitua pelo ID real
SHEET_ID_MODIFICAR = "1I-7OG4Hd03cj514QoOvmlsvFgq6Sss_ihTzcYVMmQzI"  # Substitua pelo ID real
WORKSHEET_NAME = "Sheet1"  # Ou o nome da aba que você usa

# Credenciais JSON diretamente no código (substitua pelo seu JSON real)
# Substitua pelo caminho do seu arquivo JSON (ex.: "credenciais.json")
CAMINHO_CREDENCIAIS = {
  "type": "service_account",
  "project_id": "api-reagentes",
  "private_key_id": "190d7c597b5ca724ddbe0909eb50ab290837c579",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDY+MlvQynzE3LG\neKh4538IVTgHCD0N3Y0YPmOpFg0RJnpLMtBYLOR/nY6HfwNABcCX+PKXMSlm7mYZ\n10fXy3oZMGcc+Fuk8sazJpGB05NXIjV/PZWuQV5MyHDwOaVeuFZIenpyaF2djvVM\nb6YF4jroT9Z2hJ6W7jhLTYq7XUggUIAOlO9GHAjij0ktfo8iYsInk+Emvigh2UXT\nHJi/AlGGx30nz7EnbRn1tDsJYnTNDbzXu/L+RzqcijcRI0M6jR8hPlQ7a9likb14\nEBe9AeHT+vA2yTJNYMTaoB4bE2qf4G7/cLy05GifQUC2Hg6/G4nUfC4zo31QoGBP\nHfDf/RQDAgMBAAECggEADIDFFkTaQyOmn9OOb4AHlfvqCp2y647E0PqL+WHaNKEV\n3CnFdFh4xvZoNVxgXCQdSjbtv0CsRz4v9SwcaXJ43QmWudil/QY5m+XVpLhxokxa\nXEa4zouQEduB9EzCraxeR+3LeJIDDnQjpEbRpT9wJriuffugvH5lyBEatO3oU9RG\nxhSUl9dvcEwknxMyYDKPRX9DxwHhY6rRMhcBrJ49n1uf6IQe9fOMNu59bozbCCbZ\n5IrEnZwMKXN7vzvqO0esWJZdaEKZZNKy1Z6VmnMGLAO5bLPMRqY+BWoY/yjHJj9c\nKe3dbmyOK78tAoCp6nq/XRCyM+CYsHbNZrREWGqAAQKBgQD+QqIAM/4xz37KdK56\nQoX9GglwDRJh8cxOVe9lKvZIrSFqFM3fG/lw9tTpaFa+oH3JRCWVaKnyg3jG5XvM\nwdkXIXHQmGkViF6pl2nJA4nRX7PB7V6rxp1ssPuC40GAKLhglg9lu3igJXjuG51r\nLwBSzVdd1qz2cfFcRQLdbdEQAQKBgQDadNa/SCZpH66Cr+qER2Bjg6CDCRihq/0k\noZzQK71AIgmYi7QdXDME+m2bWU7tCDem4iy76t0wy0X/dgJ34b1kSYWIVuVehZiM\nSISzDkGZcqOLONUE8lpxcPI5DkDQE1IxB+bJYMlToVsdsaFNjTvSNpcX3i/kztY0\nlq2s1EnkAwKBgQCcb17BBTxGZUW7RqL68ecCXHymBkTjIiPzpofOFOrGuE6wt/Aa\nb1m/mP5SRTHpw1Dg/h6pmGXHogAzT4ol5rastpUSJFOzPd4QNeqOFLE8ssckb+kp\ngt/kuddlJnFsaqFWO71peDi1P5jx1ue5xIdMaq5wO97bGivH+2XR2vkgAQKBgAMw\no5YlepIcaVL1OKp31Ft/p49iSZ7KwSaQyZZsnRXbqWI14ApxtzkCYylak4F4lj90\nnAyecF5vCXWihoSzoi1duXp1MmI/9ytNP8rRkXmpJ+Q3jzzEQTfY22Cj6aRgM9oN\ncHxOUoJLH+Z+GonkXxRBwdESaIah0pTwAlc8vlt7AoGAV4CqFQqNqm9PosOctw82\nyyg8n+l9N3D/znGOZbTxeCOLNWgw14PHmVjZRnUnXYMOQFVZ2DLa1YfpNJjLhlNX\ni53lZbiKS/br2qze/aoVU1uNJnn0IvcnaCyI5EN6BUeiYsf5SJ9Xloxb+bZUe3Ru\nCHzQW3ZvIGg63ACPzZ8+vuQ=\n-----END PRIVATE KEY-----\n",
  "client_email": "acesso-planilha@api-reagentes.iam.gserviceaccount.com",
  "client_id": "102197938478575606754",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/acesso-planilha%40api-reagentes.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"
}

'''página de login'''
# Configuração inicial de logging
print("\n=== INICIALIZANDO SISTEMA ===")

# Verificação crítica das variáveis
if not SHEET_ID_LOGIN:
    print("❌ ERRO CRÍTICO: SHEET_ID_LOGIN não está definido")
    sys.exit(1)

print("\n=== CONFIGURAÇÕES ===")
print(f"Planilha ID Login: {SHEET_ID_LOGIN}")
print(f"Planilha ID Alunos: {SHEET_ID_ALUNOS}")
print(f"Planilha ID Professores: {SHEET_ID_PROFESSORES}")
print(f"Planilha ID Não Controlados: {SHEET_ID_N_CONT}")
print(f"Planilha ID Controlados: {SHEET_ID_CONT}")
print(f"Planilha ID Modificar: {SHEET_ID_MODIFICAR}")
print(f"Aba da planilha: {WORKSHEET_NAME}")

# Escopo necessário
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

# Autenticação com tratamento de erros
try:
    print("\n🔑 Autenticando com Google Sheets...")
    creds = ServiceAccountCredentials.from_json_keyfile_name(CAMINHO_CREDENCIAIS, scope)
    client = gspread.authorize(creds)
    print("✅ Autenticação bem-sucedida")
except Exception as auth_error:
    print(f"❌ Falha na autenticação: {auth_error}")
    sys.exit(1)

def carregar_dados():
    """Carrega os dados da planilha com verificação completa"""
    try:
        print(f"\n📂 Acessando planilha ID: {SHEET_ID_LOGIN}")
        sheet = client.open_by_key(SHEET_ID_LOGIN)
        
        # Lista todas as abas disponíveis
        worksheets = sheet.worksheets()
        print(f"Abas disponíveis: {[ws.title for ws in worksheets]}")
        
        # Acessa a aba específica
        try:
            worksheet = sheet.worksheet(WORKSHEET_NAME)
        except gspread.WorksheetNotFound:
            print(f"❌ Aba '{WORKSHEET_NAME}' não encontrada!")
            print(f"Abas disponíveis: {[ws.title for ws in worksheets]}")
            return None
            
        # Carrega os dados
        data = worksheet.get_all_records()
        if not data:
            print("⚠️ Planilha vazia ou sem registros")
            return pd.DataFrame()
            
        df = pd.DataFrame(data)
        
        # Debug: Mostra estrutura completa
        print("\n📊 ESTRUTURA DA PLANILHA:")
        print(f"Total de linhas: {len(df)}")
        print(f"Colunas: {df.columns.tolist()}")
        print("\nPrimeiros registros:")
        print(df.head())
        
        return df
        
    except Exception as e:
        print(f"❌ ERRO ao carregar dados: {e}")
        return None

def verificar_login(email, senha):
    """Verificação robusta de login com retorno de permissões"""
    print(f"\n🔐 Tentativa de login - Email: {email}")
    
    df = carregar_dados()
    if df is None:
        return (False, False)  # (login_valido, tem_permissao)
        
    # Verifica colunas necessárias
    required_columns = {'E-mail', 'Senha', 'Permissão de acesso'}
    if not required_columns.issubset(df.columns):
        print(f"❌ Colunas faltando: {required_columns - set(df.columns)}")
        print(f"Colunas disponíveis: {df.columns.tolist()}")
        return (False, False)
        
    # Normalização
    email = str(email).strip().lower()
    senha = str(senha).strip()
    
    print(f"🔍 Procurando: Email='{email}' | Senha='{senha}'")
    
    # Converte colunas para string
    df['E-mail'] = df['E-mail'].astype(str).str.strip().str.lower()
    df['Senha'] = df['Senha'].astype(str).str.strip()
    df['Permissão de acesso'] = df['Permissão de acesso'].astype(str).str.strip().str.upper()
    
    # Busca exata
    match = df[(df['E-mail'] == email) & (df['Senha'] == senha)]
    
    if not match.empty:
        print("✅ Login válido - Registro encontrado:")
        print(match.iloc[0].to_dict())
        permissao = match.iloc[0]['Permissão de acesso'] == 'SIM'
        return (True, permissao)
    else:
        print("❌ Login inválido - Nenhum registro correspondente")
        return (False, False)

def login(page: ft.Page):
    page.title = "Sistema de Login"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.window_resizable = True
    page.window_maximized = True
    page.scroll = ft.ScrollMode.AUTO
    page.bgcolor = ft.colors.WHITE
    
    # Interface
    email_input = ft.TextField(
        label="E-mail",
        width=300,
        autofocus=True
    )
    
    senha_input = ft.TextField(
        label="Senha (11 dígitos)",
        width=300,
        password=True,
        can_reveal_password=True,
        max_length=11
    )
    
    mensagem = ft.Text("", color=ft.colors.RED)
    
    def on_login(e):
        email = email_input.value
        senha = senha_input.value
        
        if not all([email, senha]):
            mensagem.value = "Preencha todos os campos!"
        elif len(senha) != 11 or not senha.isdigit():
            mensagem.value = "Senha deve ter 11 dígitos numéricos!"
        else:
            login_valido, tem_permissao = verificar_login(email, senha)
            if login_valido:
                # Atualiza as permissões globais
                global user_permissions
                user_permissions = {
                    'logged_in': True,
                    'email': email.lower().strip(),
                    'can_edit_controlled': tem_permissao
                }
                
                mensagem.value = "✅ Login bem-sucedido!"
                mensagem.color = ft.colors.GREEN
                page.update()
                
                # Pequeno delay para mostrar a mensagem
                time.sleep(1)
                
                # Navega para a página inicial
                ir_para_pagina(page, pagina_inicial)
            else:
                mensagem.value = "❌ Credenciais inválidas"
                mensagem.color = ft.colors.RED
                
        mensagem.update()
    
    page.add(
        ft.Column([
            ft.Text("Login do Sistema", size=24, weight="bold"),
            email_input,
            senha_input,
            ft.ElevatedButton(
                "Entrar", 
                on_click=on_login,
                icon=ft.icons.LOCK_OPEN
            ),
            mensagem
        ], spacing=20)
    )

'''Página principal'''
def pagina_inicial(page):
    page.bgcolor = ft.colors.WHITE
    page.scroll = "adaptive"
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.vertical_alignment = ft.MainAxisAlignment.CENTER

    logo = ft.Image(src="logo.tif", width=300, height=150)

    botao_alunos = ft.ElevatedButton(
        text="Alunos",
        icon=ft.icons.GROUP,
        width=250,
        on_click=lambda e: ir_para_pagina(page, alunos),
        style=ft.ButtonStyle(
            shape=ft.RoundedRectangleBorder(radius=8)
        ))
    
    botao_professores = ft.ElevatedButton(
        text="professores",
        icon=ft.icons.SCHOOL,
        width=250,
        on_click=lambda e: ir_para_pagina(page, professores),
        style=ft.ButtonStyle(
            shape=ft.RoundedRectangleBorder(radius=8)
        ))
    
    botao_almoxarifado = ft.ElevatedButton(
        text="Almoxarifado",
        icon=ft.icons.CHECKLIST,
        width=250,
        on_click=lambda e: ir_para_pagina(page, almoxarifado),
        style=ft.ButtonStyle(
            shape=ft.RoundedRectangleBorder(radius=8)
        ))
    
    botao_material = ft.ElevatedButton(
        text="Material",
        icon=ft.icons.SCIENCE,
        on_click=lambda e: ir_para_pagina(page, material),
        width=250,
        style=ft.ButtonStyle(
            shape=ft.RoundedRectangleBorder(radius=8)
        ))
    
    botaosair = ft.ElevatedButton(
        text="Logout", 
        icon=ft.icons.EXIT_TO_APP, 
        on_click=lambda e: ir_para_pagina(page, login))
    
    # Mostrar informações do usuário logado
    user_info = ft.Column([
        ft.Text(f"Usuário: {user_permissions['email']}"),
        ft.Text(
            "Permissão de edição: SIM" if user_permissions['can_edit_controlled'] else "Permissão de edição: NÃO",
            color=ft.colors.GREEN if user_permissions['can_edit_controlled'] else ft.colors.RED
        )
    ], spacing=5)

    page.add(ft.Column([
        logo,
        user_info,
        botao_alunos,
        botao_professores,
        botao_almoxarifado,
        
        botaosair
    ]))

'''Página de alunos'''
def alunos(page):
    page.title = "Gerenciamento de Alunos"
    page.vertical_alignment = ft.MainAxisAlignment.START
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.window_width = 900
    page.window_height = 700
    page.scroll = ft.ScrollMode.AUTO
    page.window_resizable = True
    page.window_maximized = True
    page.bgcolor = ft.colors.WHITE

    # Variáveis de controle
    alunos_dropdown = ft.Dropdown(
        label="Selecione um aluno",
        width=400,
        options=[]
    )
    
    visualizar_btn = ft.ElevatedButton(
        text="Visualizar",
        icon=ft.icons.VISIBILITY
    )
    
    excluir_btn = ft.ElevatedButton(
        text="Excluir Aluno",
        icon=ft.icons.DELETE,
        icon_color=ft.colors.RED,
        color=ft.colors.RED
    )
    
    status_message = ft.Text("", color=ft.colors.GREEN)
    botao_pag_inical= ft.ElevatedButton(text="Pàgina inicial", icon=ft.icons.HOME, on_click=lambda e: ir_para_pagina(page, pagina_inicial))
    
    # Campos para adicionar novo aluno
    nome_completo = ft.TextField(label="Nome Completo", width=400)
    telefone = ft.TextField(label="Telefone para Contato", width=400)
    email = ft.TextField(label="E-mail", width=400)
    orientador = ft.TextField(label="Nome do Orientador", width=400)
    contato_emergencia = ft.TextField(label="Contato para Emergência", width=400)
    nome_contato_emergencia = ft.TextField(label="Nome do Contato para Emergência", width=400)
    parentesco = ft.TextField(label="Parentesco", width=400)
    
    adicionar_btn = ft.ElevatedButton(
        text="Salvar Aluno",
        icon=ft.icons.SAVE
    )
    
    # Container para exibir informações do aluno
    info_container = ft.Column(
        spacing=10,
        width=600,
        visible=False
    )

    def carregar_alunos():
        try:
            sheet = client.open_by_key(SHEET_ID_ALUNOS).worksheet(WORKSHEET_NAME)
            records = sheet.get_all_records()
            
            alunos_dropdown.options = [
                ft.dropdown.Option(record["Nome completo"]) 
                for record in records
                if "Nome completo" in record
            ]
            page.update()
        except Exception as e:
            status_message.value = f"Erro ao carregar alunos: {e}"
            status_message.color = ft.colors.RED
            page.update()

    def visualizar_aluno(e):
        if not alunos_dropdown.value:
            status_message.value = "Selecione um aluno primeiro!"
            status_message.color = ft.colors.RED
            page.update()
            return
        
        try:
            sheet = client.open_by_key(SHEET_ID_ALUNOS).worksheet(WORKSHEET_NAME)
            records = sheet.get_all_records()
            headers = sheet.row_values(1)
            
            aluno_info = next(
                (record for record in records 
                 if record.get("Nome completo") == alunos_dropdown.value), 
                None
            )
            
            if not aluno_info:
                status_message.value = "Aluno não encontrado!"
                status_message.color = ft.colors.RED
                page.update()
                return
            
            # Limpa o container antes de adicionar novas informações
            info_container.controls.clear()
            
            # Adiciona cada informação do aluno com seu cabeçalho
            for header in headers:
                if header in aluno_info:
                    info_container.controls.append(
                        ft.Text(
                            f"{header}: {aluno_info[header]}",
                            size=16,
                            weight=ft.FontWeight.BOLD
                        )
                    )
            
            info_container.visible = True
            excluir_btn.disabled = False
            status_message.value = ""
            page.update()
            
        except Exception as e:
            status_message.value = f"Erro ao visualizar aluno: {e}"
            status_message.color = ft.colors.RED
            page.update()

    def excluir_aluno(e):
        if not alunos_dropdown.value:
            status_message.value = "Selecione um aluno primeiro!"
            status_message.color = ft.colors.RED
            page.update()
            return
        
        try:
            sheet = client.open_by_key(SHEET_ID_ALUNOS).worksheet(WORKSHEET_NAME)
            cell = sheet.find(alunos_dropdown.value)
            sheet.delete_rows(cell.row)
            
            status_message.value = f"Aluno {alunos_dropdown.value} excluído com sucesso!"
            status_message.color = ft.colors.GREEN
            
            # Atualiza a lista de alunos
            carregar_alunos()
            info_container.visible = False
            excluir_btn.disabled = True
            page.update()
            
        except Exception as e:
            status_message.value = f"Erro ao excluir aluno: {e}"
            status_message.color = ft.colors.RED
            page.update()

    def adicionar_aluno(e):
        if not all([nome_completo.value, telefone.value, email.value, 
                   orientador.value, contato_emergencia.value,
                   nome_contato_emergencia.value, parentesco.value]):
            status_message.value = "Preencha todos os campos!"
            status_message.color = ft.colors.RED
            page.update()
            return
        
        try:
            sheet = client.open_by_key(SHEET_ID_ALUNOS).worksheet(WORKSHEET_NAME)
            headers = sheet.row_values(1)
            
            new_row = [
                nome_completo.value,
                telefone.value,
                email.value,
                orientador.value,
                contato_emergencia.value,
                nome_contato_emergencia.value,
                parentesco.value
            ]
            
            sheet.append_row(new_row)
            
            status_message.value = f"Aluno {nome_completo.value} adicionado com sucesso!"
            status_message.color = ft.colors.GREEN
            
            # Limpa os campos e atualiza a lista
            nome_completo.value = ""
            telefone.value = ""
            email.value = ""
            orientador.value = ""
            contato_emergencia.value = ""
            nome_contato_emergencia.value = ""
            parentesco.value = ""
            
            carregar_alunos()
            page.update()
            
        except Exception as e:
            status_message.value = f"Erro ao adicionar aluno: {e}"
            status_message.color = ft.colors.RED
            page.update()

    # Configurar eventos dos botões
    visualizar_btn.on_click = visualizar_aluno
    excluir_btn.on_click = excluir_aluno
    adicionar_btn.on_click = adicionar_aluno
    
    # Carregar alunos ao iniciar
    carregar_alunos()

    # Layout da página
    page.add(
        ft.Column(
            [
                ft.Text("Gerenciamento de Alunos", size=28, weight=ft.FontWeight.BOLD),
                
                ft.Divider(),
                
                ft.Text("Consultar Aluno", size=20, weight=ft.FontWeight.BOLD),
                alunos_dropdown,
                ft.Row(
                    [visualizar_btn, excluir_btn],
                    spacing=20,
                    alignment=ft.MainAxisAlignment.CENTER
                ),
                info_container,
                
                ft.Divider(),
                
                ft.Text("Adicionar Novo Aluno", size=20, weight=ft.FontWeight.BOLD),
                nome_completo,
                telefone,
                email,
                orientador,
                contato_emergencia,
                nome_contato_emergencia,
                parentesco,
                adicionar_btn,
                
                ft.Divider(),
                
                status_message,
                botao_pag_inical
            ],
            spacing=20,
            width=600
        )
    )


'''Página de professores'''
print("\n=== INICIANDO APLICAÇÃO ===")

print("\n=== VERIFICAÇÃO DAS VARIÁVEIS ===")
print(f"Sheet ID: {SHEET_ID_PROFESSORES}")
print(f"Worksheet Name: {WORKSHEET_NAME}")

if not SHEET_ID_PROFESSORES:
    print("❌ ERRO: SHEET_ID_PROFESSORES não configurado!")
    exit()

print("\n=== AUTENTICAÇÃO ===")
try:
    creds = ServiceAccountCredentials.from_json_keyfile_name(CAMINHO_CREDENCIAIS, scope)
    client = gspread.authorize(creds)
    print("✅ Autenticação com Google Sheets bem-sucedida")
except Exception as auth_error:
    print(f"❌ ERRO NA AUTENTICAÇÃO: {auth_error}")
    exit()

def professores(page: ft.Page):
    page.title = "Sistema de Registros"
    page.vertical_alignment = ft.MainAxisAlignment.START
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.window_width = 1000
    page.window_height = 700
    page.scroll = ft.ScrollMode.AUTO
    page.window_resizable = True  # permite redimensionar
    page.window_maximized = True  # opcional: inicia em tela cheia
    page.bgcolor = ft.colors.WHITE
    
    def carregar_dados():
        print("\n=== CARREGANDO DADOS ===")
        try:
            print(f"Tentando abrir planilha com ID: {SHEET_ID_PROFESSORES}")
            sheet = client.open_by_key(SHEET_ID_PROFESSORES)
            print(f"Planilha encontrada: {sheet.title}")
            
            print(f"Listando abas disponíveis:")
            for idx, worksheet in enumerate(sheet.worksheets()):
                print(f"{idx + 1}. {worksheet.title} (id: {worksheet.id})")
            
            print(f"\nTentando acessar aba: {WORKSHEET_NAME}")
            worksheet = sheet.worksheet(WORKSHEET_NAME)
            print(f"Aba encontrada com sucesso!")
            
            print("Obtendo registros...")
            records = worksheet.get_all_records()
            print(f"Total de registros encontrados: {len(records)}")
            
            if records:
                print("\nExemplo do primeiro registro:")
                print(records[0])
            
            return records
            
        except gspread.SpreadsheetNotFound:
            print("❌ Planilha não encontrada - verifique o SHEET_ID_PROFESSORES")
            return []
        except gspread.WorksheetNotFound:
            print(f"❌ Aba '{WORKSHEET_NAME}' não encontrada na planilha")
            return []
        except Exception as e:
            print(f"❌ ERRO AO CARREGAR DADOS: {type(e).__name__}: {e}")
            return []

    def criar_tabela(dados):
        if not dados:
            print("⚠️ Nenhum dado recebido para criar tabela")
            return ft.Text("Nenhum dado encontrado", size=20)
            
        print("\n=== CRIANDO TABELA ===")
        print(f"Total de registros para exibir: {len(dados)}")
        print(f"Colunas disponíveis: {list(dados[0].keys())}")
        
        linhas = []
        cabecalhos = list(dados[0].keys())
        
        linhas.append(
            ft.DataRow(
                cells=[ft.DataCell(ft.Text(h, weight="bold")) for h in cabecalhos]
            )
        )
        
        for registro in dados[:3]:  # Mostra apenas os 3 primeiros para debug
            print(f"Registro exemplo: {registro}")
        
        for registro in dados:
            linhas.append(
                ft.DataRow(
                    cells=[ft.DataCell(ft.Text(str(valor))) for valor in registro.values()]
                )
            )
        
        return ft.DataTable(
            columns=[ft.DataColumn(ft.Text(h)) for h in cabecalhos],
            rows=linhas,
            border=ft.border.all(1, "blue"),
            border_radius=10,
            vertical_lines=ft.border.BorderSide(1, "blue"),
            horizontal_lines=ft.border.BorderSide(1, "blue"),
            heading_row_color=ft.colors.BLUE_100,
        )

    def pagina_inicial2(e):
        print("\n=== Registro de professores ===")
        page.clean()
        page.add(
            ft.Column(
                [
                    ft.Text("Registro de professores", size=30, weight="bold"),
                    ft.ElevatedButton(
                        "Ver Registros",
                        on_click=mostrar_registros,
                        icon=ft.icons.TABLE_CHART,
                        width=200
                    ),
                    ft.ElevatedButton(
                        "Página inicial",
                        on_click=lambda e: ir_para_pagina(page, pagina_inicial),
                        icon=ft.icons.HOME,
                        width=200,
                        
                    )
                ],
                alignment=ft.MainAxisAlignment.CENTER,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=40
            )
        )

    def mostrar_registros(e):
        print("\n=== MOSTRANDO REGISTROS ===")
        dados = carregar_dados()
        page.clean()
        page.add(
            ft.Column(
                [
                    ft.Row(
                        [
                            ft.IconButton(
                                icon=ft.icons.ARROW_BACK,
                                on_click=pagina_inicial2,
                                tooltip="Voltar"
                            ),
                            ft.Text("Registros da Planilha", size=25, weight="bold")
                        ],
                        alignment=ft.MainAxisAlignment.START
                    ),
                    ft.Divider(),
                    criar_tabela(dados),
                    ft.ElevatedButton(
                        "Voltar",
                        on_click=pagina_inicial2,
                        icon=ft.icons.HOME,
                        width=200
                    )
                ],
                scroll=ft.ScrollMode.AUTO,
                expand=True
            )
        )

    # Inicia na página inicial
    print("\n=== INICIANDO INTERFACE ===")
    pagina_inicial2(None)

    

'''Página de almoxarifado'''
def reagentes(page):
    # Criando os botões com ícones
    botao_controlados = ft.ElevatedButton(
        text="Reagentes Controlados",
        icon=ft.icons.WARNING,
        width=300,
        height=60,
        on_click=lambda e: ir_para_pagina(page, controlado)
    )
    
    botao_naocontrolados = ft.ElevatedButton(
        text="Reagentes não controlados", 
        icon=ft.icons.SCIENCE, 
        width=300,
        height=60,
        on_click=lambda e: ir_para_pagina(page, naocontrolado)
    )
    
    botao_volar_pag_ini = ft.ElevatedButton(
        text="Página Inicial",
        icon=ft.icons.HOME,
        width=300,
        height=60,
        on_click=lambda e: ir_para_pagina(page, pagina_inicial)
    )
    
    texto = ft.Text(value="Tipos de reagentes")
    
    # Organizando os botões em uma coluna centralizada
    botoes_coluna = ft.Column(
        controls=[
            texto,
            botao_controlados,
            botao_naocontrolados,
            botao_volar_pag_ini
        ],
        alignment=ft.MainAxisAlignment.CENTER,
        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        spacing=20
    )
    
    # Centralizando a coluna na página
    container = ft.Container(
        content=botoes_coluna,
        alignment=ft.alignment.center,
        expand=True
    )
    
    page.add(container)

def almoxarifado(page: ft.Page):
    page.title = "Gestão de Reagentes"
    page.scroll = ft.ScrollMode.AUTO
    page.bgcolor = ft.colors.WHITE
    page.window_resizable = True
    page.window_maximized = True
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    reagentes(page)

'''Reagentes não controlados'''
# Autenticação
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(CAMINHO_CREDENCIAIS, scope)
client = gspread.authorize(creds)

def naocontrolado(page: ft.Page):
    page.title = "Gerenciamento de Reagentes"
    page.window_width = 1000
    page.window_height = 700
    page.scroll = ft.ScrollMode.AUTO
    page.bgcolor = ft.colors.WHITE

    # Variáveis de estado
    reagente_editando = None
    aba_editando = None

    # Funções principais
    def carregar_reagentes():
        """Busca dados diretamente da planilha"""
        try:
            # Força nova autenticação para evitar cache
            client = gspread.authorize(creds)
            sheet = client.open_by_key(SHEET_ID_N_CONT)
            
            # Tenta ler com tratamento de timeout
            try:
                liquidos = sheet.worksheet("liquidos").get_all_records()
                solidos = sheet.worksheet("solidos").get_all_records()
            except gspread.exceptions.APIError:
                print("Timeout - Tentando novamente...")
                time.sleep(2)
                liquidos = sheet.worksheet("liquidos").get_all_records()
                solidos = sheet.worksheet("solidos").get_all_records()

            return {"liquidos": liquidos, "solidos": solidos}
        except Exception as e:
            print(f"ERRO na leitura: {str(e)}")
            return {"liquidos": [], "solidos": []}

    def pesquisar_reagente(e=None):
        """Pesquisa reagentes com dados atualizados"""
        termo = pesquisa_input.value.strip().lower()
        resultado_pesquisa.controls.clear()
        edit_container.visible = False
        
        dados = carregar_reagentes()
        aba = aba_selecionada.value
        reagentes = dados.get(aba, [])

        for reagente in reagentes:
            if termo in str(reagente.get("Produto", "")).lower():
                card = ft.Card(
                    content=ft.Container(
                        content=ft.Column(
                            [
                                ft.Text(reagente.get("Produto", ""), weight="bold"),
                                ft.Text(f"Quantidade: {reagente.get('Quantidade (L)', reagente.get('Quantidade (Kg)', 'N/A'))} {'L' if aba == 'liquidos' else 'Kg'}"),
                                ft.Text(f"Armário: {reagente.get('Armário', 'N/A')}"),
                                ft.TextButton(
                                    "FISPQ",
                                    on_click=lambda e, url=reagente.get("FISPQ", ""): webbrowser.open(url) if url else None
                                ),
                                ft.ElevatedButton(
                                    "Alterar Quantidade",
                                    on_click=lambda e, r=reagente, a=aba: editar_reagente(r, a)
                                )
                            ],
                            spacing=5
                        ),
                        padding=10,
                        width=400
                    )
                )
                resultado_pesquisa.controls.append(card)
        
        if not resultado_pesquisa.controls:
            resultado_pesquisa.controls.append(
                ft.Text("Nenhum reagente encontrado", color=ft.colors.RED)
            )
        
        page.update()

    def editar_reagente(reagente, aba):
        """Prepara a interface para edição"""
        nonlocal reagente_editando, aba_editando
        reagente_editando = reagente
        aba_editando = aba
        
        if aba == "liquidos":
            quantidade_atual = float(reagente.get("Quantidade (L)", 0))
            unidade = "L"
        else:
            quantidade_atual = float(reagente.get("Quantidade (Kg)", 0))
            unidade = "Kg"
        
        txt_produto_edit.value = f"Produto: {reagente.get('Produto', '')}"
        txt_quantidade_atual.value = f"Quantidade atual: {quantidade_atual} {unidade}"
        quantidade_gasta_input.label = f"Quantidade gasta ({unidade})"
        quantidade_gasta_input.value = ""
        edit_container.visible = True
        page.update()

    def salvar_alteracao(e):
        """Salva alterações com verificação em tempo real"""
        nonlocal reagente_editando, aba_editando
        
        try:
            gasto = float(quantidade_gasta_input.value)
            
            if aba_editando == "liquidos":
                quantidade_atual = float(reagente_editando.get("Quantidade (L)", 0))
                unidade = "L"
                col = 2
            else:
                quantidade_atual = float(reagente_editando.get("Quantidade (Kg)", 0))
                unidade = "Kg"
                col = 3
            
            if gasto <= 0:
                raise ValueError("Quantidade deve ser positiva")
            if gasto > quantidade_atual:
                raise ValueError("Quantidade insuficiente")
            
            nova_quantidade = quantidade_atual - gasto
            
            # 1. Atualização na planilha
            sheet = client.open_by_key(SHEET_ID_N_CONT)
            worksheet = sheet.worksheet(aba_editando)
            cell = worksheet.find(reagente_editando["Produto"])
            worksheet.update_cell(cell.row, col, str(nova_quantidade))
            
            # 2. Leitura DIRETA com 3 tentativas
            valor_confirmado = None
            for tentativa in range(3):
                try:
                    time.sleep(2)  # Espera maior para garantir
                    valor_confirmado = worksheet.cell(cell.row, col).value
                    print(f"Tentativa {tentativa+1}: Valor lido = {valor_confirmado}")
                    if valor_confirmado == str(nova_quantidade):
                        break
                except Exception as e:
                    print(f"Erro na leitura (tentativa {tentativa+1}): {str(e)}")
            
            # 3. Fallback se não conseguir ler
            valor_final = valor_confirmado if valor_confirmado else str(nova_quantidade)
            
            # 4. Atualização visual IMEDIATA e permanente
            edit_container.visible = False
            for control in resultado_pesquisa.controls:
                if isinstance(control, ft.Card):
                    card_content = control.content.content
                    for item in card_content.controls:
                        if isinstance(item, ft.Text) and item.value.startswith("Quantidade:"):
                            item.value = f"Quantidade: {valor_final} {unidade}"
                            # Modificação PERMANENTE no objeto
                            if aba_editando == "liquidos":
                                reagente_editando["Quantidade (L)"] = float(valor_final)
                            else:
                                reagente_editando["Quantidade (Kg)"] = float(valor_final)
                            break
            
            page.snack_bar = ft.SnackBar(
                ft.Text(f"Quantidade atualizada para {valor_final} {unidade} (confirmado)"),
                duration=4000
            )
            page.snack_bar.open = True
            page.update()
            
        except Exception as e:
            page.snack_bar = ft.SnackBar(
                ft.Text(f"Erro crítico: {str(e)}"),
                bgcolor=ft.colors.RED,
                duration=5000
            )
            page.snack_bar.open = True
            page.update()

    def cancelar_edicao(e):
        edit_container.visible = False
        page.update()

    # Componentes da interface
    aba_selecionada = ft.Dropdown(
        options=[ft.dropdown.Option("liquidos"), ft.dropdown.Option("solidos")],
        value="liquidos",
        width=200
    )
    
    pesquisa_input = ft.TextField(label="Pesquisar reagente", width=400, autofocus=True)
    resultado_pesquisa = ft.Column(scroll=ft.ScrollMode.AUTO, expand=True)
    
    txt_produto_edit = ft.Text("", size=18, weight="bold")
    txt_quantidade_atual = ft.Text("")
    quantidade_gasta_input = ft.TextField(
        label="Quantidade gasta", 
        width=200, 
        keyboard_type=ft.KeyboardType.NUMBER
    )
    
    edit_container = ft.Container(
        visible=False,
        content=ft.Column(
            [
                ft.Text("Editar Quantidade", size=20, weight="bold"),
                txt_produto_edit,
                txt_quantidade_atual,
                quantidade_gasta_input,
                ft.Row(
                    [
                        ft.ElevatedButton("Salvar", on_click=salvar_alteracao, icon=ft.icons.SAVE),
                        ft.ElevatedButton("Cancelar", on_click=cancelar_edicao, icon=ft.icons.CANCEL, color=ft.colors.RED),
                        ft.ElevatedButton(text="Página inicial",on_click=lambda e: ir_para_pagina(page, pagina_inicial))
                    ],
                    spacing=20
                )
            ],
            spacing=10
        ),
        padding=20,
        margin=10,
        border=ft.border.all(1, ft.colors.GREY_300),
        border_radius=10
    )

    # Layout principal
    page.add(
        ft.Column(
            [
                ft.Row(
                    [  
                        aba_selecionada,
                        pesquisa_input,
                        ft.ElevatedButton("Pesquisar", on_click=pesquisar_reagente),
                        ft.ElevatedButton(text="Página inicial", on_click=lambda e: ir_para_pagina(page, pagina_inicial))
                    ],
                    spacing=10,
                    alignment=ft.MainAxisAlignment.CENTER
                ),
                ft.Divider(),
                ft.Text(value="O campo de alteração abrirá ao final desta página"), 
                resultado_pesquisa,
                edit_container
            ],
            spacing=20,
            expand=True
        )
    )

if __name__ == "__naocontrolado__":
    print("=== INICIANDO APLICAÇÃO ===")
    print(f"Planilha ID: {SHEET_ID_N_CONT}")
    

'''Página de visualização de reagentes controlados'''
# Conectar à planilha do Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(CAMINHO_CREDENCIAIS, scope)
cliente = gspread.authorize(creds)
planilha = cliente.open_by_key(SHEET_ID_CONT).sheet1

def controlado(page: ft.Page):
    page.title = "Visualizador de Reagentes"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    
    # Widgets da interface
    reagentes_dropdown = ft.Dropdown(
        label="Selecione um reagente",
        width=400,
        autofocus=True
    )

    # Dicionário para armazenar as informações
    info_widgets = {
        "Produto": ft.Text(),
        "Quantidade": ft.Text(),
        "Unidade": ft.Text(),
        "Controlado por": ft.Text(),
        "FISPQ": ft.Row()  # Agora é uma Row para o ícone
    }

    def carregar_reagentes():
        dados = planilha.get_all_records()
        reagentes = sorted(set([linha["Produto"] for linha in dados]))
        reagentes_dropdown.options = [ft.dropdown.Option(reagente) for reagente in reagentes]
        page.update()

    def reagente_selecionado(e):
        selecionado = reagentes_dropdown.value
        dados = planilha.get_all_records()
        
        for linha in dados:
            if linha["Produto"] == selecionado:
                # Atualiza os textos
                info_widgets["Produto"].value = f"Produto: {linha['Produto']}"
                info_widgets["Quantidade"].value = f"Quantidade: {linha['Quantidade']}"
                info_widgets["Unidade"].value = f"Unidade: {linha['Unidade']}"
                info_widgets["Controlado por"].value = f"Controlado por: {linha['Controlado por']}"
                
                # Atualiza o widget FISPQ com ícone clicável
                fispq_url = linha.get("FISPQ", "")
                if fispq_url.startswith(('http://', 'https://')):
                    info_widgets["FISPQ"].controls = [
                        ft.Text("FISPQ: "),
                        ft.IconButton(
                            icon=ft.icons.PICTURE_AS_PDF,
                            icon_color=ft.colors.BLUE,
                            tooltip="Abrir FISPQ no navegador",
                            on_click=lambda e, url=fispq_url: webbrowser.open(url)
                        )
                    ]
                else:
                    info_widgets["FISPQ"].controls = [ft.Text("FISPQ: Não disponível")]
                
                break
                
        page.update()

    reagentes_dropdown.on_change = reagente_selecionado
    
    # Layout principal
    conteudo = [
        ft.Text("Reagentes controlados", size=30, weight="bold", color=ft.colors.BLUE_900),
        ft.Container(
            content=reagentes_dropdown,
            padding=10,
            alignment=ft.alignment.center
        ),
        ft.Container(
            content=ft.Column(
                [
                    info_widgets["Produto"],
                    info_widgets["Quantidade"],
                    info_widgets["Unidade"],
                    info_widgets["Controlado por"],
                    info_widgets["FISPQ"]
                ],
                spacing=10
            ),
            padding=20,
            width=500,
            border=ft.border.all(1, ft.colors.GREY_300),
            border_radius=10
        )
    ]
    
    # Adicionar botão de edição apenas se o usuário tiver permissão
    if user_permissions['can_edit_controlled']:
        conteudo.append(
            ft.ElevatedButton(
                text="Edição de reagentes controlados", 
                icon=ft.icons.DANGEROUS, 
                tooltip="Somente usuários autorizados",
                on_click=lambda e: ir_para_pagina(page, edi_controlados)
            )
        )
    
    # Botão de página inicial (sempre visível)
    conteudo.append(
        ft.ElevatedButton(
            text="Página inicial", 
            icon=ft.icons.HOME, 
            on_click=lambda e: ir_para_pagina(page, pagina_inicial)
        )
    )
    
    page.add(
        ft.Column(
            conteudo,
            spacing=20,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        )
    )

    carregar_reagentes()

'''Página de edição de controlados'''
# Escopo necessário
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

# Autenticando com a conta de serviço
creds = ServiceAccountCredentials.from_json_keyfile_name(CAMINHO_CREDENCIAIS, scope)
client = gspread.authorize(creds)

def edi_controlados(page: ft.Page):
    # Verificação de segurança - só permite acesso se o usuário tiver permissão
    if not user_permissions['can_edit_controlled']:
        page.add(ft.Text("❌ Acesso não autorizado", size=24, color=ft.colors.RED))
        page.add(ft.ElevatedButton(
            text="Voltar",
            icon=ft.icons.ARROW_BACK,
            on_click=lambda e: ir_para_pagina(page, controlado)
        ))
        return
    
    page.title = "Controle de estoque dos reagentes controlados"
    page.vertical_alignment = ft.MainAxisAlignment.START
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.window_width = 900
    page.window_height = 700
    page.scroll = ft.ScrollMode.AUTO
    page.bgcolor = ft.colors.WHITE
    page.window_resizable = True
    page.window_maximized = True

    # Variáveis de controle
    user_name = ft.TextField(label="Nome do Usuário", width=400)
    reagent_name = ft.Dropdown(label="Selecione o Reagente", width=400, options=[])
    quantity = ft.TextField(label="Quantidade", hint_text="Quantidade em g ou L", width=400)
    save_button = ft.ElevatedButton(text="Salvar Registro", icon=ft.icons.SAVE)
    status_message = ft.Text("", color=ft.colors.GREEN)
    new_reagent = ft.TextField(label="Novo Reagente", hint_text="Digite o nome do novo reagente", width=400)
    add_reagent_button = ft.ElevatedButton(text="Adicionar Reagente",icon=ft.icons.ADD)
    delete_reagent_button = ft.ElevatedButton(text="Excluir Reagente Selecionado", color=ft.colors.RED, icon=ft.icons.DELETE)

    def load_reagents():
        try:
            sheet = client.open_by_key(SHEET_ID_MODIFICAR).worksheet(WORKSHEET_NAME)
            headers = sheet.row_values(1)
            if len(headers) >= 3:
                reagent_columns = headers[3:]
                reagent_name.options = [ft.dropdown.Option(reagent) for reagent in reagent_columns]
                page.update()
        except Exception as e:
            status_message.value = f"Erro ao carregar reagentes: {e}"
            status_message.color = ft.colors.RED
            page.update()

    def clear_existing_sums(sheet):
        try:
            total_cells = sheet.findall("Total", in_column=1)
            for cell in sorted(total_cells, key=lambda x: x.row, reverse=True):
                sheet.delete_rows(cell.row)
        except Exception as e:
            print(f"Erro ao limpar somas anteriores: {e}")

    def add_monthly_sums(sheet):
        try:
            all_values = sheet.get_all_values()
            if len(all_values) <= 1:
                return
                
            headers = all_values[0]
            reagent_columns = headers[3:]
            
            months_data = {}
            for i, row in enumerate(all_values[1:], start=2):
                if not row or not row[0]:
                    continue
                    
                try:
                    date = datetime.strptime(row[0], "%d/%m/%Y")
                    month_year = date.strftime("%m/%Y")
                    
                    if month_year not in months_data:
                        months_data[month_year] = []
                    months_data[month_year].append((i, row))
                except ValueError as e:
                    print(f"Ignorando linha com data inválida: {e}")
                    continue
            
            for month_year, month_rows in months_data.items():
                if not month_rows:
                    continue
                    
                last_row_num = month_rows[-1][0]
                
                totals = {reagent: 0 for reagent in reagent_columns}
                for row_num, row in month_rows:
                    for i, reagent in enumerate(reagent_columns):
                        col_num = 3 + i
                        if col_num < len(row):
                            value = row[col_num]
                            try:
                                if value and (isinstance(value, (int, float)) or (isinstance(value, str) and value.replace('.', '', 1).isdigit())):
                                    totals[reagent] += float(value)
                            except ValueError as e:
                                print(f"Erro ao processar valor {value}: {e}")
                                continue
                
                sum_row = [f"Total {month_year}", sum(totals.values()), "SISTEMA"]
                sum_row += [totals.get(reagent, "") for reagent in reagent_columns]
                
                try:
                    sheet.insert_row(sum_row, last_row_num + 1)
                    sheet.format(
                        f"A{last_row_num + 1}:{gspread.utils.rowcol_to_a1(last_row_num + 1, len(sum_row))}",
                        {"textFormat": {"bold": True}}
                    )
                except Exception as e:
                    print(f"Erro ao inserir linha de total: {e}")
                    continue
                
                for other_month, other_rows in months_data.items():
                    if other_month == month_year:
                        continue
                    for i, (row_num, row) in enumerate(other_rows):
                        if row_num > last_row_num:
                            other_rows[i] = (row_num + 1, row)
            
            # Verificação
            try:
                last_rows = sheet.get_all_values()[-3:]
                print("\nVerificação - Últimas linhas:")
                for row in last_rows:
                    print(row)
            except Exception as e:
                print(f"Erro na verificação: {e}")

        except Exception as e:
            print(f"Erro em add_monthly_sums: {e}")
            status_message.value = f"Erro ao calcular totais: {str(e)}"
            status_message.color = ft.colors.RED
            page.update()

    def save_data(e):
        if not all([user_name.value, reagent_name.value, quantity.value]):
            status_message.value = "Preencha todos os campos!"
            status_message.color = ft.colors.RED
            page.update()
            return

        try:
            # Abre ambas as planilhas
            sheet_modificar = client.open_by_key(SHEET_ID_MODIFICAR).worksheet(WORKSHEET_NAME)
            sheet_estoque = client.open_by_key(SHEET_ID_CONT).worksheet(WORKSHEET_NAME)
            
            clear_existing_sums(sheet_modificar)
            
            today = datetime.now().strftime("%d/%m/%Y")
            headers = sheet_modificar.row_values(1)
            new_row = [""] * len(headers)
            new_row[0] = today
            new_row[2] = user_name.value

            try:
                reagent_col = headers.index(reagent_name.value)
                new_row[reagent_col] = float(quantity.value)
            except ValueError:
                status_message.value = "Reagente não encontrado na planilha. Recarregue a lista."
                status_message.color = ft.colors.RED
                page.update()
                return

            # 1. Salva na planilha principal (SHEET_ID_MODIFICAR)
            sheet_modificar.append_row(new_row)
            
            # 2. Atualiza o estoque na planilha SHEET_ID_CONT
            try:
                # Encontra o reagente na planilha de estoque
                estoque_data = sheet_estoque.get_all_records()
                for i, row in enumerate(estoque_data, start=2):
                    if row.get("Produtos", "").lower() == reagent_name.value.lower():
                        # Subtrai a quantidade usada
                        quantidade_atual = float(row.get("Quantidade", 0))
                        quantidade_usada = float(quantity.value)
                        nova_quantidade = quantidade_atual - quantidade_usada
                        
                        # Atualiza a planilha de estoque
                        sheet_estoque.update_cell(i, estoque_data[0].keys().index("Quantidade") + 1, str(nova_quantidade))
                        break
            except Exception as estoque_error:
                print(f"Erro ao atualizar estoque: {estoque_error}")
                # Não interrompe o fluxo principal, apenas registra o erro

            update_monthly_spending(sheet_modificar, today)
            add_monthly_sums(sheet_modificar)

            status_message.value = "Registro salvo e estoque atualizado com sucesso!"
            status_message.color = ft.colors.GREEN
            quantity.value = ""
            page.update()

        except Exception as e:
            status_message.value = f"Erro ao salvar registro: {e}"
            status_message.color = ft.colors.RED
            page.update()


    def add_reagent(e):
        if not new_reagent.value:
            status_message.value = "Digite o nome do novo reagente!"
            status_message.color = ft.colors.RED
            page.update()
            return

        try:
            sheet = client.open_by_key(SHEET_ID_MODIFICAR).worksheet(WORKSHEET_NAME)
            headers = sheet.row_values(1)
            if new_reagent.value in headers:
                status_message.value = "Este reagente já existe!"
                status_message.color = ft.colors.RED
                page.update()
                return

            sheet.add_cols(1)
            sheet.update_cell(1, len(headers)+1, new_reagent.value)
            status_message.value = f"Reagente '{new_reagent.value}' adicionado com sucesso!"
            status_message.color = ft.colors.GREEN
            new_reagent.value = ""
            load_reagents()
            page.update()

        except Exception as e:
            status_message.value = f"Erro ao adicionar reagente: {e}"
            status_message.color = ft.colors.RED
            page.update()

    def delete_reagent(e):
        if not reagent_name.value:
            status_message.value = "Selecione um reagente para excluir!"
            status_message.color = ft.colors.RED
            page.update()
            return

        try:
            sheet = client.open_by_key(SHEET_ID_MODIFICAR).worksheet(WORKSHEET_NAME)
            headers = sheet.row_values(1)
            try:
                col_index = headers.index(reagent_name.value) + 1
            except ValueError:
                status_message.value = "Reagente não encontrado na planilha."
                status_message.color = ft.colors.RED
                page.update()
                return

            sheet.delete_columns(col_index)
            status_message.value = f"Reagente '{reagent_name.value}' removido com sucesso!"
            status_message.color = ft.colors.GREEN
            load_reagents()
            page.update()

        except Exception as e:
            status_message.value = f"Erro ao remover reagente: {e}"
            status_message.color = ft.colors.RED
            page.update()

    def update_monthly_spending(sheet, date_str):
        try:
            date = datetime.strptime(date_str, "%d/%m/%Y")
            month_year = date.strftime("%m/%Y")
            data = sheet.get_all_records()
            total = 0

            for row in data:
                try:
                    row_date = datetime.strptime(row["Data"], "%d/%m/%Y")
                    if row_date.strftime("%m/%Y") == month_year:
                        for key, value in row.items():
                            if key not in ["Data", "gastos do mês", "Usuário"]:
                                try:
                                    if isinstance(value, (int, float)) or str(value).replace('.','',1).isdigit():
                                        total += float(value)
                                except:
                                    continue
                except:
                    continue

            for i, row in enumerate(data, start=2):
                try:
                    row_date = datetime.strptime(sheet.cell(i, 1).value, "%d/%m/%Y")
                    if row_date.strftime("%m/%Y") == month_year:
                        sheet.update_cell(i, 2, total)
                except:
                    continue

        except Exception as e:
            print(f"Erro ao atualizar gastos mensais: {e}")

    save_button.on_click = save_data
    add_reagent_button.on_click = add_reagent
    delete_reagent_button.on_click = delete_reagent

    load_reagents()

    page.add(
        ft.Column(
            [
                ft.Text("Registro de uso de reagentes controlados", size=28, weight=ft.FontWeight.BOLD),
                ft.Divider(),
                ft.Text("Registro de Uso", size=20, weight=ft.FontWeight.BOLD),
                user_name,
                reagent_name,
                quantity,
                save_button,
                ft.Divider(),
                ft.Text("Gerenciamento de Reagentes", size=20, weight=ft.FontWeight.BOLD),
                new_reagent,
                ft.Row(
                    [add_reagent_button, delete_reagent_button],
                    spacing=20,
                    alignment=ft.MainAxisAlignment.CENTER
                ),
                ft.Divider(),
                status_message,
                ft.ElevatedButton(text="Página inicial", icon=ft.icons.HOME,on_click=lambda e: ir_para_pagina(page, pagina_inicial),)
            ],
            spacing=20,
            width=600,
            expand=True,
        )
    )


'''Página de material'''
def material(page):
    page.title = "Material"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.bgcolor = ft.colors.WHITE
    
    # Botões
    btn_nao_controlados = ft.ElevatedButton(
        text="Reagentes não controlados",
        icon=ft.icons.SCIENCE,
        on_click=lambda e: ir_para_pagina(page, naocontrolado),
        width=250
    )
    
    btn_controlados = ft.ElevatedButton(
        text="Reagentes controlados",
        icon=ft.icons.WARNING,
        on_click=lambda e: ir_para_pagina(page, controlado),
        width=250
    )
    
    btn_voltar = ft.ElevatedButton(
        text="Página inicial",
        icon=ft.icons.HOME,
        on_click=lambda e: ir_para_pagina(page, pagina_inicial),
        width=250
    )
    
    page.add(
        ft.Column(
            [
                ft.Text("Selecione o tipo de material", size=20),
                btn_nao_controlados,
                btn_controlados,
                btn_voltar
            ],
            spacing=20,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER
        )
    )

"Iniciar com a página de login"
def main(page: ft.Page):
    login(page)

# Executa o aplicativo
ft.app(target=main)

