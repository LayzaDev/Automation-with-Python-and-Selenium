#!/usr/bin/env python
#coding: utf-8

import os
import time
import zipfile
import pyautogui
import win32com.client
import pygetwindow as gw
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

caminho_perfil = r"C:\Users\USER\AppData\Local\Google\Chrome\User Data"
caminho_subperfil = "Default"
caminho_arq_dbf = r"C:\Users\USER\Downloads\ArquivoDBF"

def configurar_navegador():
    # Configurando o navegador Chrome com o perfil padrão do Chrome
    options = Options() # usado para definir as configurações do Chrome antes de inici-alo com o Selenium
    options.add_argument(f"user-data-dir={caminho_perfil}") # Acessa o diretório base em que o Chrome armazena os usuários
    options.add_argument(f"profile-directory={caminho_subperfil}") # Define o subperfil a ser acessado (no caso, o perfil será o padrão)

    # Dicionário de preferências
    prefs = {
        "download.default_directory": caminho_arq_dbf, # Define a pasta em que o arquivo baixado será salvo
        "download.prompt_for_download": False,           # Baixar o arquivo sem o prompt de confirmação (salvar como...)
        "download.directory_upgrade": True,              # Atualiza automaticamente o diretório do download
        "safebrowsing.enabled": True,                    # Mantém o recurso de navegação segura do Chrome ativado, evitando problemas com arquivos não seguros
        "profile.default_content_setting_values.automatic_downloads": 1,  # Permite downloads automáticos
        "plugins.always_open_pdf_externally": True       # Para abrir PDFs automaticamente fora do navegador
    }

    options.add_experimental_option("prefs", prefs) # adicionando o dicionário de preferências às opções do Chrome
    return webdriver.Chrome(options=options)

# Busca pelo arquivo zip mais recente armazenado na pasta ArquivoDBF
def encontrar_arq_zip(caminho_arquivo, extensao=".zip"):
    arquivos = [os.path.join(caminho_arquivo, f) for f in os.listdir(caminho_arquivo) if f.endswith(extensao)]
    if not arquivos:
        raise FileNotFoundError(f"Nenhum arquivo '{extensao}' foi encontrado na pasta {caminho_arquivo}.")
    print(f"Arquivo '{extensao}' encontrado com sucesso.")
    return max(arquivos, key=os.path.getmtime) # pega o timestamp da ultima modificação do arq e compara com os outros arquivos

# Descompacta o arquivo zip e extrai seu conteúdo para a pasta de destino
def descompactar_arq_zip(arquivo_zip, destino_extracao):
    os.makedirs(destino_extracao, exist_ok=True)
    with zipfile.ZipFile(arquivo_zip, 'r') as zip_ref:
        zip_ref.extractall(destino_extracao)
    print(f"Arquivo ZIP '{arquivo_zip}' descompactado em '{destino_extracao}'.")

# Busca pelo arquivo .dbf que foi extraído mais recentemente
def encontrar_arq_dbf(caminho_arquivo, extensao=".dbf"):
    arquivos = [os.path.join(caminho_arquivo, f) for f in os.listdir(caminho_arquivo) if f.endswith(extensao)]
    if not arquivos:
        raise FileNotFoundError(f"Nenhum arquivo com a extensão '{extensao}' foi encontrado.")
    print(f"Arquivo '{extensao}' encontrado com sucesso.")
    return max(arquivos, key=os.path.getmtime)

# Cria uma instancia do Excel no Windows, torna ela visível e tenta abrir um arquivo '.xlsx'
def abrir_excel(caminho_excel):
    excel = win32com.client.Dispatch("Excel.Application") # Abre uma instância do Excel no **WINDOWS**
    excel.Visible = True # Faz o excel aparecer/abrir na tela
    try:
        wb = excel.Workbooks.Open(caminho_excel) # Equivale a ir no Excel e fazer "Arquivo > Abrir"
        return excel, wb
    except Exception as e:
        print(f"Erro ao abrir arquivo no Excel: {e}")
        raise    

# Aguarda alguns segundos e depois simula o atalho 'Alt + F4' para fechar as janelas indesejadas que estão ativas na tela.
def fechar_janelas_indesejadas():
    try:
        time.sleep(5) # Dá um tempo pra janela indesejada aparecer na tela
        pyautogui.hotkey('alt', 'f4') # Simula o atalho 'Alt + F4', que fecha janelas ativas
        print("Janela indesejada fechada com sucesso.")
    except Exception as e:
        print(f"Erro ao fechar a janela indesejada: {e}")

# Busca todas as janelas abertas com o título "Visual Studio Code" e as minimiza. 
def minimizar_janela():
    titulo_vscode = "Visual Studio Code"
    janela_vscode = [janela for janela in gw.getWindowsWithTitle(titulo_vscode) if janela.title]

    if not janela_vscode:
        print(f"Nenhuma janela com o titulo '{titulo_vscode}' encontrada.")
    else:
        for janela in janela_vscode:
            janela.minimize()
            print(f"A janela '{janela.title}' foi minimizada.")


def posicionar_janelas():
    janelas = gw.getWindowsWithTitle("Excel") # Retorna uma lista de janelas abertos que contêm "Excel" no titulo

    if len(janelas) < 2: # Verifica se existe pelo menos duas janelas abertas para fazer o posicionamento lado a lado
        raise Exception("É necessário ao menos duas janelas do Excel abertas para ajustar.")

    janelas = sorted(janelas, key=lambda x: x.title) # Ordena as janelas pelo titulo para garantir uma ordenação
    largura_total, altura_total = pyautogui.size() # Obtem a largura e a altura da tela do computador
    metade_largura_janela = largura_total // 2 # Calcula a metade da largura da tela

    # Configurando a janela do lado esquerdo
    janelas[0].moveTo(0, 0)  # Move a primeira janela para o canto superior esquerdo da tela (x = 0, y = 0)
    janelas[0].resizeTo(metade_largura_janela, altura_total) # Faz a janela ocupar metade esquerda da tela e a altura total

    # Configurar janela do lado direito
    janelas[1].moveTo(metade_largura_janela, 0)  # Move a janela para começar na metade da tela
    janelas[1].resizeTo(metade_largura_janela, altura_total) # redimensiona para ocupar a metade direita da tela e a altura total.
    
    print("Janelas posicionadas com sucesso.")