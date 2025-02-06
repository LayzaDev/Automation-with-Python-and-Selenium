import os
import time
import zipfile
import pyautogui
import win32com.client
import pygetwindow as gw
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.options import Options

caminho_perfil = r"C:\Users\USER\AppData\Local\Google\Chrome\User Data"
caminho_subperfil = "Default"
caminho_arq_dbf = r"C:\Users\USER\Downloads\ArquivoDBF"

options = Options()
options.add_argument(f"user-data-dir={caminho_perfil}")
options.add_argument(f"profile-directory={caminho_subperfil}")

prefs = {
    "download.default_directory": caminho_arq_dbf,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "profile.default_content_setting_values.automatic_downloads": 1,
    "plugins.always_open_pdf_externally": True
}

options.add_experimental_option("prefs", prefs)

def encontrar_arquivo_zip_mais_recente(caminho_pasta):
    arquivos = [os.path.join(caminho_pasta, f) for f in os.listdir(caminho_pasta) if f.endswith(".zip")]
    if not arquivos:
        raise FileNotFoundError(f"Nenhum arquivo ZIP foi encontrado na pasta {caminho_pasta}")
    return max(arquivos, key=os.path.getmtime)

def descompactar_zip(arquivo_zip, destino_extracao):
    os.makedirs(destino_extracao, exist_ok=True)
    with zipfile.ZipFile(arquivo_zip, 'r') as zip_ref:
        zip_ref.extractall(destino_extracao)
    print(f"Arquivo ZIP '{arquivo_zip}' descompactado em '{destino_extracao}'.")

def encontrar_arq_dbf(caminho_pasta, extensao=".dbf"):
    arquivos = [os.path.join(caminho_pasta, f) for f in os.listdir(caminho_pasta) if f.endswith(extensao)]
    if not arquivos:
        raise FileNotFoundError(f"Nenhum arquivo com a extensão '{extensao}' foi encontrado.")
    return arquivos[0]

def abrir_excel(caminho_arquivo):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    try:
        wb = excel.Workbooks.Open(caminho_arquivo)
        return excel, wb
    except Exception as e:
        print(f"Erro ao abrir arquivo no Excel: {e}")
        raise  

def fechar_janelas_indesejadas():
    try:
        time.sleep(5)
        pyautogui.hotkey('alt', 'f4')
        print("Janela do arquivo ZIP fechada.")
    except Exception as e:
        print(f"Erro ao fechar a janela ZIP: {e}")

def minimizar_janela():
    titulo_vscode = "Visual Studio Code"
    janela_vscode = [janela for janela in gw.getWindowsWithTitle(titulo_vscode) if janela.title]

    if not janela_vscode:
        print("Nenhuma janela com o titulo '{titulo_vscode}' encontrada.")
    else:
        for janela in janela_vscode:
            janela.minimize()
            print(f"A janela '{janela.title}' foi minimizada.")

def posicionar_janelas():
    janelas = gw.getWindowsWithTitle("Excel")

    if len(janelas) < 2:
        raise Exception("Menos de duas janelas do Excel abertas para ajustar.")

    # Ordenar janelas para garantir consistência
    janelas = sorted(janelas, key=lambda x: x.title)

    largura_tela, altura_tela = pyautogui.size()
    largura_meia_tela = largura_tela // 2

    # Configurar primeira janela (lado esquerdo)
    janelas[0].moveTo(0, 0)  # Posição inicial no canto superior esquerdo
    janelas[0].resizeTo(largura_meia_tela, altura_tela)  # Metade da largura, altura total

    # Configurar segunda janela (lado direito)
    janelas[1].moveTo(largura_meia_tela, 0)  # Metade da largura no eixo X
    janelas[1].resizeTo(largura_meia_tela, altura_tela)  # Metade da largura, altura total
    
    print("Janelas posicionadas com sucesso.")

try:
    os.startfile(caminho_arq_dbf)
    time.sleep(1)
    arquivo_zip_mais_recente = encontrar_arquivo_zip_mais_recente(caminho_arq_dbf)
    descompactar_zip(arquivo_zip_mais_recente, caminho_arq_dbf)
    fechar_janelas_indesejadas()
    print("Arquivo DBF descompactado")

    time.sleep(2)

    planilha_daily_reports = r"C:\Users\USER\OneDrive\Área de Trabalho\Daily Reports.xlsx"
    
    if not os.path.exists(planilha_daily_reports):
        raise FileNotFoundError(f"Arquivo {planilha_daily_reports} não encontrada.")
    
    planilha_arquivo_dbf = encontrar_arq_dbf(caminho_arq_dbf, extensao=".dbf")
    time.sleep(1)
    excel_dbf, wb_dbf = abrir_excel(planilha_arquivo_dbf)
    os.startfile(planilha_daily_reports)
    time.sleep(1)

    os.startfile(planilha_arquivo_dbf)
    time.sleep(0.5)

    fechar_janelas_indesejadas()
    time.sleep(0.5)

    posicionar_janelas()
    time.sleep(1)

    minimizar_janela()

except Exception as e:
    print(f"Erro: {e}")
