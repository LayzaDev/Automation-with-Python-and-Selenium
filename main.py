import time
import os
from selenium.webdriver.common.by import By
from automacao import configurar_navegador, encontrar_arq_zip, descompactar_arq_zip, minimizar_janela, encontrar_arq_dbf, abrir_excel, fechar_janelas_indesejadas, posicionar_janelas, caminho_arq_dbf
navegador = configurar_navegador()

def manipula_navegador():
    navegador.get("https://sinan.saude.gov.br/sinan/login/login.jsf")
    time.sleep(1.3)
    print("Entrei no site") 
    navegador.find_element(By.XPATH, '//*[@id="form"]/fieldset/div[4]/input').click()
    time.sleep(1)
    print("btn 1") 
    navegador.find_element(By.XPATH, '//*[@id="barraMenu:j_id28"]/tbody/tr/td[12]').click()
    time.sleep(1)
    print("btn 2")
    navegador.find_element(By.XPATH, '//*[@id="barraMenu:j_id52_span"]').click()
    time.sleep(1.2)
    print("btn 3") 
    navegador.find_element(By.XPATH, '//*[@id="barraMenu:j_id53:anchor"]').click()
    time.sleep(1.1)
    print("btn 4") 
    navegador.find_element(By.XPATH, '//*[@id="form:consulta_dataInicialPopupButton"]').click()
    time.sleep(1)
    print("btn 5") 
    navegador.find_element(By.XPATH, '//*[@id="form:consulta_dataInicialHeader"]/table/tbody/tr/td[3]/div').click()
    time.sleep(0.7)
    print("btn 6")
    navegador.find_element(By.XPATH, '//*[@id="form:consulta_dataInicialDateEditorLayoutM0"]').click()
    time.sleep(0.6)
    print("btn 7")
    navegador.find_element(By.XPATH, '//*[@id="form:consulta_dataInicialDateEditorButtonOk"]').click()
    time.sleep(0.8)
    print("btn 8") 
    navegador.find_element(By.XPATH, '//*[@id="form:consulta_dataInicialDayCell3"]').click()
    time.sleep(0.8)
    print("btn 9") 
    navegador.find_element(By.XPATH, '//*[@id="form:consulta_dataFinalPopupButton"]').click()
    time.sleep(1)
    print("btn 10") 
    navegador.find_element(By.XPATH, '//*[@id="form:consulta_dataFinalFooter"]/table/tbody/tr/td[5]/div').click()
    time.sleep(1)
    print("btn 11")
    navegador.find_element(By.XPATH, '//*[@id="form:tipoUf"]').click()
    time.sleep(0.8)
    print("btn 12")
    navegador.find_element(By.XPATH, '//*[@id="form:tipoUf"]/option[4]').click()
    time.sleep(1)
    print("btn 13") 
    navegador.find_element(By.XPATH, '//*[@id="form:j_id128"]').click()
    time.sleep(5)
    print("btn 14")
    print("Entrando no modo crítico") 
    navegador.find_element(By.XPATH, '//*[@id="barraMenu:j_id52_span"]').click()
    time.sleep(2.5)
    print("Ok 1") 
    navegador.find_element(By.XPATH, '//*[@id="barraMenu:j_id56:anchor"]').click()
    time.sleep(4)
    print("Ok 2") 
    valor = navegador.find_element(By.XPATH, '//*[@id="form:j_id68:0:j_id69"]/center')
    time.sleep(1)
    print(f"VALOR: {valor}")
    navegador.find_element(By.XPATH, '//*[@id="form:j_id101"]').click()
    time.sleep(5)
    print("Ok 3")
    navegador.find_element(By.XPATH, '//*[@id="form:j_id101"]').click()
    time.sleep(6)
    print("Ok 4")
    navegador.find_element(By.XPATH, '//*[@id="form:j_id101"]').click()
    time.sleep(3)

    links  = navegador.find_elements(By.XPATH, '//a[contains(text(), "Baixar arquivo DBF")]')
    if links:
        links[-1].click()
    else:
        print("Nenhum link encontrado!")

    time.sleep(1)
    print("Ok 5")

try:
    
    manipula_navegador()

    fechar_janelas_indesejadas()
    time.sleep(1)

    os.startfile(caminho_arq_dbf)
    time.sleep(1)

    arquivo_zip_mais_recente = encontrar_arq_zip(caminho_arq_dbf)
    time.sleep(0.5)

    descompactar_arq_zip(arquivo_zip_mais_recente, caminho_arq_dbf)
    time.sleep(0.5)
    print("Arquivo DBF descompactado")
    fechar_janelas_indesejadas()

    time.sleep(0.5)
    minimizar_janela()
    print("Janela vscode minimizada")

    planilha_daily_reports = r"C:\Users\USER\OneDrive\Daily Reports.xlsx"
    
    if not os.path.exists(planilha_daily_reports):
        raise FileNotFoundError(f"Arquivo {planilha_daily_reports} não encontrada.")
    
    planilha_arquivo_dbf = encontrar_arq_dbf(caminho_arq_dbf, extensao=".dbf")
    time.sleep(0.6)
    excel_dbf, wb_dbf = abrir_excel(planilha_arquivo_dbf)
    os.startfile(planilha_daily_reports)
    time.sleep(0.6)

    os.startfile(planilha_arquivo_dbf)
    time.sleep(0.5)

    fechar_janelas_indesejadas()
    time.sleep(0.3)

    posicionar_janelas()
    time.sleep(0.5)
except Exception as e:
    print(f"Erro: {e}")

input("Pressione CTRL + Z para encerrar...")
navegador.quit()