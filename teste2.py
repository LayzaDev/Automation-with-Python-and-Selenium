#!/usr/bin/env python
#coding: utf-8

import os
import time
import pygetwindow as gw
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.options import Options
    
try:
    arquivo_origem = r"C:\Users\USER\Downloads\ArquivoDBF\planilha1.xlsx"
    arquivo_destino = r"C:\Users\USER\Downloads\ArquivoDBF\planilha2.xlsx"

    os.startfile(arquivo_origem)
    time.sleep(2)
    os.startfile(arquivo_destino)
    time.sleep(2)

    nome_coluna = "DT_NOTIFIC"

    df_origem = pd.read_excel(arquivo_origem)

    if nome_coluna not in df_origem.columns:
        print(f"Coluna '{nome_coluna}' não encontrada na planilha de origem.")
        exit()

    if not pd.api.types.is_datetime64_any_dtype(df_origem[nome_coluna]):
        df_origem[nome_coluna] = pd.to_datetime(df_origem[nome_coluna], origin='1899-12-30', unit='D', errors='coerce')

    coluna_dados = df_origem[[nome_coluna]].dropna()

    try:
        df_destino = pd.read_excel(arquivo_destino)
        df_destino = pd.concat([df_destino, coluna_dados], axis=1)
    except FileNotFoundError:
        df_destino = coluna_dados

    with pd.ExcelWriter(arquivo_destino, engine="openpyxl", date_format="dd-mm-yyyy", datetime_format="dd-mm-yyyy") as writer:
        df_destino.to_excel(writer, index=False)

    print(f"Dados da coluna '{nome_coluna}' copiados com formato de data (dd-mm-yyyy) para '{arquivo_destino}'.")

except Exception as e:
    print(f"Erro: {e}")