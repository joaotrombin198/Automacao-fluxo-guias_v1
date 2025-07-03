import json
import re
import requests
import time
import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

# Inicializa o navegador
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

navegador = webdriver.Chrome(options=chrome_options)
navegador.maximize_window()
navegador.get("https://rda-hml.unimedsc.com.br/autsc2/Login.do")

def AcessaAut():
    # Informa dados de login e acessa o sistema
    WebDriverWait(navegador, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="ds_login"]'))
    ).send_keys("admin198")
    navegador.find_element(By.XPATH, '//*[@id="passwordTemp"]').send_keys("#Unimed198")

    WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="Button_DoLogin"]'))
    ).click()


def obter_data_atual():
    return datetime.now().strftime("%Y-%m-%d")

def get_element_text(driver, by, value):
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((by, value))
        )
        return element.text
    except TimeoutException:
        return None


def get_element_value(driver, by, value):
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((by, value))
        )
        return element.get_attribute('value')
    except TimeoutException:
        return None


def verifica_elemento_xpath(driver, xpath):
    try:
        driver.find_element(By.XPATH, xpath)
        return True
    except NoSuchElementException:
        return False


def registrar_dados_guia(numero_guia, codigo_atual, qt_solicitada, qt_anterior, qt_previsto):
    data_hora_atual = datetime.now().strftime('%d/%m/%Y - %H:%M:%S')
    nome_arquivo_dados = f'dados_guias{datetime.now().strftime("%d-%m-%Y")}.txt'

    novo_registro = (f"{data_hora_atual}\tGuia: {numero_guia} |\tCod procedimento: {codigo_atual} |"
                     f"\tQtd Solicitada: {qt_solicitada} |\tQtd Anterior: {qt_anterior} | \tQtd Prevista (API): {qt_previsto}")

    with open(nome_arquivo_dados, 'a', encoding='utf-8') as file:
        file.write(novo_registro + "\n")


def enviar_json(json_data, url):
    headers = {'Content-Type': 'application/json',
               'ANALYZER-API-AUTHORIZATION': '7CXdWHlT0tJPymFRsun0QXaT5UZm+tORJw/osBrZeC2BdXAEFvqQ6OX863Jps23j'}
    response = requests.post(url, data=json.dumps(json_data), headers=headers, verify=False)
    return response.status_code, response.text


def registrar_log_txt(numero_guia, erro_guia):
    data_atual = obter_data_atual()
    nome_arquivo_log = f'guia_erros_{data_atual}.txt'
    registros_existentes = set()

    if os.path.exists(nome_arquivo_log):
        with open(nome_arquivo_log, 'r', encoding='utf-8') as file:
            for line in file:
                registros_existentes.add(line.strip())

    novo_registro = f"{datetime.now().strftime('%d/%m/%Y - %H:%M:%S')}\t{numero_guia} - {erro_guia}"
    if novo_registro not in registros_existentes:
        with open(nome_arquivo_log, 'a', encoding='utf-8') as file:
            file.write(novo_registro + "\n")
