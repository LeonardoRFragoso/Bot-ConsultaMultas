import os
import time
import pandas as pd
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from dotenv import load_dotenv

# Carregar variáveis de ambiente do arquivo .env
load_dotenv()

# Configurações
FILE_PATH = r"C:\Users\leona\OneDrive\Documentos\Bot-ConsultaMultas\Consulta Multas.xlsx"
API_KEY = os.getenv("API_KEY")
SITE_KEY = os.getenv("SITE_KEY")
PAGE_URL = os.getenv("PAGE_URL")
DRIVER_PATH = os.getenv("DRIVER_PATH")

def iniciar_navegador():
    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--start-maximized")
    service = Service(DRIVER_PATH)
    return webdriver.Chrome(service=service, options=options)

def obter_token_captcha(api_key, site_key, page_url):
    print("Enviando reCAPTCHA para resolver...")
    url = "https://2captcha.com/in.php"
    payload = {
        'key': api_key,
        'method': 'userrecaptcha',
        'googlekey': site_key,
        'pageurl': page_url,
        'json': 1
    }
    response = requests.post(url, data=payload).json()
    if response.get('status') != 1:
        raise Exception(f"Erro ao enviar reCAPTCHA: {response.get('request')}")

    captcha_id = response.get('request')
    print(f"ID do reCAPTCHA enviado: {captcha_id}")
    token_url = f"https://2captcha.com/res.php?key={api_key}&action=get&id={captcha_id}&json=1"

    for _ in range(30):  # Tentativas de resolução (150 segundos no total)
        time.sleep(5)
        token_response = requests.get(token_url).json()
        if token_response.get('status') == 1:
            print("Token do CAPTCHA resolvido com sucesso.")
            return token_response.get('request')
        elif token_response.get('request') != 'CAPCHA_NOT_READY':
            raise Exception(f"Erro ao obter solução: {token_response.get('request')}")

    raise Exception("Tempo limite ao resolver o CAPTCHA.")

def consulta_multas(driver, renavam, cpf_cnpj):
    try:
        print(f"Iniciando consulta para RENAVAM: {renavam}, CPF/CNPJ: {cpf_cnpj}")
        driver.get(PAGE_URL)

        WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="frameConsulta"]'))
        )

        # Preencher os campos
        renavam_field = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="MultasRenavam"]'))
        )
        renavam_field.send_keys(str(renavam))

        cpf_field = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="MultasCpfcnpj"]'))
        )
        cpf_field.send_keys(str(cpf_cnpj))

        # Resolver o reCAPTCHA
        token = obter_token_captcha(API_KEY, SITE_KEY, PAGE_URL)
        captcha_response = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'g-recaptcha-response'))
        )
        driver.execute_script("arguments[0].innerHTML = arguments[1];", captcha_response, token)

        # Clicar no botão Consultar
        consultar_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="btPesquisar"]'))
        )
        consultar_button.click()
        print("Botão 'Consultar' clicado com sucesso.")

        # Aguardar o resultado da consulta
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="caixaInformacao"]'))
        )

        # Verificar o resultado da consulta
        try:
            resultado = driver.find_element(By.XPATH, '//*[contains(text(),"Não há multa")]').text
            print(f"Resultado para RENAVAM {renavam}: {resultado}")
        except:
            try:
                resultado = driver.find_element(By.XPATH, '//*[contains(text(),"AS INFORMAÇÕES SÃO VÁLIDAS ATÉ O MOMENTO DA CONSULTA")]').text
                print(f"Resultado para RENAVAM {renavam}: {resultado}")
            except:
                resultado = driver.find_element(By.XPATH, '//*[@id="caixaInformacao"]').text
                print(f"Resultado para RENAVAM {renavam}: {resultado}")

        return resultado

    except Exception as e:
        print("Erro durante a consulta:", e)
        return "Erro na consulta"

def main():
    # Ler a planilha
    df = pd.read_excel(FILE_PATH)

    # Validar as variáveis carregadas
    if not API_KEY or not SITE_KEY or not DRIVER_PATH:
        raise Exception("API_KEY, SITE_KEY ou DRIVER_PATH não configurados corretamente no .env")

    # Inicializar o navegador
    driver = iniciar_navegador()

    # Iterar sobre as linhas da planilha
    for index, row in df.iterrows():
        renavam = row['RENAVAM']
        cnpj = row['CNPJ']
        print(f"Consultando para RENAVAM: {renavam}, CNPJ: {cnpj}")
        resultado = consulta_multas(driver, renavam, cnpj)

        # Atualizar a planilha com o resultado
        df.at[index, 'Resultado'] = resultado

    # Salvar os resultados em um novo arquivo
    output_file = FILE_PATH.replace(".xlsx", "_resultado.xlsx")
    df.to_excel(output_file, index=False)
    print(f"Resultados salvos em: {output_file}")

    # Fechar o navegador
    driver.quit()

if __name__ == "__main__":
    main()
