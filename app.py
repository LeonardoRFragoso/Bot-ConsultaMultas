import time
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service

# Configurações
API_KEY = "13a1288df1b6d6145a00476d48bf1c2b"  # Substitua pela sua chave da API do 2Captcha
SITE_KEY = "6LfP47IUAAAAAIwbI5NOKHyvT9Pda17dl0nXl4xv"  # Substitua pela sitekey do reCAPTCHA no site
PAGE_URL = "https://www.detran.rj.gov.br/_monta_aplicacoes.asp?cod=11&tipo=consulta_multa"
RENAVAM = "329933248"
CNPJ = "15025071000116"

def iniciar_navegador():
    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--start-maximized")

    driver_path = r"C:\Users\leona\OneDrive\Documentos\Bot-ConsultaMultas\chromedriver.exe"
    service = Service(driver_path)
    driver = webdriver.Chrome(service=service, options=options)
    return driver

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

    response = requests.post(url, data=payload)
    response_data = response.json()

    if response_data['status'] == 1:
        captcha_id = response_data['request']
        print(f"ID do reCAPTCHA enviado: {captcha_id}")
    else:
        raise Exception(f"Erro ao enviar reCAPTCHA: {response_data['request']}")

    token_url = f"https://2captcha.com/res.php?key={api_key}&action=get&id={captcha_id}&json=1"
    for _ in range(20):  # Tentativas de resolução (5 segundos por tentativa)
        time.sleep(5)
        token_response = requests.get(token_url).json()
        if token_response['status'] == 1:
            print("Token do CAPTCHA resolvido obtido com sucesso.")
            return token_response['request']
        elif token_response['request'] != 'CAPCHA_NOT_READY':
            raise Exception(f"Erro ao obter solução: {token_response['request']}")

    raise Exception("Tempo limite ao resolver o CAPTCHA.")

def consulta_multas(driver, renavam, cpf_cnpj):
    try:
        print(f"Iniciando consulta para RENAVAM: {renavam}, CPF/CNPJ: {cpf_cnpj}")
        driver.get(PAGE_URL)

        WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="frameConsulta"]'))
        )
        print("Trocando para o IFrame principal.")

        renavam_field = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="MultasRenavam"]'))
        )
        renavam_field.send_keys(str(renavam))
        print("Campo RENAVAM preenchido.")

        cpf_field = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="MultasCpfcnpj"]'))
        )
        cpf_field.send_keys(str(cpf_cnpj))
        print("Campo CPF/CNPJ preenchido.")

        # Obter o token do reCAPTCHA dinamicamente
        token = obter_token_captcha(API_KEY, SITE_KEY, PAGE_URL)
        print(f"Token do CAPTCHA obtido: {token}")

        # Inserir o token no campo g-recaptcha-response
        captcha_response = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'g-recaptcha-response'))
        )
        driver.execute_script(
            "arguments[0].innerHTML = arguments[1];", captcha_response, token
        )
        print("Token do CAPTCHA resolvido inserido.")

        # Clicar no botão CONSULTAR
        consultar_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="btPesquisar"]'))
        )
        consultar_button.click()
        print("Botão 'Consultar' clicado com sucesso.")

        # Aguardar os resultados
        print("Aguardando resultados...")
        resultado = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="caixaInformacao"]'))
        ).text
        print("Consulta realizada com sucesso!")
        print("Resultado:", resultado)

    except Exception as e:
        print("Erro durante a consulta:", e)
    finally:
        print("Navegador permanecerá aberto para interação manual. Feche-o manualmente.")
        input("Pressione Enter no terminal para encerrar o navegador...")
        driver.quit()

def main():
    try:
        driver = iniciar_navegador()
        consulta_multas(driver, RENAVAM, CNPJ)
    except Exception as e:
        print(f"Erro no processo principal: {e}")

if __name__ == "__main__":
    main()
