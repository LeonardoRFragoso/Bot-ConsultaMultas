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
    print("Iniciando o navegador...")
    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--start-maximized")
    service = Service(DRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)
    print("Navegador iniciado com sucesso!")
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
    response = requests.post(url, data=payload).json()
    
    if response.get('status') != 1:
        print(f"Erro ao enviar reCAPTCHA: {response.get('request')}")
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
            print(f"Erro ao obter solução: {token_response.get('request')}")
            raise Exception(f"Erro ao obter solução: {token_response.get('request')}")
    
    print("Tempo limite ao resolver o CAPTCHA.")
    raise Exception("Tempo limite ao resolver o CAPTCHA.")

def extrair_multas(driver):
    print("Iniciando extração das multas...")
    multas = []
    i = 1
    while True:
        try:
            tabela_xpath = f"//*[@id='caixaTabela']/div[4]/table[{i}]"
            tabela = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, tabela_xpath))
            )
            linhas = tabela.find_elements(By.TAG_NAME, 'tr')

            for linha in linhas:
                dados = linha.find_elements(By.TAG_NAME, 'td')
                if len(dados) > 1:  # Filtra linhas que contêm dados
                    multa = {
                        'Auto de Infração': dados[0].text,
                        'Auto de Renavam': dados[1].text,
                        'Data de Pagamento com Desconto': dados[2].text,
                        'Enquadramento da Infração': dados[3].text,
                        'Valor Original': dados[6].text,
                        'Valor a Ser Pago': dados[7].text
                    }
                    multas.append(multa)
            print(f"Tabela {i} extraída com sucesso.")
            i += 1  # Avançar para a próxima tabela
        except Exception as e:
            print(f"Sem mais tabelas a serem processadas ou erro: {e}")
            break  # Interrompe quando não encontrar mais tabelas ou erro
    
    print(f"{len(multas)} multa(s) extraída(s).")
    return multas

def consulta_multas(driver, renavam, cpf_cnpj):
    try:
        print(f"Iniciando consulta para RENAVAM: {renavam}, CPF/CNPJ: {cpf_cnpj}")
        driver.get(PAGE_URL)

        # Esperar o iframe ser carregado completamente
        iframe = WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="frameConsulta"]'))
        )
        print("Trocado para o IFrame com sucesso!")

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

        # Aguardar pelo menos 10 segundos antes de verificar se as tabelas ou a mensagem foram carregadas
        time.sleep(10)

        # Verificar se a mensagem "Não há multa" aparece
        try:
            mensagem_sem_multas = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[contains(text(),"Não há multa")]'))
            )
            print("Mensagem de não há multas encontrada. Recarregando a página e buscando próximo RENAVAM.")
            driver.refresh()  # Recarregar a página
            return []  # Retornar lista vazia, indicando que não há multas
        except Exception as e:
            print("Nenhuma mensagem de não há multas encontrada, buscando por tabelas...")

        # Extrair os dados das multas
        multas = extrair_multas(driver)
        print(f"Multas encontradas para RENAVAM {renavam}: {len(multas)}")

        return multas

    except Exception as e:
        print("Erro durante a consulta:", e)
        return []

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
        multas = consulta_multas(driver, renavam, cnpj)

        # Armazenar as multas encontradas na planilha
        df.at[index, 'Resultado'] = multas if multas else "Sem multas registradas"

    # Salvar os resultados em um novo arquivo
    output_file = FILE_PATH.replace(".xlsx", "_resultado.xlsx")
    df.to_excel(output_file, index=False)
    print(f"Resultados salvos em: {output_file}")  # Print informando que os resultados foram salvos

    # Fechar o navegador
    driver.quit()

if __name__ == "__main__":
    main()
