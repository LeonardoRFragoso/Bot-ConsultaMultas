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
from organizar_multas import organizar_dados  # Importa o script de organização

# Carregar variáveis de ambiente do arquivo .env
load_dotenv()

# Configurações
FILE_PATH = r"C:\Users\leonardo.fragoso\Documents\Bot-ConsultaMultas\ConsultaMultas.xlsx"
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

def consulta_multas(driver, renavam, cnpj):
    try:
        print(f"Iniciando consulta para RENAVAM: {renavam}, CNPJ: {cnpj}")
        driver.get(PAGE_URL)

        # Garantir que o iframe está disponível e mudar para ele
        iframe = WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="frameConsulta"]'))
        )
        print("Trocado para o IFrame com sucesso!")

        # Preencher os campos de RENAVAM e CNPJ
        renavam_field = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="MultasRenavam"]'))
        )
        renavam_field.send_keys(str(renavam))

        cpf_field = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="MultasCpfcnpj"]'))
        )
        cpf_field.send_keys(str(cnpj))  # Usando 'CNPJ' como no cabeçalho da planilha

        # Resolver o CAPTCHA
        token = obter_token_captcha(API_KEY, SITE_KEY, PAGE_URL)
        captcha_response = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'g-recaptcha-response'))
        )
        driver.execute_script("arguments[0].innerHTML = arguments[1];", captcha_response, token)

        # Clicar no botão 'Consultar'
        consultar_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="btPesquisar"]'))
        )
        consultar_button.click()
        print("Botão 'Consultar' clicado com sucesso.")
        time.sleep(10)

        # Verificar a mensagem de "não há multas"
        try:
            mensagem = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '//*[contains(text(),"Não há multa")]'))
            )
            print("Mensagem de 'Não há multas' encontrada. Pulando para o próximo RENAVAM.")
            return []
        except:
            print("Nenhuma mensagem de 'Não há multas' encontrada. Buscando por tabelas...")

        # Verificar a mensagem de "Este veículo não consta no cadastro"
        try:
            mensagem_nao_consta = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="caixaInformacao"]'))
            )
            if "não consta no cadastro" in mensagem_nao_consta.text:
                print("Mensagem de 'Veículo não consta no cadastro' encontrada. Pulando para o próximo RENAVAM.")
                return []
        except:
            pass

        # Agora que o botão foi clicado e o sistema retornou as informações, vamos procurar as divs com as multas
        return extrair_multas_dos_iframes(driver)

    except Exception as e:
        print(f"Erro durante a consulta para RENAVAM {renavam}: {e}")
        return []

def extrair_multas_dos_iframes(driver):
    print("Iniciando extração de multas...")
    multas = []

    try:
        # Aguardar até que a div 'caixaTabela' esteja disponível dentro do iframe
        caixa_tabela = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="caixaTabela"]'))
        )
        print("Elemento 'caixaTabela' encontrado.")

        # Buscar todas as tabelas dentro da div caixaTabela
        tabelas = caixa_tabela.find_elements(By.XPATH, './/table[@class="tabelaDescricao"]')
        print(f"{len(tabelas)} tabela(s) de multas encontrada(s).")

        for tabela_index, tabela in enumerate(tabelas, start=1):
            try:
                linhas = tabela.find_elements(By.TAG_NAME, 'tr')
                for linha in linhas:
                    colunas = linha.find_elements(By.TAG_NAME, 'td')
                    if len(colunas) > 1:  # Apenas processar linhas com dados relevantes
                        multa = {
                            'Auto de Infração': colunas[0].text if len(colunas) > 0 else '',
                            'Auto de Renavam': colunas[1].text if len(colunas) > 1 else '',
                            'Data para Pagamento com Desconto': colunas[2].text if len(colunas) > 2 else '',
                            'Enquadramento da Infração': colunas[3].text if len(colunas) > 3 else '',
                            'Data da Infração': colunas[4].text if len(colunas) > 4 else '',
                            'Hora': colunas[5].text if len(colunas) > 5 else '',
                            'Placa Relacionada': colunas[6].text if len(colunas) > 6 else '',
                            'Valor Original R$': colunas[7].text if len(colunas) > 7 else '',
                            'Valor a Ser Pago R$': colunas[8].text if len(colunas) > 8 else '',
                            'Órgão Emissor': colunas[9].text if len(colunas) > 9 else ''
                        }
                        multas.append(multa)
                print(f"Tabela {tabela_index} processada com sucesso.")
            except Exception as e:
                print(f"Erro ao processar tabela {tabela_index}: {e}")

    except Exception as e:
        print(f"Erro ao localizar a div 'caixaTabela' ou as tabelas: {e}")

    # Retornar ao contexto principal
    driver.switch_to.default_content()
    print(f"{len(multas)} multa(s) extraída(s) no total.")
    return multas

def main():
    # Ler a planilha
    df = pd.read_excel(FILE_PATH)

    # Validar as variáveis carregadas
    if not API_KEY or not SITE_KEY or not DRIVER_PATH:
        raise Exception("API_KEY, SITE_KEY ou DRIVER_PATH não configurados corretamente no .env")

    # Inicializar o navegador
    driver = iniciar_navegador()

    # Criar um dicionário para organizar as abas do Excel
    resultados = {}
    sem_multas = []

    # Iterar sobre as linhas da planilha
    for index, row in df.iterrows():
        renavam = row['RENAVAM']
        cnpj = row['CNPJ']  # Ajustado para usar 'CNPJ' como no cabeçalho da planilha
        print(f"Consultando para RENAVAM: {renavam}, CNPJ: {cnpj}")
        multas = consulta_multas(driver, renavam, cnpj)

        if multas:
            # Criar DataFrame para o RENAVAM atual
            resultados[renavam] = pd.DataFrame(multas)
        else:
            # Adicionar à lista de RENAVAM sem multas
            sem_multas.append({'RENAVAM': renavam, 'CNPJ': cnpj, 'Status': 'Sem multas registradas'})

    # Salvar os resultados em um arquivo Excel com múltiplas abas
    with pd.ExcelWriter("resultados_organizados.xlsx") as writer:
        for renavam, multas_df in resultados.items():
            multas_df.to_excel(writer, sheet_name=str(renavam), index=False)
        
        if sem_multas:
            # Salvar os RENAVAM sem multas em uma aba separada
            pd.DataFrame(sem_multas).to_excel(writer, sheet_name="Sem Multas", index=False)

    print("Resultados salvos em: resultados_organizados.xlsx")

    # Fechar o navegador
    driver.quit()

    # Chamar o script de organização
    print("Organizando os dados...")
    organizar_dados()
    print("Organização concluída!")

if __name__ == "__main__":
    main()
