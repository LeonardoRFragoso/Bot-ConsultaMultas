import os
import time
import pandas as pd
import requests
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from dotenv import load_dotenv
from datetime import datetime

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
        # Localizar a div 'caixaTabela' e capturar todas as tabelas de multas
        caixa_tabela = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="caixaTabela"]'))
        )
        tabelas = caixa_tabela.find_elements(By.XPATH, './/table[@class="tabelaDescricao"]')
        print(f"{len(tabelas)} tabela(s) encontrada(s).")

        for tabela_index, tabela in enumerate(tabelas, start=1):
            try:
                linhas = tabela.find_elements(By.TAG_NAME, 'tr')
                if not linhas:
                    print(f"Tabela {tabela_index} vazia ou não contém linhas.")
                    continue

                multa = {}
                for linha_index, linha in enumerate(linhas):
                    colunas = linha.find_elements(By.TAG_NAME, 'td')

                    if not colunas:
                        print(f"Linha {linha_index + 1} da tabela {tabela_index} não contém colunas.")
                        continue
                    
                    # Atualizar lógica de mapeamento com base na estrutura visualizada
                    if linha_index == 0 and len(colunas) >= 3:
                        multa['Auto_de_Inflacao'] = colunas[0].text.split(":", 1)[-1].strip()
                        multa['Auto_de_Renainf'] = colunas[1].text.split(":", 1)[-1].strip()
                        multa['Data_para_Pagamento_com_Desconto'] = colunas[2].text.split(":", 1)[-1].strip()

                    elif linha_index == 1 and len(colunas) >= 3:
                        multa['Enquadramento_da_Inflacao'] = colunas[0].text.split(":", 1)[-1].strip()
                        multa['Data_da_Inflacao'] = colunas[1].text.split(":", 1)[-1].strip()
                        multa['Hora'] = colunas[2].text.split(":", 1)[-1].strip()

                    elif linha_index == 2 and len(colunas) >= 2:
                        multa['Descricao'] = colunas[0].text.split(":", 1)[-1].strip()
                        multa['Placa_Relacionada'] = colunas[1].text.split(":", 1)[-1].strip()

                    elif linha_index == 3 and len(colunas) >= 3:
                        multa['Local_da_Inflacao'] = colunas[0].text.split(":", 1)[-1].strip()
                        multa['Valor_Original_R'] = colunas[1].text.split(":", 1)[-1].strip()
                        multa['Valor_a_Ser_Pago_R'] = colunas[2].text.split(":", 1)[-1].strip()

                    elif linha_index == 4 and len(colunas) >= 3:
                        multa['Status_do_Pagamento'] = colunas[0].text.split(":", 1)[-1].strip()
                        multa['Orgao_Emissor'] = colunas[1].text.split(":", 1)[-1].strip()
                        multa['Agente_Emissor'] = colunas[2].text.split(":", 1)[-1].strip()

                # Validar se os campos essenciais foram preenchidos
                if multa.get('Auto_de_Inflacao') and multa.get('Auto_de_Renainf'):
                    multas.append(multa)
                    print(f"Tabela {tabela_index} processada com sucesso.")
                else:
                    print(f"Tabela {tabela_index} ignorada devido a dados incompletos.")

            except Exception as e:
                print(f"Erro ao processar tabela {tabela_index}: {e}")

    except Exception as e:
        print(f"Erro ao localizar as tabelas de multas: {e}")

    return multas

def salvar_dados_em_json(multas, output_path="multas.json"):
    try:
        if os.path.exists(output_path):
            with open(output_path, "r", encoding="utf-8") as file:
                dados_existentes = json.load(file)
        else:
            dados_existentes = []

        dados_existentes.extend(multas)

        with open(output_path, "w", encoding="utf-8") as file:
            json.dump(dados_existentes, file, ensure_ascii=False, indent=4)
        
        print(f"Dados salvos no arquivo {output_path} com sucesso!")
    except Exception as e:
        print(f"Erro ao salvar dados em JSON: {e}")

def main():
    # Ler a planilha
    df = pd.read_excel(FILE_PATH)

    if not API_KEY or not SITE_KEY or not DRIVER_PATH:
        raise Exception("API_KEY, SITE_KEY ou DRIVER_PATH não configurados corretamente no .env")

    driver = iniciar_navegador()
    resultados = []
    sem_multas = []

    for index, row in df.iterrows():
        renavam = row['RENAVAM']
        cnpj = row['CNPJ']
        multas = consulta_multas(driver, renavam, cnpj)

        if multas:
            resultados.extend(multas)
        else:
            sem_multas.append({'RENAVAM': renavam, 'CNPJ': cnpj, 'Status': 'Sem multas registradas'})

    # Salvar os resultados no arquivo JSON
    salvar_dados_em_json(resultados)

    # Salvar RENAVAM sem multas em uma aba separada (opcional)
    if sem_multas:
        pd.DataFrame(sem_multas).to_excel("SemMultas.xlsx", index=False)

    print("Processo concluído!")
    driver.quit()

if __name__ == "__main__":
    main()
