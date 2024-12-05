import sys
import os
from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from twocaptcha import TwoCaptcha
import pandas as pd
import time

# Configurações globais
FRAME_CONSULTA_XPATH = "//*[@id='frameConsulta']"
CAPTCHA_IFRAME_XPATH = "//*[@id='divCaptcha']/div/div/iframe"
RECAPTCHA_BUTTON_XPATH = "//*[@id='recaptcha-anchor']"
CONSULTAR_BUTTON_XPATH = "//*[@id='btPesquisar']"
API_KEY = "13a1288df1b6d6145a00476d48bf1c2b"

def resolver_captcha(site_key, page_url):
    solver = TwoCaptcha(API_KEY)
    print("Resolving CAPTCHA via 2Captcha...")
    try:
        result = solver.recaptcha(sitekey=site_key, url=page_url)
        print("CAPTCHA resolved:", result)
        return result['code']
    except Exception as e:
        print("CAPTCHA resolution failed:", e)
        raise

def encontrar_g_recaptcha_response(driver):
    """Busca o elemento 'g-recaptcha-response' no HTML e nos iframes."""
    print("Searching for 'g-recaptcha-response' element in the page and all iframes...")
    
    # Primeiro, tenta localizar o elemento no HTML principal da página
    try:
        recaptcha_response_element = driver.find_element(By.ID, "g-recaptcha-response")
        print("'g-recaptcha-response' element found in the main HTML.")
        return recaptcha_response_element  # Retorna o elemento encontrado
    except Exception:
        print("'g-recaptcha-response' element not found in the main HTML.")

    # Se não encontrar no HTML principal, tenta localizar nos iframes
    driver.switch_to.default_content()  # Garante que começamos do contexto principal
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for iframe in iframes:
        try:
            driver.switch_to.frame(iframe)
            print(f"Checking iframe: {iframe.get_attribute('src')}")
            # Tenta localizar o elemento dentro do iframe atual
            recaptcha_response_element = driver.find_element(By.ID, "g-recaptcha-response")
            print("'g-recaptcha-response' element found in iframe.")
            return recaptcha_response_element  # Retorna o elemento encontrado
        except Exception:
            # Se não encontrar o elemento, volta ao contexto principal e tenta o próximo iframe
            driver.switch_to.default_content()
            continue
    
    print("Error: 'g-recaptcha-response' element not found in the page or any iframe!")
    raise Exception("'g-recaptcha-response' element not found in the page or any iframe!")

def inserir_token_captcha(driver):
    try:
        # Alterar para o iframe principal (frameConsulta)
        print("Switching to main iframe (frameConsulta)...")
        WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, FRAME_CONSULTA_XPATH))
        )

        # Alterar para o iframe do CAPTCHA para interagir com o checkbox
        print("Switching to reCAPTCHA iframe to interact with the checkbox...")
        WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, CAPTCHA_IFRAME_XPATH))
        )

        # Clicar no botão "Eu não sou um robô"
        print("Clicking reCAPTCHA checkbox...")
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, RECAPTCHA_BUTTON_XPATH))
        ).click()

        # Voltar ao contexto do iframe principal (frameConsulta)
        print("Switching back to main iframe (frameConsulta)...")
        driver.switch_to.default_content()
        driver.switch_to.frame(driver.find_element(By.XPATH, FRAME_CONSULTA_XPATH))

        # Obter o sitekey do CAPTCHA
        print("Extracting sitekey from the main iframe...")
        sitekey_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='divCaptcha']"))
        )
        site_key = sitekey_element.get_attribute("data-sitekey")
        print(f"Sitekey extracted: {site_key}")

        # Resolver o CAPTCHA usando a API 2Captcha
        page_url = driver.current_url
        captcha_token = resolver_captcha(site_key, page_url)

        # Localizar o elemento 'g-recaptcha-response' dinamicamente (tanto no HTML quanto nos iframes)
        recaptcha_response_element = encontrar_g_recaptcha_response(driver)

        # Inserir o token no elemento encontrado
        print("Inserting token into 'g-recaptcha-response'...")
        driver.execute_script(""" 
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """, recaptcha_response_element, captcha_token)

        # Confirmar se o valor foi atualizado corretamente no elemento
        token_value = driver.execute_script("return arguments[0].value;", recaptcha_response_element)
        if token_value == captcha_token:
            print("Token successfully inserted into 'g-recaptcha-response'.")
        else:
            print(f"Error: Token insertion failed! Expected: {captcha_token}, but found: {token_value}")
            raise Exception("Token insertion verification failed!")

        # Confirmar que o CAPTCHA foi validado com sucesso
        print("Checking for reCAPTCHA checkmark...")
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "recaptcha-checkbox-checkmark"))
        )
        print("reCAPTCHA validated successfully.")

        # Voltar ao iframe principal antes de clicar no botão CONSULTAR
        print("Switching back to main iframe (frameConsulta) before clicking CONSULTAR...")
        driver.switch_to.default_content()
        driver.switch_to.frame(driver.find_element(By.XPATH, FRAME_CONSULTA_XPATH))

        # Remover o overlay que está bloqueando o clique no botão CONSULTAR
        print("Removing potential overlays blocking the CONSULTAR button...")
        driver.execute_script(""" 
            var overlay = document.querySelector('div[style*="z-index: 2000000000"]');
            if (overlay) {
                overlay.remove();
            }
        """)

        # Clicar no botão CONSULTAR
        print("Clicking 'CONSULTAR' button...")
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, CONSULTAR_BUTTON_XPATH))
        ).click()
        print("Button clicked.")

    except Exception as e:
        print("Error inserting CAPTCHA token:", e)
        raise


def consulta_multas(renavam, cpf_cnpj):
    url = "https://www.detran.rj.gov.br/_monta_aplicacoes.asp?cod=11&tipo=consulta_multa"

    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(60)

    try:
        print(f"Processing RENAVAM: {renavam}, CPF/CNPJ: {cpf_cnpj}")
        driver.get(url)
        
        # Aceitar cookies, se necessário
        try:
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='adopt-accept-all-button']"))
            ).click()
        except:
            print("No cookies to accept.")

        # Preencher o formulário
        driver.switch_to.frame("frameConsulta")
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "MultasRenavam"))
        ).send_keys(str(renavam))
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "MultasCpfcnpj"))
        ).send_keys(str(cpf_cnpj))
        driver.switch_to.default_content()

        # Inserir CAPTCHA
        inserir_token_captcha(driver)

        # Verificar o resultado
        print("Waiting for results...")
        driver.switch_to.default_content()
        driver.switch_to.frame("frameConsulta")
        result_element = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='caixaInformacao']"))
        )
        result = result_element.text
        print("Results loaded successfully:", result)
        return result
    except Exception as e:
        print("Error:", e)
        return f"Error processing RENAVAM {renavam}"
    finally:
        driver.quit()

if __name__ == "__main__":
    input_file = "Consulta Multas.xlsx"
    output_file = "Resultados Multas.xlsx"

    try:
        df = pd.read_excel(input_file)
        results = []
        for _, row in df.iterrows():
            renavam = row["RENAVAM"]
            cpf_cnpj = row["CPF/CNPJ"]
            result = consulta_multas(renavam, cpf_cnpj)
            results.append({"RENAVAM": renavam, "CPF/CNPJ": cpf_cnpj, "Result": result})
        pd.DataFrame(results).to_excel(output_file, index=False)
        print("Results saved.")
    except Exception as e:
        print("Error:", e)
