import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests
import time
import os


# Função para resolver o CAPTCHA via 2Captcha
def resolver_captcha(api_key, site_key, page_url):
    print("Enviando CAPTCHA para resolução via 2Captcha...")

    response = requests.post(
        "http://2captcha.com/in.php",
        data={
            "key": api_key,
            "method": "userrecaptcha",
            "googlekey": site_key,
            "pageurl": page_url
        }
    )
    
    if "OK|" not in response.text:
        raise Exception(f"Erro ao enviar CAPTCHA: {response.text}")
    
    captcha_id = response.text.split("|")[1]
    print(f"CAPTCHA enviado. ID: {captcha_id}")

    for _ in range(60):  # Aguarda até 60 segundos
        time.sleep(5)
        res = requests.get(f"http://2captcha.com/res.php?key={api_key}&action=get&id={captcha_id}")
        if res.text.startswith("OK|"):
            print("CAPTCHA resolvido com sucesso.")
            return res.text.split("|")[1]
        elif res.text == "ERROR_CAPTCHA_UNSOLVABLE":
            print("CAPTCHA impossível de resolver. Tentando novamente...")
            break

    raise Exception("Tempo excedido para resolver o CAPTCHA.")


# Função principal
def consulta_multas(renavam, cpf_cnpj):
    url = "https://www.detran.rj.gov.br/_monta_aplicacoes.asp?cod=11&tipo=consulta_multa"
    site_key = "6LeIxAcTAAAAAJcZVRqyHh71UMIEGNQ_MXjiZKhI"
    api_key = "13a1288df1b6d6145a00476d48bf1c2b"

    # Configurar as opções do Chrome
    options = uc.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-infobars")

    # Inicializar o driver
    driver = uc.Chrome(options=options)
    driver.get(url)

    try:
        print(f"\nIniciando consulta para RENAVAM: {renavam}, CPF/CNPJ: {cpf_cnpj}")

        # Fechar modal de cookies, se presente
        try:
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "adopt-accept-all-button"))
            ).click()
            print("Modal de cookies fechado com sucesso.")
        except Exception:
            print("Modal de cookies não encontrado ou não disponível. Continuando...")

        # Localizar o iframe do formulário
        try:
            form_iframe = WebDriverWait(driver, 15).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "frameConsulta"))
            )
            print("Contexto alterado para o iframe do formulário.")
        except Exception as e:
            print(f"Erro ao localizar o iframe do formulário: {e}")
            salvar_arquivos_debug(driver, renavam)
            return f"Erro ao acessar o formulário para RENAVAM {renavam}."

        # Preencher o campo RENAVAM
        try:
            campo_renavam = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.ID, "MultasRenavam"))
            )
            campo_renavam.clear()
            campo_renavam.send_keys(renavam)
            print("Campo RENAVAM preenchido.")
        except Exception as e:
            print(f"Erro ao preencher o campo RENAVAM: {e}")
            salvar_arquivos_debug(driver, renavam)
            return f"Erro ao preencher o RENAVAM {renavam}."

        # Preencher o campo CPF/CNPJ
        try:
            campo_cpf_cnpj = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.ID, "MultasCpfcnpj"))
            )
            campo_cpf_cnpj.clear()
            campo_cpf_cnpj.send_keys(cpf_cnpj)
            print("Campo CPF/CNPJ preenchido.")
        except Exception as e:
            print(f"Erro ao preencher o campo CPF/CNPJ: {e}")
            salvar_arquivos_debug(driver, renavam)
            return f"Erro ao preencher o CPF/CNPJ para RENAVAM {renavam}."

        # Voltar para o contexto principal
        driver.switch_to.default_content()

        # Resolver o CAPTCHA
        try:
            captcha_token = resolver_captcha(api_key, site_key, url)
        except Exception as e:
            print(f"Erro ao resolver o CAPTCHA: {e}")
            salvar_arquivos_debug(driver, renavam)
            return f"Erro ao resolver o CAPTCHA para RENAVAM {renavam}."

        # Inserir o token resolvido e verificar CAPTCHA
        try:
            print("Alterando para o contexto do iframe do CAPTCHA...")
            
            # Localizar o iframe do reCAPTCHA
            iframe = WebDriverWait(driver, 20).until(
                EC.frame_to_be_available_and_switch_to_it(
                    (By.XPATH, "//iframe[contains(@src, 'https://www.google.com/recaptcha/api2/anchor')]")
                )
            )
            print("Confirmação: Contexto alterado para o iframe do CAPTCHA com sucesso.")

            # Salvar o DOM do iframe do CAPTCHA
            with open("debug/dom_captcha.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            print("DOM do iframe do CAPTCHA salvo para depuração.")

            # Tentar localizar e clicar no botão "Não sou um robô"
            print("Tentando localizar o botão 'Não sou um robô'...")
            captcha_checkbox = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.ID, "recaptcha-anchor"))
            )
            captcha_checkbox.click()
            print("Botão 'Não sou um robô' clicado com sucesso.")

            # Inserir o token resolvido
            print("Tentando inserir o token resolvido no campo 'g-recaptcha-response'...")
            driver.execute_script(
                'document.getElementById("g-recaptcha-response").style.display = "block";'
                f'document.getElementById("g-recaptcha-response").innerHTML = "{captcha_token}";'
            )
            print("Token resolvido inserido no CAPTCHA com sucesso.")
        except Exception as e:
            print(f"Erro ao alterar para o contexto ou interagir com o CAPTCHA: {e}")
            salvar_arquivos_debug(driver, renavam)
            return f"Erro ao interagir com o CAPTCHA para RENAVAM {renavam}."

        # Voltar ao contexto principal e clicar no botão
        try:
            driver.switch_to.default_content()
            form_iframe = WebDriverWait(driver, 15).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "frameConsulta"))
            )
            print("Contexto alterado com sucesso para o iframe do formulário.")

            botao_consultar = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#btPesquisar"))
            )
            botao_consultar.click()
            print("Botão 'Consultar' clicado com sucesso.")
        except Exception as e:
            print(f"Erro ao clicar no botão 'Consultar': {e}")
            salvar_arquivos_debug(driver, renavam)
            return f"Erro ao clicar no botão 'Consultar' para RENAVAM {renavam}."

        # Aguardar resultados
        try:
            driver.switch_to.default_content()
            resultados_element = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="caixaGeral"]/div[2]/div[2]'))
            )
            resultados = resultados_element.text
            print("Consulta realizada com sucesso.")
            return resultados
        except Exception as e:
            print(f"Erro ao obter os resultados: {e}")
            salvar_arquivos_debug(driver, renavam)
            return f"Erro ao obter os resultados para RENAVAM {renavam}."

    except Exception as e:
        print(f"Erro inesperado durante a consulta: {e}")
        salvar_arquivos_debug(driver, renavam)
        return f"Erro geral ao realizar a consulta para RENAVAM {renavam}."

    finally:
        driver.quit()


# Função para salvar arquivos de depuração
def salvar_arquivos_debug(driver, renavam):
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    os.makedirs("debug", exist_ok=True)

    driver.save_screenshot(f"debug/screenshot_{renavam}_{timestamp}.png")
    with open(f"debug/dom_{renavam}_{timestamp}.html", "w", encoding="utf-8") as f:
        f.write(driver.page_source)

    print("Arquivos de depuração salvos.")
