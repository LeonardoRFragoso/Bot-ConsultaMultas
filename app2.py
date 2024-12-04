import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests
import pandas as pd
import time

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

    # Aguardar a solução do CAPTCHA
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

    options = uc.ChromeOptions()
    options.add_argument("--start-maximized")

    driver = uc.Chrome(options=options)
    driver.get(url)

    try:
        print(f"Iniciando consulta para RENAVAM: {renavam}, CPF/CNPJ: {cpf_cnpj}")

        # Fechar modal de cookies
        try:
            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="adopt-accept-all-button"]'))
            ).click()
            print("Modal de cookies fechado com sucesso.")
        except Exception:
            print("Modal de cookies não encontrado ou não disponível. Continuando...")

        # Verificar e listar todos os iframes
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        print(f"Número de iframes na página: {len(iframes)}")

        # Alterar contexto para o iframe do formulário
        form_iframe = next(
            (iframe for iframe in iframes if "nadaConsta" in (iframe.get_attribute("src") or "")),
            None
        )

        if form_iframe:
            driver.switch_to.frame(form_iframe)
            print("Contexto alterado para o iframe do formulário.")
        else:
            print("Erro: Nenhum iframe correspondente encontrado.")
            return "Erro no iframe"

        # Preencher os campos
        campo_renavam = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "MultasRenavam"))
        )
        campo_renavam.clear()
        campo_renavam.send_keys(renavam)
        print("Campo RENAVAM preenchido.")

        campo_cpf_cnpj = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "MultasCpfcnpj"))
        )
        campo_cpf_cnpj.clear()
        campo_cpf_cnpj.send_keys(cpf_cnpj)
        print("Campo CPF/CNPJ preenchido.")

        driver.switch_to.default_content()

        # Resolver o CAPTCHA
        captcha_token = resolver_captcha(api_key, site_key, url)

        captcha_iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'recaptcha')]"))
        )
        driver.switch_to.frame(captcha_iframe)
        driver.execute_script(
            f'document.getElementById("g-recaptcha-response").innerHTML = "{captcha_token}";'
        )
        print("CAPTCHA resolvido e inserido no formulário.")

        driver.switch_to.default_content()

        # Aguarde para garantir que o botão esteja ativo
        time.sleep(3)

        botao_consultar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="btPesquisar"]'))
        )
        botao_consultar.click()
        print("Formulário enviado.")

        # Aguardar os resultados
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "resultado"))
        )
        resultados = driver.find_element(By.ID, "resultado").text
        print("Consulta realizada com sucesso.")
        return resultados

    except Exception as e:
        print(f"Erro geral durante a consulta: {e}")
        return "Erro geral"

    finally:
        driver.quit()

# Carregar dados da planilha Excel
file_path = "Consulta Multas.xlsx"
veiculos_df = pd.read_excel(file_path, sheet_name="Planilha1")

# Processar consultas
resultados = []
for _, row in veiculos_df.iterrows():
    renavam = row['RENAVAN']
    cpf_cnpj = row['CNPJ']
    resultado = consulta_multas(renavam, cpf_cnpj)
    resultados.append({
        'PLACA': row['PLACA'],
        'RENAVAN': renavam,
        'CNPJ': cpf_cnpj,
        'Resultado': resultado
    })

# Salvar resultados em Excel
resultados_df = pd.DataFrame(resultados)
resultados_df.to_excel("resultados_multas.xlsx", index=False)
print("Consulta finalizada. Resultados salvos em 'resultados_multas.xlsx'.")
