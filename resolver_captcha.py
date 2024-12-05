from selenium import webdriver
from selenium.webdriver.common.by import By
import time

# Inicializa o driver
driver = webdriver.Chrome()

# Acesse o site
driver.get('https://www.detran.rj.gov.br/_monta_aplicacoes.asp?cod=11&tipo=consulta_multa')

# Aguarde o CAPTCHA ser resolvido manualmente
print("Por favor, resolva o reCAPTCHA manualmente no navegador.")
time.sleep(60)

# Captura o texto exibido no resultado
try:
    resultado = driver.find_element(By.XPATH, "//div[contains(text(), 'Não há multa')]").text
    print("Resultado capturado:", resultado)
except Exception as e:
    print("Erro ao capturar o resultado:", e)

# Finaliza o driver
driver.quit()
