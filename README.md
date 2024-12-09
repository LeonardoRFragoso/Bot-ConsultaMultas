
# Bot-ConsultaMultas

Um bot automatizado para consulta de multas no site do Detran-RJ utilizando Selenium, 2Captcha e dados extraídos de uma planilha Excel.

## Funcionalidades
- Preenchimento automático de formulários de consulta no site do Detran-RJ.
- Resolução automática de CAPTCHAs utilizando a API do 2Captcha.
- Processamento de dados de entrada a partir de uma planilha Excel.
- Exportação de resultados para um arquivo Excel.

---

## Requisitos

- **Python 3.8 ou superior**a
- **Google Chrome** instalado
- **ChromeDriver** compatível com a versão do navegador
- API Key da **2Captcha**

---

## Instalação

1. Clone este repositório:
   ```bash
   git clone https://github.com/LeonardoRFragoso/Bot-ConsultaMultas.git
   cd Bot-ConsultaMultas
   ```

2. Crie um ambiente virtual:
   ```bash
   python -m venv venv
   ```

3. Ative o ambiente virtual:

   **Windows**:
   ```bash
   venv\Scripts\activate
   ```

   **Linux/Mac**:
   ```bash
   source venv/bin/activate
   ```

4. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

---

## Como usar

1. Insira os dados de entrada na planilha `Consulta Multas.xlsx` (colunas de exemplo: `RENAVAN`, `CNPJ`, `PLACA`).

2. Adicione sua API Key da 2Captcha no arquivo `app.py`:
   ```python
   api_key = "SUA_API_KEY_AQUI"
   ```

3. Execute o script:
   ```bash
   python app.py
   ```

4. Após a execução, os resultados estarão disponíveis no arquivo `resultados_multas.xlsx`.

---

## Estrutura do Projeto

```
Bot-ConsultaMultas/
├── Consulta Multas.xlsx      # Planilha de entrada
├── README.md                 # Documentação do projeto
├── app.py                    # Script principal
├── requirements.txt          # Dependências do projeto
└── venv/                     # Ambiente virtual (ignorado no Git)
```

---

## Contribuição

1. Faça um fork do repositório.
2. Crie uma branch para a sua feature:
   ```bash
   git checkout -b minha-feature
   ```
3. Commit suas alterações:
   ```bash
   git commit -m "Descrição da feature"
   ```
4. Faça um push para a branch:
   ```bash
   git push origin minha-feature
   ```
5. Abra um Pull Request.

---

## Observações

- Certifique-se de que o **ChromeDriver** está na sua variável de ambiente `PATH`.
- Para evitar problemas, mantenha o **Chrome** atualizado e utilize a versão correspondente do **ChromeDriver**.

---

## Licença

Este projeto é licenciado sob a MIT License.

---

### O que foi atualizado no README:

1. **Instruções sobre como configurar o ambiente virtual**.
2. **Dependência de uma API Key da 2Captcha e como configurá-la**.
3. **Uso claro da planilha de entrada e arquivo de saída**.
4. **Adição de estrutura de projeto** para facilitar a navegação.
5. **Passos para contribuição** para que outros possam colaborar com o projeto.
