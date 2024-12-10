import pandas as pd

def organizar_dados():
    """
    Organiza os dados extraídos de multas em um arquivo Excel no formato especificado.
    - Padroniza as colunas conforme o exemplo fornecido.
    - Cria uma aba "Sem Multas" e mantém os dados formatados.
    """
    # Arquivo de entrada gerado pela extração
    input_file = "resultados_organizados.xlsx"
    print(f"Lendo o arquivo {input_file}...")

    # Abrindo todas as abas do arquivo de entrada
    dataframes = pd.read_excel(input_file, sheet_name=None)  # Lê todas as abas

    # Arquivo de saída
    output_file = "resultados_final.xlsx"
    print(f"Organizando os dados para o arquivo {output_file}...")

    # Estrutura padrão de colunas
    colunas_ordenadas = [
        'Auto de Infração', 'Auto de Renainf', 'Data para Pagamento com Desconto',
        'Enquadramento da Infração', 'Data da Infração', 'Hora', 'Descrição',
        'Placa Relacionada', 'Local da Infração', 'Valor Original R$',
        'Valor a Ser Pago R$', 'Status de Pagamento', 'Órgão Emissor', 'Agente Emissor'
    ]

    # Estrutura padrão para a aba "Sem Multas"
    colunas_sem_multas = ['RENAVAM', 'CNPJ', 'Status']

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        for renavam, sheet_df in dataframes.items():
            print(f"Processando aba: {renavam}...")

            if renavam == "Sem Multas":
                # Garantir que a aba "Sem Multas" siga o formato esperado
                sheet_df.columns = sheet_df.columns.str.strip()  # Remover espaços em branco nos nomes das colunas
                for coluna in colunas_sem_multas:
                    if coluna not in sheet_df.columns:
                        sheet_df[coluna] = ""
                sheet_df = sheet_df[colunas_sem_multas]
                sheet_df.to_excel(writer, sheet_name=renavam, index=False)
                continue

            # Garantir que todas as colunas necessárias estejam no DataFrame
            sheet_df.columns = sheet_df.columns.str.strip()  # Remover espaços em branco nos nomes das colunas
            for coluna in colunas_ordenadas:
                if coluna not in sheet_df.columns:
                    sheet_df[coluna] = ""

            # Processar valores e normalizar dados
            sheet_df['Valor Original R$'] = pd.to_numeric(sheet_df['Valor Original R$'], errors='coerce').fillna(0)
            sheet_df['Valor a Ser Pago R$'] = pd.to_numeric(sheet_df['Valor a Ser Pago R$'], errors='coerce').fillna(0)

            # Ordenar as colunas
            df_multas = sheet_df[colunas_ordenadas]

            # Escrever no arquivo de saída
            df_multas.to_excel(writer, sheet_name=str(renavam), index=False)

    print(f"Dados organizados e salvos em: {output_file}")
