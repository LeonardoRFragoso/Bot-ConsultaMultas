import pandas as pd
from consulta_multas import consulta_multas

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
