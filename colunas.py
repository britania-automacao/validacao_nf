import pandas as pd

# Carrega os dois relatórios
df_antigo = pd.read_csv("dados/entrada_totvs/REB405 1.csv", sep=";", encoding="latin1", nrows=0)
df_novo = pd.read_excel("dados/entrada_totvs/REB405_142852.xlsx", nrows=0) 

# Transforma as colunas em conjuntos (sets)
colunas_antigo = set(df_antigo.columns)
colunas_novo = set(df_novo.columns)

print("--- DIAGNÓSTICO DE ESTRUTURA DO TOTVS ---")
print(f"Colunas que EXISTEM no CSV antigo, mas SUMIRAM no Excel novo:\n{colunas_antigo - colunas_novo}\n")
print(f"Colunas que SÃO NOVAS no Excel, e NÃO EXISTIAM no CSV antigo:\n{colunas_novo - colunas_antigo}")