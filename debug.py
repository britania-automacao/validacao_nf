import os
import sys
import pandas as pd
from pathlib import Path

# Adiciona a pasta src para puxar suas configurações reais
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))
from config import CaminhosProjeto

raiz = Path(__file__).resolve().parent
caminhos = CaminhosProjeto(diretorio_base=raiz)
pasta_ofb = caminhos.entrada_totvs / "OFB013"

print(f"Procurando arquivos na pasta: {pasta_ofb}")
arquivos = list(pasta_ofb.rglob("*.csv")) + list(pasta_ofb.rglob("*.xlsx"))

if not arquivos:
    print("[ERRO] Nenhum arquivo foi achado nessa pasta.")
else:
    # Vamos ler o primeiro arquivo que ele achar
    arquivo_alvo = arquivos[0]
    print(f"Lendo o arquivo: {arquivo_alvo.name}")
    
    try:
        # Lê o CSV da mesma forma que a automação
        df = pd.read_csv(arquivo_alvo, sep=';', dtype=str, encoding='latin1', on_bad_lines='skip')
        df.columns = [str(c).replace('"', '').replace("'", "").strip() for c in df.columns]
        
        if 'Chave Acesso' in df.columns:
            print("\n[SUCESSO] Coluna 'Chave Acesso' encontrada no layout!")
            print("Mostrando os primeiros valores reais de chave de acesso encontrados:")
            print("--------------------------------------------------")
            
            # Pega as chaves ignorando o que estiver vazio
            chaves = df['Chave Acesso'].dropna().astype(str).str.strip()
            chaves_preenchidas = chaves[chaves != ''].head(10).tolist()
            
            if not chaves_preenchidas:
                print("ALERTA CRÍTICO: A coluna 'Chave Acesso' EXISTE no TOTVS,")
                print("mas TODAS AS LINHAS estão 100% VAZIAS!")
                print("É por isso que o robô zerou a base.")
            else:
                for v in chaves_preenchidas:
                    print(f"-> {v}")
            print("--------------------------------------------------")
        else:
            print("\n[ERRO FATAL] A coluna 'Chave Acesso' NÃO EXISTE NESSE ARQUIVO!")
            print("Verifique se o layout exportado do TOTVS não mudou o nome da coluna.")
            print("Colunas que o robô conseguiu enxergar:")
            print(df.columns.tolist())
            
    except Exception as e:
        print(f"Falha ao ler o arquivo: {e}")