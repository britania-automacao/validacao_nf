# src/services_etl.py
import os
import re
import time
import requests
import urllib3
import pandas as pd
import xml.etree.ElementTree as ET
from pathlib import Path
from dotenv import load_dotenv

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from config import CaminhosProjeto

import logging
import sys

# --- INÍCIO DA CONFIGURAÇÃO DE LOG ---
pasta_historico = Path(__file__).resolve().parent.parent / "dados" / "historico"
pasta_historico.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(message)s',
    handlers=[
        logging.FileHandler(pasta_historico / "etl_servicos.log", encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

def print(*args, **kwargs):
    sep = kwargs.get("sep", " ")
    logger.info(sep.join(map(str, args)))
# --- FIM DA CONFIGURAÇÃO DE LOG ---

# Desativa os avisos de requisição insegura no terminal
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

load_dotenv()

class LeitorSimplesNacional:
    def __init__(self, diretorio_base: Path):
        self.pasta = diretorio_base / "dados" / "consulta_simples_nacional"
        self.pasta.mkdir(parents=True, exist_ok=True)

    def _limpar_cnpj(self, cnpj_bruto) -> str:
        if pd.isna(cnpj_bruto): return ""
        val = str(cnpj_bruto).strip()
        if val.endswith('.0'): val = val[:-2]
        if 'E' in val.upper(): 
            try: val = str(int(float(val)))
            except: pass
        
        limpo = ''.join(filter(str.isdigit, val))
        return limpo.zfill(14) if limpo else ""

    def _consultar_api(self, cnpj: str) -> str:
        """Consulta APIs de CNPJ com sistema de fallback (Cascata) obedecendo a ordem solicitada."""
        if not cnpj or len(cnpj) != 14: 
            return "PENDENTE_CONSULTA"
            
        # 1. ReceitaWS | 2. Minha Receita | 3. BrasilAPI
        apis = [
            {"nome": "ReceitaWS", "url": f"https://receitaws.com.br/v1/cnpj/{cnpj}"},
            {"nome": "MinhaReceita", "url": f"https://minhareceita.org/{cnpj}"},
            {"nome": "BrasilAPI", "url": f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}"}
        ]

        for api in apis:
            nome_api = api["nome"]
            url = api["url"]
            
            try:
                response = requests.get(url, timeout=10, verify=False)
                
                if response.status_code == 200:
                    dados = response.json()
                    
                    # Tratamento específico para a estrutura JSON do ReceitaWS
                    if nome_api == "ReceitaWS":
                        if dados.get("status") == "ERROR":
                            print(f"  [API AVISO] {nome_api} retornou erro (CNPJ inválido/rejeitado). Pulando para a próxima...")
                            continue
                            
                        simples = dados.get("simples")
                        if isinstance(simples, dict):
                            opcao = simples.get("optante")
                            return "SIM" if opcao is True else "NÃO"
                        return "NÃO"
                        
                    # Tratamento para Minha Receita e BrasilAPI (ambas usam 'opcao_pelo_simples')
                    elif nome_api in ["MinhaReceita", "BrasilAPI"]:
                        opcao = dados.get("opcao_pelo_simples")
                        return "SIM" if opcao is True else "NÃO"
                        
                elif response.status_code == 429:
                    print(f"  [API AVISO] Limite atingido (429) na {nome_api} para o CNPJ {cnpj}. Acionando fallback...")
                    continue 
                    
                else:
                    print(f"  [API AVISO] {nome_api} retornou erro {response.status_code}. Acionando fallback...")
                    continue 
                    
            except Exception as e:
                print(f"  [API ERRO] Falha de conexão na {nome_api}: {e}. Acionando fallback...")
                continue
                
        # Se as 3 APIs falharem para o mesmo CNPJ, mantém pendente para a próxima execução
        print(f"  [ALERTA MÁXIMO] Todas as 3 APIs falharam para o CNPJ {cnpj}.")
        return "PENDENTE_CONSULTA"

    def carregar(self, cnpjs_alvo=None) -> pd.DataFrame:
        arquivos = list(self.pasta.glob("*.xlsx")) + list(self.pasta.glob("*.csv"))
        
        df_original = pd.DataFrame()
        col_cnpj, col_simples = "CNPJ", "SIMPLES"
        caminho_salvar = None

        if arquivos:
            caminho_salvar = arquivos[0]
            try:
                print(f"[ETL SIMPLES] Lendo base de Consulta Simples Nacional: {caminho_salvar.name}")
                if caminho_salvar.suffix.lower() == '.xlsx':
                    df_original = pd.read_excel(caminho_salvar, dtype=str)
                else:
                    df_original = pd.read_csv(caminho_salvar, sep=',', encoding='utf-8', dtype=str)
                    if len(df_original.columns) == 1:
                        df_original = pd.read_csv(caminho_salvar, sep=';', encoding='latin1', dtype=str)
                        
                col_cnpj = [c for c in df_original.columns if 'CNPJ' in c.upper()][0]
                col_simples = [c for c in df_original.columns if 'SIMPLES' in c.upper()][0]
                
                df_original['CNPJ_Limpo'] = df_original[col_cnpj].apply(self._limpar_cnpj)
                df_original['Simples_SN'] = df_original[col_simples].str.strip().str.upper()
                
            except Exception as e:
                print(f"[ERRO SIMPLES] Falha ao estruturar a base do Simples Nacional: {e}")
                df_original = pd.DataFrame(columns=['CNPJ_Limpo', 'Simples_SN'])
        else:
            print("[AVISO] Nenhum arquivo de Simples Nacional encontrado. Criando base zerada...")
            caminho_salvar = self.pasta / "consulta_simples_nacional.xlsx"
            df_original = pd.DataFrame(columns=['CNPJ', 'SIMPLES', 'CNPJ_Limpo', 'Simples_SN'])

        # --- LÓGICA DE API (AUTO-HEALING) ---
        if cnpjs_alvo:
            cnpjs_existentes = set()
            if not df_original.empty and 'CNPJ_Limpo' in df_original.columns:
                
                # Remove pendentes para forçar nova consulta
                df_original = df_original[df_original['Simples_SN'] != 'PENDENTE_CONSULTA'].copy()
                cnpjs_existentes = set(df_original['CNPJ_Limpo'].dropna().tolist())

            cnpjs_faltantes = {c for c in cnpjs_alvo if isinstance(c, str) and len(c) == 14} - cnpjs_existentes

            if cnpjs_faltantes:
                print(f"[API SIMPLES] {len(cnpjs_faltantes)} CNPJs pendentes/novos detectados! Consultando cascata de APIs...")
                novos_registros = []
                
                for idx, cnpj in enumerate(cnpjs_faltantes, 1):
                    status = self._consultar_api(cnpj)
                    nova_linha = {
                        col_cnpj: cnpj, 
                        col_simples: status,
                        'CNPJ_Limpo': cnpj,
                        'Simples_SN': status
                    }
                    novos_registros.append(nova_linha)
                    
                    if idx % 10 == 0:
                        print(f"  -> Consultados {idx}/{len(cnpjs_faltantes)}...")
                    
                    # REGRA APLICADA: Pausa rígida de 20 segundos entre consultas.
                    # Matemática: 60 segundos / 20 = Exatamente 3 requisições por minuto.
                    # Isso garante que a ReceitaWS e a MinhaReceita nunca estourem o limite.
                    time.sleep(20) 

                df_novos = pd.DataFrame(novos_registros)
                df_original = pd.concat([df_original, df_novos], ignore_index=True)
                
                try:
                    df_salvar = df_original.drop(columns=['CNPJ_Limpo', 'Simples_SN'], errors='ignore')
                    df_salvar[col_cnpj] = df_salvar[col_cnpj].apply(lambda x: str(x).zfill(14) if pd.notna(x) else x)
                    if caminho_salvar.suffix.lower() == '.xlsx':
                        df_salvar.to_excel(caminho_salvar, index=False)
                    else:
                        df_salvar.to_csv(caminho_salvar, sep=';', index=False, encoding='utf-8-sig')
                    print(f"[SUCESSO] Arquivo '{caminho_salvar.name}' atualizado permanentemente com os novos CNPJs!")
                except Exception as e:
                    print(f"[ERRO] Não foi possível salvar o arquivo atualizado: {e}")

        if not df_original.empty and 'CNPJ_Limpo' in df_original.columns:
            df_retorno = df_original.drop_duplicates(subset=['CNPJ_Limpo'], keep='last')
            return df_retorno[['CNPJ_Limpo', 'Simples_SN']]
        else:
            return pd.DataFrame(columns=['CNPJ_Limpo', 'Simples_SN'])


class GeradorCsvGoverno:
    def __init__(self, caminhos: CaminhosProjeto):
        self.caminhos = caminhos
        self.arquivo_saida_csv = self.caminhos.diretorio_base / "dados" / "Base_XML_Governo_Consolidada.csv"

    def _get_node(self, raiz, tags_possiveis):
        for elem in raiz.iter():
            local_name = elem.tag.split('}')[-1]
            if local_name in tags_possiveis:
                return elem
        return None

    def _get_text(self, raiz, tags_possiveis, default=None):
        nodo = self._get_node(raiz, tags_possiveis)
        if nodo is not None and nodo.text:
            return nodo.text.strip()
        return default

    def _get_float(self, raiz, tags_possiveis):
        valor_str = self._get_text(raiz, tags_possiveis)
        if valor_str:
            try: return float(valor_str)
            except ValueError: return 0.0
        return 0.0

    def _normalizar_numero_nota(self, numero_bruto: str) -> str:
        if not numero_bruto: return None
        num_limpo = ''.join(filter(str.isdigit, str(numero_bruto)))
        if not num_limpo: return str(numero_bruto).strip()
        return str(int(num_limpo))

    def _achatar_xml(self, nodo, prefixo=''):
        dados = {}
        if nodo.attrib:
            for attr_nome, attr_valor in nodo.attrib.items():
                chave = f"{prefixo}_{attr_nome}" if prefixo else attr_nome
                dados[chave] = attr_valor
                
        for filho in nodo:
            tag = filho.tag.split('}')[-1]
            chave = f"{prefixo}_{tag}" if prefixo else tag
            
            if len(filho) > 0:
                dados.update(self._achatar_xml(filho, chave))
            
            if filho.text and filho.text.strip():
                dados[chave] = filho.text.strip()
                
        return dados

    def gerar_csv(self) -> Path:
        print("[ETL XML] Verificando XMLs Nacionais para atualização da Base CSV...")
        lista_arquivos_xml = list(self.caminhos.xml_servicos_nacional.rglob("*.xml"))
        
        if not lista_arquivos_xml:
            print("[AVISO] A pasta de XMLs está vazia.")
            return None

        arquivos_ja_processados = set()
        if self.arquivo_saida_csv.exists():
            try:
                df_existente = pd.read_csv(self.arquivo_saida_csv, sep=';', usecols=['Origem_XML'], encoding='utf-8-sig')
                arquivos_ja_processados = set(df_existente['Origem_XML'].dropna().tolist())
                print(f"[ETL XML] A base atual possui {len(arquivos_ja_processados)} notas registadas.")
            except Exception as e:
                print(f"[AVISO] Não foi possível ler o CSV atual. Erro: {e}")

        registros_novos = []
        for caminho in lista_arquivos_xml:
            if caminho.name in arquivos_ja_processados:
                continue

            try:
                arvore = ET.parse(caminho)
                raiz = arvore.getroot()

                # NOVO: Procura múltiplas notas (Lotes) dentro do mesmo XML
                nodos_nota = []
                for elem in raiz.iter():
                    tag_limpa = elem.tag.split('}')[-1].lower()
                    if tag_limpa in ['nfse', 'nfs-e']:
                        nodos_nota.append(elem)
                
                # Fallback para padrão ABRASF se não achar o padrão Nacional
                if not nodos_nota:
                    for elem in raiz.iter():
                        tag_limpa = elem.tag.split('}')[-1].lower()
                        if tag_limpa == 'compnfse':
                            nodos_nota.append(elem)
                            
                # Fallback final: se não achar blocos, assume que a raiz é a nota
                if not nodos_nota:
                    nodos_nota = [raiz]

                # Itera sobre cada nota encontrada dentro do arquivo
                for nodo in nodos_nota:
                    registro = self._achatar_xml(nodo)

                    nome_arquivo = caminho.name
                    match_nsu = re.search(r'nsu_(\d+)', nome_arquivo, re.IGNORECASE)
                    nsu_xml = match_nsu.group(1) if match_nsu else ( ''.join(filter(str.isdigit, nome_arquivo)) or caminho.stem )

                    numero_bruto = self._get_text(nodo, ['nNFSe', 'Numero', 'nNFS-e'])
                    numero_normalizado = self._normalizar_numero_nota(numero_bruto)

                    cnpj = None
                    nome_xml = "Fornecedor Desconhecido"
                    
                    nodo_forn = self._get_node(nodo, ['prest', 'PrestadorServico', 'emit', 'Emitente'])
                    if nodo_forn is not None:
                        cnpj_bruto = self._get_text(nodo_forn, ['CNPJ', 'Cnpj'])
                        if cnpj_bruto: 
                            cnpj = ''.join(filter(str.isdigit, str(cnpj_bruto))).zfill(14)
                        nome_xml = self._get_text(nodo_forn, ['xNome', 'RazaoSocial'], "Fornecedor Desconhecido")

                    dh_emi = self._get_text(nodo, ['dhEmi', 'DataEmissao', 'Competencia'])
                    data_emissao = dh_emi[:10] if dh_emi and len(dh_emi) >= 10 else ""
                    
                    valor_bruto = self._get_float(nodo, ['vServ', 'ValorServicos'])
                    if valor_bruto == 0.0: 
                        valor_bruto = self._get_float(nodo, ['vBC', 'BaseCalculo'])

                    val_nacional = self._get_text(nodo, ['opSimpNac'])
                    val_abrasf = self._get_text(nodo, ['OptanteSimplesNacional'])
                    
                    simples_xml_resolvido = "NÃO"
                    if val_nacional in ['2', '3', '4']: 
                        simples_xml_resolvido = "SIM"
                    elif val_nacional == '1': 
                        simples_xml_resolvido = "NÃO"
                    elif val_abrasf == '1':   
                        simples_xml_resolvido = "SIM"
                    elif val_abrasf == '2':
                        simples_xml_resolvido = "NÃO"

                    flag_ret_iss = self._get_text(nodo, ['tpRetISSQN', 'IssRetido'], '1')

                    if cnpj and numero_normalizado:
                        # Extrai o Ano e Mês (YYYY-MM) para a chave
                        mes_ano_xml = data_emissao[:7] if len(data_emissao) >= 7 else "0000-00"
                        chave = f"{cnpj}_{numero_normalizado}_{mes_ano_xml}"
                        
                        registro["Chave"] = chave
                        registro["NSU_XML"] = nsu_xml
                        registro["CNPJ_XML"] = cnpj
                        registro["Nota_XML"] = numero_normalizado
                        registro["Nome_XML"] = nome_xml
                        registro["DataEmissao_XML"] = data_emissao
                        registro["Valor_XML"] = valor_bruto
                        registro["BaseCalculo_XML"] = self._get_float(nodo, ['vBC', 'BaseCalculo'])
                        registro["ISS_XML"] = self._get_float(nodo, ['vISSQN', 'ValorIss'])
                        registro["Flag_Ret_ISS"] = flag_ret_iss
                        registro["Simples_XML_Resolvido"] = simples_xml_resolvido
                        registro["IRRF_XML"] = self._get_float(nodo, ['vRetIRRF', 'ValorIr'])
                        registro["INSS_XML"] = self._get_float(nodo, ['vRetCP', 'ValorInss'])
                        registro["PIS_XML"] = self._get_float(nodo, ['vRetPIS', 'ValorPis'])
                        registro["COFINS_XML"] = self._get_float(nodo, ['vRetCOFINS', 'ValorCofins'])
                        registro["CSLL_XML"] = self._get_float(nodo, ['vRetCSLL', 'ValorCsll'])
                        registro["Origem_XML"] = caminho.name
                        
                        registros_novos.append(registro)

            except Exception:
                pass 

        if registros_novos:
            df_novos = pd.DataFrame(registros_novos)
            
            if self.arquivo_saida_csv.exists() and len(arquivos_ja_processados) > 0:
                try:
                    df_existente = pd.read_csv(self.arquivo_saida_csv, sep=';', encoding='utf-8-sig', dtype=str)
                    df_final = pd.concat([df_existente, df_novos], ignore_index=True)
                    df_final.to_csv(self.arquivo_saida_csv, sep=';', index=False, encoding='utf-8-sig')
                    print(f"[ETL XML] Processamento Rápido e Seguro! +{len(df_novos)} NOVAS notas adicionadas à Base CSV.")
                except Exception as e:
                    print(f"[ERRO XML] Falha ao atualizar CSV de forma segura: {e}")
            else:
                df_novos.to_csv(self.arquivo_saida_csv, sep=';', index=False, encoding='utf-8-sig')
                print(f"[ETL XML] Sucesso! CSV Base do Governo criado do zero com {len(df_novos)} notas.")
        else:
            print("[ETL XML] Nenhuma nota nova na pasta. O CSV do Governo já está 100% atualizado!")
            
        return self.arquivo_saida_csv


class LeitorBaseGoverno:
    def __init__(self, caminho_csv: Path):
        self.caminho_csv = caminho_csv

    def carregar(self) -> pd.DataFrame:
        if not self.caminho_csv or not self.caminho_csv.exists():
            return pd.DataFrame()
            
        print("[ETL GOVERNO] Carregando dados da Base Consolidada (CSV)...")
        df = pd.read_csv(self.caminho_csv, sep=';', encoding='utf-8-sig', dtype=str)
        
        # Atualiza dinamicamente as chaves legadas do CSV para o novo padrão (CNPJ_Nota_Ano-Mes)
        if 'CNPJ_XML' in df.columns and 'Nota_XML' in df.columns and 'DataEmissao_XML' in df.columns:
            df['MesAno_XML'] = df['DataEmissao_XML'].apply(lambda x: str(x)[:7] if pd.notna(x) and len(str(x)) >= 7 else "0000-00")
            df['Chave'] = df['CNPJ_XML'] + "_" + df['Nota_XML'] + "_" + df['MesAno_XML']
        
        cols_numericas = ['Valor_XML', 'BaseCalculo_XML', 'ISS_XML', 'IRRF_XML', 'INSS_XML', 'PIS_XML', 'COFINS_XML', 'CSLL_XML']
        for col in cols_numericas:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
                
        return df


class LeitorRelatorioTotvs:
    def __init__(self, caminhos: CaminhosProjeto):
        self.pasta = caminhos.entrada_totvs / "REB405"
        self.pasta.mkdir(parents=True, exist_ok=True)

    def _limpar_cnpj(self, cnpj_bruto) -> str:
        if pd.isna(cnpj_bruto): return ""
        val = str(cnpj_bruto).strip()
        if val.endswith('.0'): val = val[:-2]
        if 'E' in val.upper(): 
            try: val = str(int(float(val)))
            except: pass
        limpo = ''.join(filter(str.isdigit, val))
        return limpo.zfill(14) if limpo else ""

    def _limpar_numero_nota(self, numero_bruto) -> str:
        if pd.isna(numero_bruto): return ""
        num_limpo = ''.join(filter(str.isdigit, str(numero_bruto)))
        if num_limpo: return str(int(num_limpo))
        return str(numero_bruto).strip()

    def _converter_moeda(self, df, nome_coluna):
        if nome_coluna not in df.columns:
            return 0.0
            
        def limpar_moeda(valor):
            if pd.isna(valor): return 0.0
            v_str = str(valor).strip()
            if ',' in v_str:
                v_str = v_str.replace('.', '').replace(',', '.')
            try:
                return float(v_str)
            except ValueError:
                return 0.0

        return df[nome_coluna].apply(limpar_moeda)

    def processar_relatorio(self) -> pd.DataFrame:
        print(f"[ETL TOTVS] A iniciar a leitura do ERP na pasta: {self.pasta.name}...")
        arquivos_totvs = list(self.pasta.rglob("*.xlsx")) + list(self.pasta.rglob("*.csv"))
        if not arquivos_totvs:
            print(f"[AVISO] Nenhum relatório do TOTVS encontrado na pasta {self.pasta.name}.")
            return pd.DataFrame()

        lista_dfs = []

        for caminho_relatorio in arquivos_totvs:
            try:
                print(f"[ETL TOTVS] A processar arquivo: {caminho_relatorio.name}")
                if caminho_relatorio.suffix.lower() == '.xlsx':
                    df_temp = pd.read_excel(caminho_relatorio, dtype=str)
                else:
                    try:
                        df_temp = pd.read_csv(caminho_relatorio, sep=';', encoding='utf-8', dtype=str)
                    except UnicodeDecodeError:
                        df_temp = pd.read_csv(caminho_relatorio, sep=';', encoding='latin1', dtype=str)
                
                lista_dfs.append(df_temp)
            except Exception as e:
                print(f"[ERRO TOTVS] Falha ao ler o arquivo {caminho_relatorio.name}: {e}")

        if not lista_dfs:
            print("[ERRO TOTVS] Nenhum arquivo pôde ser lido com sucesso.")
            return pd.DataFrame()

        df_totvs = pd.concat(lista_dfs, ignore_index=True)

        try:
            df_totvs.dropna(how='all', inplace=True)
            df_totvs.columns = df_totvs.columns.str.strip().str.replace('\n', ' ').str.replace('\r', '')

            if 'Docto' in df_totvs.columns and 'CNPJ emit' in df_totvs.columns:
                df_totvs['Nota_TOTVS'] = df_totvs['Docto'].apply(self._limpar_numero_nota)
                df_totvs['CNPJ_TOTVS'] = df_totvs['CNPJ emit'].apply(self._limpar_cnpj)
                
                # Extrai o Mês e Ano da Dt.Emiss do TOTVS para compor a chave (Converte DD/MM/YYYY para YYYY-MM)
                def extract_mes_ano(dt_str):
                    if pd.isna(dt_str) or not dt_str: return "0000-00"
                    dt_str = str(dt_str).strip()
                    if len(dt_str) >= 10 and dt_str[2] == '/' and dt_str[5] == '/':
                        return f"{dt_str[6:10]}-{dt_str[3:5]}"
                    if len(dt_str) >= 10 and dt_str[4] == '-':
                        return dt_str[:7]
                    return "0000-00"
                
                df_totvs['MesAno_TOTVS'] = df_totvs.get('Dt.Emiss', pd.Series(dtype=str)).apply(extract_mes_ano)
                
                # Chave agora é CNPJ + Nota + Ano-Mês
                df_totvs['Chave'] = df_totvs.apply(
                    lambda row: f"{row['CNPJ_TOTVS']}_{row['Nota_TOTVS']}_{row['MesAno_TOTVS']}" 
                    if row['CNPJ_TOTVS'] and row['Nota_TOTVS'] else None, axis=1
                )
                
                df_totvs = df_totvs.dropna(subset=['Chave']).copy()
                
                df_totvs['Valor_TOTVS'] = self._converter_moeda(df_totvs, 'Tot. Nota')
                df_totvs['BaseCalc_TOTVS'] = self._converter_moeda(df_totvs, 'Base Calc.')
                
                # Impostos Retidos
                df_totvs['ISS_TOTVS'] = self._converter_moeda(df_totvs, 'ISS')
                df_totvs['IRRF_TOTVS'] = self._converter_moeda(df_totvs, 'IRRF')
                df_totvs['INSS_TOTVS'] = self._converter_moeda(df_totvs, 'INSS')
                df_totvs['CSRF_TOTVS'] = self._converter_moeda(df_totvs, 'CSRF')
                
                # Alíquotas para validação da regra de Simples/CSRF
                df_totvs['Aliq_IRRF_TOTVS'] = self._converter_moeda(df_totvs, 'Aliq IRRF')
                df_totvs['Aliq_INSS_TOTVS'] = self._converter_moeda(df_totvs, 'Aliq INSS')
                df_totvs['Aliq_ISS_TOTVS'] = self._converter_moeda(df_totvs, 'Aliq ISS')
                df_totvs['Aliq_CSRF_TOTVS'] = self._converter_moeda(df_totvs, 'Aliq CSRF')
                
                print(f"[ETL TOTVS] A consolidar linhas fatiadas por Centro de Custo/Itens...")
                cols_numericas = [
                    'Valor_TOTVS', 'BaseCalc_TOTVS', 'ISS_TOTVS', 'IRRF_TOTVS', 'INSS_TOTVS', 'CSRF_TOTVS'
                ]
                
                regras_agrupamento = {}
                for coluna in df_totvs.columns:
                    if coluna in cols_numericas:
                        regras_agrupamento[coluna] = 'sum'
                    elif coluna != 'Chave':
                        regras_agrupamento[coluna] = 'first'
                        
                df_totvs = df_totvs.groupby('Chave', as_index=False).agg(regras_agrupamento)
                
                print(f"[ETL TOTVS] Lançamentos finais enviados para o cruzamento: {len(df_totvs)}")
                return df_totvs
            else:
                print("[ERRO TOTVS] Colunas 'Docto' ou 'CNPJ emit' não encontradas nos arquivos combinados.")
                return pd.DataFrame()
        except Exception as e:
            print(f"[ERRO TOTVS] Falha ao estruturar relatórios combinados: {e}")
            return pd.DataFrame()


class AuditorFiscal:
    def cruzar_dados(self, df_xml: pd.DataFrame, df_totvs: pd.DataFrame, df_simples: pd.DataFrame) -> pd.DataFrame:
        print("[AUDITORIA] A executar a Malha Fina Tributária Avançada...")
        if df_xml.empty or df_totvs.empty:
            return pd.DataFrame()
            
        df_cruzamento = pd.merge(df_xml, df_totvs, on='Chave', how='outer', indicator=True)
        
        df_cruzamento['CNPJ_Merge'] = df_cruzamento['CNPJ_TOTVS'].fillna(df_cruzamento['CNPJ_XML'])
        
        if not df_simples.empty:
            df_cruzamento = pd.merge(df_cruzamento, df_simples, left_on='CNPJ_Merge', right_on='CNPJ_Limpo', how='left')
        else:
            df_cruzamento['Simples_SN'] = 'PENDENTE_CONSULTA'
            
        for col, col_fallback in [('CNPJ emit', 'CNPJ_XML'), ('Docto', 'Nota_XML'), ('Nome Abrev', 'Nome_XML'), ('Tot. Nota', 'Valor_XML')]:
            if col in df_cruzamento.columns:
                df_cruzamento[col] = df_cruzamento[col].fillna(df_cruzamento[col_fallback])
        
        def formata_data_totvs(dt_str):
            if pd.isna(dt_str) or not dt_str: return ""
            dt_str = str(dt_str).strip()
            if len(dt_str) >= 10 and dt_str[2] == '/' and dt_str[5] == '/':
                return f"{dt_str[6:10]}-{dt_str[3:5]}-{dt_str[0:2]}"
            if len(dt_str) >= 10 and dt_str[4] == '-':
                return dt_str[:10]
            return dt_str

        def validar_impostos(row):
            if row['_merge'] == 'left_only': return "" 
            
            erros_encontrados = []
            
            dt_xml = str(row.get('DataEmissao_XML', '')).strip()[:10]
            dt_totvs = formata_data_totvs(row.get('Dt.Emiss', ''))
            
            val_simples_planilha = str(row.get('Simples_SN', 'PENDENTE_CONSULTA')).strip().upper()
            val_simples_xml = str(row.get('Simples_XML_Resolvido', 'NÃO')).strip().upper()
            is_simples = (val_simples_planilha == 'SIM') or (val_simples_xml == 'SIM')
            
            if val_simples_planilha == 'PENDENTE_CONSULTA' or val_simples_xml == 'PENDENTE_CONSULTA':
                erros_encontrados.append("SIMPLES_RET_INDEVIDA") 
                
            teve_erro_simples = False
            
            # --- REGRA 1: Optantes pelo Simples Nacional NÃO devem sofrer retenção de impostos ---
            if is_simples:
                retencoes_indevidas = []
                if row.get('IRRF_TOTVS', 0.0) > 0 and row.get('Aliq_IRRF_TOTVS', 0.0) > 0: retencoes_indevidas.append('IRRF')
                if row.get('INSS_TOTVS', 0.0) > 0 and row.get('Aliq_INSS_TOTVS', 0.0) > 0: retencoes_indevidas.append('INSS')
                if row.get('ISS_TOTVS', 0.0) > 0 and row.get('Aliq_ISS_TOTVS', 0.0) > 0: retencoes_indevidas.append('ISS')
                if row.get('CSRF_TOTVS', 0.0) > 0 and row.get('Aliq_CSRF_TOTVS', 0.0) > 0: retencoes_indevidas.append('CSRF')
                
                if retencoes_indevidas:
                    erros_encontrados.append("SIMPLES_RET_INDEVIDA")
                    teve_erro_simples = True
            
            # --- REGRA 2: Validação da Alíquota de CSRF (deve ser 4.65%) ---
            if row.get('CSRF_TOTVS', 0.0) > 0:
                aliq_csrf = row.get('Aliq_CSRF_TOTVS', 0.0)
                if not (4.6 <= aliq_csrf <= 4.7):
                    erros_encontrados.append("CSRF_ALIQ_ERRADA")

            limite_isencao = 10.0 

            # --- CASO A: NOTA TEM XML E EXISTE NO ERP ---
            if row['_merge'] == 'both':
                if dt_xml and dt_totvs and dt_xml != dt_totvs:
                    erros_encontrados.append("DATA_EMISSAO")

                if abs(row.get('Valor_XML', 0.0) - row.get('Valor_TOTVS', 0.0)) > 0.10:
                    erros_encontrados.append("VALOR_BRUTO")
                    
                flag_ret = str(row.get('Flag_Ret_ISS', '1'))
                iss_esperado = row.get('ISS_XML', 0.0) if flag_ret == '2' else 0.0
                if abs(iss_esperado - row.get('ISS_TOTVS', 0.0)) > limite_isencao:
                    erros_encontrados.append("ISS")
                    
                if not teve_erro_simples:
                    if abs(row.get('IRRF_XML', 0.0) - row.get('IRRF_TOTVS', 0.0)) > limite_isencao:
                        erros_encontrados.append('IR')
                    if abs(row.get('INSS_XML', 0.0) - row.get('INSS_TOTVS', 0.0)) > limite_isencao:
                        erros_encontrados.append('INSS')

            # --- CASO B: NOTA SÓ EXISTE NO ERP (SEM XML) ---
            elif row['_merge'] == 'right_only':
                tem_outro_erro = teve_erro_simples or ("CSRF_ALIQ_ERRADA" in erros_encontrados)
                
                if not teve_erro_simples:
                    erros_totvs = []
                    # Para notas cegas, validamos internamente o que o TOTVS diz que reteve
                    if row.get('IRRF_TOTVS', 0.0) > limite_isencao: erros_totvs.append("IR")
                    if row.get('INSS_TOTVS', 0.0) > limite_isencao: erros_totvs.append("INSS")
                    if row.get('ISS_TOTVS', 0.0) > limite_isencao: erros_totvs.append("ISS")
                    
                    if erros_totvs:
                        erros_encontrados.append("_".join(erros_totvs))
                        tem_outro_erro = True
                        
                # Nova regra: Retirar das inconsistências as notas que apenas faltam XML
                if tem_outro_erro:
                    erros_encontrados.append("FALTA_XML")

            return ",".join(erros_encontrados) if erros_encontrados else "OK"
            
        def status_curto(row):
            if row['_merge'] == 'left_only': return 'IGNORAR'
            elif row['Observações'] != 'OK': 
                if 'FALTA_XML' in row['Observações']: return 'FALTA XML'
                return 'DIVERGENTE'
            else: return 'CONCILIADO'
            
        def mapear_apontamento(obs):
            if obs in ['OK', '']: return "Sem Erro"
            if "SIMPLES_RET_INDEVIDA" in obs: return "Sem Consulta Simples nacional"
            if "CSRF_ALIQ_ERRADA" in obs: return "Alíquota CSRF Divergente"
            
            tem_ir, tem_inss, tem_iss = "IR" in obs, "INSS" in obs, "ISS" in obs
            
            if tem_ir and tem_iss and tem_inss: return "Verf. IR/ISS/INSS"
            if tem_ir and tem_inss: return "Verf. IR/INSS"
            if tem_ir and tem_iss: return "Verf. IR/ISS"
            if tem_iss and tem_inss: return "Verf. ISS/INSS"
            
            if tem_ir: return "Verf. IR"
            if tem_iss: return "Verf. ISS"
            if tem_inss: return "Verf. INSS"
            
            if "FALTA_XML" in obs: return "Nota sem XML na Receita"
            if "DATA_EMISSAO" in obs and not "VALOR_BRUTO" in obs: return "Divergência de Data"
            if "VALOR_BRUTO" in obs: return "Outras Divergências"
            
            return "Outras Divergências"
                
        df_cruzamento['Observações'] = df_cruzamento.apply(validar_impostos, axis=1)
        df_cruzamento['Valid. FINAL'] = df_cruzamento.apply(status_curto, axis=1)
        df_cruzamento['Apontamento_Mapeado'] = df_cruzamento['Observações'].apply(mapear_apontamento)
        
        return df_cruzamento

    def exportar_relatorio_divergencias(self, df_resultado: pd.DataFrame, caminho_saida: Path) -> Path:
        print(f"[EXPORTAÇÃO] A desenhar painel executivo e a pintar colunas de erro...")
        if df_resultado.empty: return None

        df_lancadas = df_resultado[df_resultado['_merge'].isin(['both', 'right_only'])].copy()
        if df_lancadas.empty: return None
            
        df_erros = df_lancadas[df_lancadas['Valid. FINAL'].isin(['DIVERGENTE', 'FALTA XML'])].copy()

        cols_financeiras = ['Valor_TOTVS', 'IRRF_TOTVS', 'CSRF_TOTVS', 'INSS_TOTVS', 'ISS_TOTVS']
        for c in cols_financeiras:
            if c in df_lancadas.columns:
                df_lancadas[c] = pd.to_numeric(df_lancadas[c], errors='coerce').fillna(0)

        total_nfs = len(df_lancadas)
        total_verificar = len(df_erros)
        
        est_col = 'Est' if 'Est' in df_lancadas.columns else 'Est Fiscal'
        if est_col in df_lancadas.columns:
            df_estab = df_lancadas.groupby(est_col).size().reset_index(name='Total NFS')
            erros_estab = df_erros.groupby(est_col).size().reset_index(name='Para Verificar')
            df_estab = pd.merge(df_estab, erros_estab, on=est_col, how='left').fillna(0)
            df_estab['Para Verificar'] = df_estab['Para Verificar'].astype(int)
        else:
            df_estab = pd.DataFrame(columns=['Estab', 'Total NFS', 'Para Verificar'])
        
        lista_apontamentos = [
            (1, "Sem Consulta Simples nacional"),
            (2, "Verf. IR/ISS/INSS"),
            (3, "Verf. IR/INSS"),
            (4, "Verf. IR/ISS"),
            (5, "Verf. ISS/INSS"),
            (6, "Verf. IR"),
            (7, "Verf. ISS"),
            (8, "Verf. INSS"),
            (9, "Alíquota CSRF Divergente"),
            (10, "Nota sem XML na Receita"),
            (11, "Divergência de Data")
        ]
        erros_contagem = df_erros['Apontamento_Mapeado'].value_counts().to_dict()
        dados_apontamentos = []
        for ordem, desc in lista_apontamentos:
            qtd = erros_contagem.get(desc, 0)
            dados_apontamentos.append((ordem, desc, qtd))
            
        if 'CNPJ emit' in df_lancadas.columns and 'Nome Abrev' in df_lancadas.columns:
            df_forn = df_lancadas.groupby(['CNPJ emit', 'Nome Abrev']).size().reset_index(name='Quantidade')
            df_forn = df_forn.sort_values(by='Quantidade', ascending=False).head(10)
        else:
            df_forn = pd.DataFrame(columns=['CNPJ emit', 'Nome Abrev', 'Quantidade'])
        
        if 'Cód.Item' in df_lancadas.columns and 'Descrição Item' in df_lancadas.columns:
            df_serv = df_lancadas.groupby(['Cód.Item', 'Descrição Item']).size().reset_index(name='Quantidades')
            df_serv = df_serv.sort_values(by='Quantidades', ascending=False).head(5)
        else:
            df_serv = pd.DataFrame(columns=['Cód.Item', 'Desc. Tipo de serviço', 'Quantidades'])
            
        totais_gerais = [
            ("Valor Total NFS", df_lancadas['Valor_TOTVS'].sum()),
            ("Total IRRF", df_lancadas['IRRF_TOTVS'].sum()),
            ("Total CSRF", df_lancadas['CSRF_TOTVS'].sum()),
            ("TOTAL INSS", df_lancadas['INSS_TOTVS'].sum()),
            ("TOTAL I.S.S", df_lancadas['ISS_TOTVS'].sum())
        ]

        arquivo_saida = caminho_saida / "relatorio_servicos.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Dashboard"
        ws.sheet_view.showGridLines = False

        fill_title = PatternFill(fill_type="solid", fgColor="0E2841")
        fill_header = PatternFill(fill_type="solid", fgColor="FFC000")
        font_title = Font(bold=True, color="FFFFFF")
        font_header = Font(bold=True, color="000000")
        borda = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        def criar_titulo_secao(ws, col_start, col_end, row, texto):
            ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
            c = ws.cell(row=row, column=col_start, value=texto)
            c.font = font_title
            c.fill = fill_title
            c.alignment = Alignment(horizontal="center", vertical="center")
            for col in range(col_start, col_end + 1):
                ws.cell(row=row, column=col).border = borda

        def criar_tabela(ws, col_start, row_start, headers, dados):
            for i, h in enumerate(headers):
                c = ws.cell(row=row_start, column=col_start+i, value=h)
                c.font = font_header
                c.fill = fill_header
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = borda
            for r_idx, row_data in enumerate(dados, 1):
                for c_idx, val in enumerate(row_data):
                    c = ws.cell(row=row_start+r_idx, column=col_start+c_idx, value=val)
                    c.border = borda
                    if isinstance(val, (int, float)): 
                        c.alignment = Alignment(horizontal="center", vertical="center")

        larguras = {'B': 12, 'C': 15, 'D': 18, 'E': 3, 'F': 8, 'G': 35, 'H': 15, 'I': 3, 'J': 18, 'K': 35, 'L': 15, 'M': 3, 'N': 20, 'O': 20}
        for col, width in larguras.items():
            ws.column_dimensions[col].width = width

        criar_titulo_secao(ws, 2, 4, 2, "RESUMO DE NFS")
        ws.cell(row=3, column=2, value="Total de NFS").font = Font(bold=True)
        ws.cell(row=3, column=3, value=total_nfs).alignment = Alignment(horizontal="left")
        ws.cell(row=4, column=2, value="Total NFS Para Verificar").font = Font(bold=True)
        ws.cell(row=4, column=3, value=total_verificar).alignment = Alignment(horizontal="left")
        
        criar_tabela(ws, 2, 6, ["Estab", "Total NFS", "Para Verificar"], df_estab.values.tolist())
        
        criar_titulo_secao(ws, 6, 8, 2, "APONTAMENTOS")
        criar_tabela(ws, 6, 3, ["Ordem", "Descrição", "Quantidades"], dados_apontamentos)
        
        criar_titulo_secao(ws, 10, 12, 2, "10 MAIORES FORNECEDORES DE EMISSÃO")
        criar_tabela(ws, 10, 3, ["Emit.", "Nome Abrev", "Quantidade"], df_forn.values.tolist())
        
        row_serv = 3 + len(df_forn) + 2
        criar_titulo_secao(ws, 10, 12, row_serv, "5 MAIORES TIPO DE SERVIÇOS")
        criar_tabela(ws, 10, row_serv+1, ["Cód.Item", "Desc. Tipo de serviço", "Quantidades"], df_serv.values.tolist())

        criar_titulo_secao(ws, 14, 15, 2, "TOTAL GERAL")
        criar_tabela(ws, 14, 3, ["Descrição", "Valor"], totais_gerais)
        
        for row in range(4, 4 + len(totais_gerais)):
            c_val = ws.cell(row=row, column=15)
            c_val.number_format = 'R$ #,##0.00'

        ws_val = wb.create_sheet("Validação")
        
        colunas_comparacao = [
            'CNPJ_TOTVS', 'Nota_TOTVS',
            'DataEmissao_XML', 'Dt.Emiss',
            'Valor_XML', 'Valor_TOTVS',
            'ISS_XML', 'ISS_TOTVS',
            'IRRF_XML', 'IRRF_TOTVS',
            'INSS_XML', 'INSS_TOTVS',
            'CSRF_TOTVS', 'Aliq_CSRF_TOTVS'
        ]

        colunas_tecnicas = [
            'Chave', 'CNPJ_XML', 'Nota_XML', 'NSU_XML', 'Nome_XML', 'Origem_XML', '_merge', 
            'Apontamento_Mapeado', 'CNPJ_Merge', 'CNPJ_Limpo', 'Simples_SN', 'Flag_Ret_ISS', 
            'Simples_XML_Resolvido', 'BaseCalculo_XML', 'BaseCalc_TOTVS',
            'Aliq_IRRF_TOTVS', 'Aliq_INSS_TOTVS', 'Aliq_ISS_TOTVS'
        ]
        
        colunas_originais = [col for col in df_resultado.columns if col not in colunas_tecnicas and col not in colunas_comparacao and col not in ['Observações', 'Valid. FINAL']]
        colunas_finais = ['Observações', 'Valid. FINAL'] + colunas_comparacao + colunas_originais
        
        ws_val.append(colunas_finais)
        for col_idx, col_name in enumerate(colunas_finais, 1):
            celula = ws_val.cell(row=1, column=col_idx)
            celula.font = font_title
            celula.fill = fill_title
            ws_val.column_dimensions[celula.column_letter].width = 20
            
        if not df_erros.empty:
            df_erros['Ordem_Sort'] = df_erros['Valid. FINAL'].apply(lambda x: 1 if x == 'FALTA XML' else 0)
            df_erros = df_erros.sort_values(by=['Ordem_Sort', 'Apontamento_Mapeado']).drop(columns=['Ordem_Sort'])
            
            df_export = df_erros[colunas_finais].copy()
            df_export = df_export.fillna("")
            
            col_idx_map = {col_name: idx for idx, col_name in enumerate(colunas_finais, 1)}
            fill_erro = PatternFill(fill_type="solid", fgColor="FFC7CE") 
            font_erro = Font(color="9C0006", bold=True)                  
            
            for r_idx, r in enumerate(df_export.itertuples(index=False), start=2):
                ws_val.append(list(r))
                
                obs = str(r[0])
                
                def pintar_celulas(col_xml, col_totvs):
                    if col_xml in col_idx_map:
                        ws_val.cell(row=r_idx, column=col_idx_map[col_xml]).fill = fill_erro
                        ws_val.cell(row=r_idx, column=col_idx_map[col_xml]).font = font_erro
                    if col_totvs in col_idx_map:
                        ws_val.cell(row=r_idx, column=col_idx_map[col_totvs]).fill = fill_erro
                        ws_val.cell(row=r_idx, column=col_idx_map[col_totvs]).font = font_erro

                if "DATA_EMISSAO" in obs: pintar_celulas("DataEmissao_XML", "Dt.Emiss")
                if "VALOR_BRUTO" in obs: pintar_celulas("Valor_XML", "Valor_TOTVS")
                if "ISS" in obs: pintar_celulas("ISS_XML", "ISS_TOTVS")
                if "IR" in obs: pintar_celulas("IRRF_XML", "IRRF_TOTVS")
                if "INSS" in obs: pintar_celulas("INSS_XML", "INSS_TOTVS")
                
                if "CSRF_ALIQ_ERRADA" in obs:
                    pintar_celulas("CSRF_TOTVS", "Aliq_CSRF_TOTVS")
                
                if "SIMPLES_RET_INDEVIDA" in obs:
                    for fed_col in ["IRRF_TOTVS", "INSS_TOTVS", "CSRF_TOTVS", "ISS_TOTVS"]:
                        if fed_col in col_idx_map:
                            try:
                                val = r[col_idx_map[fed_col] - 1]
                                if val and float(val) > 0:
                                    ws_val.cell(row=r_idx, column=col_idx_map[fed_col]).fill = fill_erro
                                    ws_val.cell(row=r_idx, column=col_idx_map[fed_col]).font = font_erro
                            except ValueError:
                                pass
                
        ws_val.freeze_panes = "A2"
        ultima_coluna = get_column_letter(ws_val.max_column)
        ws_val.auto_filter.ref = f"A1:{ultima_coluna}{ws_val.max_row}"

        wb.save(arquivo_saida)
        print(f"[SUCESSO] Dashboard Pixel Perfect gerado em: {arquivo_saida}")
        return arquivo_saida

if __name__ == "__main__":
    RAIZ_PROJETO = Path(__file__).resolve().parent.parent
    config_caminhos = CaminhosProjeto(diretorio_base=RAIZ_PROJETO)
    
    gerador_csv = GeradorCsvGoverno(caminhos=config_caminhos)
    caminho_base_governo = gerador_csv.gerar_csv()
    
    tabela_xml = LeitorBaseGoverno(caminho_csv=caminho_base_governo).carregar()
    tabela_totvs = LeitorRelatorioTotvs(caminhos=config_caminhos).processar_relatorio()
    
    cnpjs_para_consultar = set()
    if not tabela_xml.empty and 'CNPJ_XML' in tabela_xml.columns:
        cnpjs_para_consultar.update(tabela_xml['CNPJ_XML'].dropna().tolist())
    if not tabela_totvs.empty and 'CNPJ_TOTVS' in tabela_totvs.columns:
        cnpjs_para_consultar.update(tabela_totvs['CNPJ_TOTVS'].dropna().tolist())
        
    tabela_simples = LeitorSimplesNacional(diretorio_base=RAIZ_PROJETO).carregar(cnpjs_alvo=cnpjs_para_consultar)
    
    auditor = AuditorFiscal()
    resultado = auditor.cruzar_dados(df_xml=tabela_xml, df_totvs=tabela_totvs, df_simples=tabela_simples)
    
    if not resultado.empty:
        print("\n" + "="*50)
        print("📊 RESUMO DE CRUZAMENTO DE DADOS (MATCH)")
        print("="*50)
        
        qtd_both = len(resultado[resultado['_merge'] == 'both'])
        qtd_so_totvs = len(resultado[resultado['_merge'] == 'right_only'])
        
        print(f"-> Notas processadas do TOTVS: {len(tabela_totvs)}")
        print(f"-> Notas processadas do Governo: {len(tabela_xml)}")
        print(f"-> Notas conciliadas perfeitamente: {qtd_both}")
        print(f"-> Notas no TOTVS SEM XML no Governo: {qtd_so_totvs}")
        
        config_caminhos.saida_auditoria.mkdir(parents=True, exist_ok=True)
        
        arquivo_excel = auditor.exportar_relatorio_divergencias(df_resultado=resultado, caminho_saida=config_caminhos.saida_auditoria)
        
        print("\n[!] Painéis de Serviços Gerados e Guardados na Pasta de Auditoria.")