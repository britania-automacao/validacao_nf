# src/services_etl.py
import os
import re
import time
import requests
import urllib3
import numpy as np
import pandas as pd
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

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
load_dotenv()


class LeitorRegrasRetencao:
    def __init__(self, caminhos: CaminhosProjeto):
        self.caminhos = caminhos
        p_rede = Path(r"P:\FISCAL\GERAL\05 - KPIs RECEBIMENTO FISCAL\8. Automacao Validacao NF\planilha_de_regras")
        p_local = caminhos.diretorio_base / "planilha_de_regras"
        try:
            if p_rede.exists():
                self.pasta = p_rede
            else:
                self.pasta = p_local
        except (PermissionError, OSError):
            self.pasta = p_local
            
        self.arquivo = self.pasta / "regras_retencao.xlsx"
        if not self.arquivo.exists() and (p_local / "regras_retencao.xlsx").exists():
            self.arquivo = p_local / "regras_retencao.xlsx"
        
    def carregar(self) -> dict:
        if not self.arquivo.exists():
            print(f"[AVISO] Planilha de regras não encontrada no caminho: {self.arquivo}. A validação por Cód. Item será ignorada.")
            return {}
            
        try:
            print(f"[ETL REGRAS] Lendo regras de retenção do arquivo na rede...")
            df_regras = pd.read_excel(self.arquivo)
            
            df_regras = df_regras.dropna(subset=['COD ITEM'])
            df_regras['COD ITEM'] = df_regras['COD ITEM'].astype(str).str.replace('.0', '', regex=False).str.strip()
            
            regras = {}
            for _, row in df_regras.iterrows():
                cod = row['COD ITEM']
                regras[cod] = {
                    'IR': str(row.get('Ret. IR', '')).strip().upper(),
                    'CS': str(row.get('Ret. CS', '')).strip().upper(),
                    'ISS': str(row.get('Ret. I.S.S', '')).strip().upper(),
                    'INSS': str(row.get('Ret. INSS', '')).strip().upper()
                }
            print(f"[ETL REGRAS] Sucesso! Foram mapeadas regras para {len(regras)} Códigos de Item.")
            return regras
        except Exception as e:
            print(f"[ERRO REGRAS] Falha ao ler a planilha de regras de retenção na rede: {e}")
            return {}


class LeitorSimplesNacional:
    def __init__(self, caminhos: CaminhosProjeto):
        self.caminhos = caminhos
        try:
            self.pasta = caminhos.consulta_simples_nacional
            if not self.pasta.exists():
                self.pasta.mkdir(parents=True, exist_ok=True)
        except (PermissionError, OSError):
            self.pasta = caminhos.diretorio_base / "dados" / "consulta_simples_nacional"
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
        if not cnpj or len(cnpj) != 14: 
            return "PENDENTE_CONSULTA"
            
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'application/json'
        }

        # Cascata de APIs otimizada: prioriza as mais rápidas com maior cota gratuita
        apis = [
            {"nome": "CNPJ.ws", "url": f"https://publica.cnpj.ws/cnpj/{cnpj}"},
            {"nome": "BrasilAPI", "url": f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}"},
            {"nome": "MinhaReceita", "url": f"https://minhareceita.org/{cnpj}"},
            {"nome": "ReceitaWS", "url": f"https://receitaws.com.br/v1/cnpj/{cnpj}"}
        ]

        for api in apis:
            nome_api = api["nome"]
            url = api["url"]
            try:
                try:
                    response = requests.get(url, headers=headers, timeout=8, verify=True)
                except requests.exceptions.SSLError:
                    response = requests.get(url, headers=headers, timeout=8, verify=False)
                if response.status_code == 200:
                    dados = response.json()
                    
                    if nome_api == "CNPJ.ws":
                        simples_info = dados.get("simples")
                        if isinstance(simples_info, dict):
                            optante = simples_info.get("simples")
                            if isinstance(optante, str):
                                return "SIM" if optante.strip().lower() == "sim" else "NÃO"
                            elif isinstance(optante, bool):
                                return "SIM" if optante else "NÃO"
                        return "NÃO"
                        
                    elif nome_api in ["BrasilAPI", "MinhaReceita"]:
                        opcao = dados.get("opcao_pelo_simples")
                        if opcao is not None:
                            return "SIM" if opcao is True else "NÃO"
                            
                    elif nome_api == "ReceitaWS":
                        if dados.get("status") == "ERROR": continue
                        simples = dados.get("simples")
                        if isinstance(simples, dict):
                            opcao = simples.get("optante")
                            return "SIM" if opcao is True else "NÃO"
                        return "NÃO"
                        
                elif response.status_code == 429:
                    # Se tomar rate limit em uma API, pula imediatamente para a próxima
                    continue
            except Exception:
                continue
        return "PENDENTE_CONSULTA"

    def carregar(self, cnpjs_alvo=None, mapa_emitente=None, mapa_simples_xml: dict | None = None) -> pd.DataFrame:
        arquivos = []
        try:
            arquivos = list(self.pasta.glob("*.xlsx")) + list(self.pasta.glob("*.csv"))
        except Exception:
            pass
            
        # Fallback para pasta local se a pasta configurada estiver inacessível ou vazia
        if not arquivos:
            pasta_local = self.caminhos.diretorio_base / "dados" / "consulta_simples_nacional"
            if pasta_local.exists():
                try:
                    arquivos = list(pasta_local.glob("*.xlsx")) + list(pasta_local.glob("*.csv"))
                    if arquivos:
                        self.pasta = pasta_local
                except Exception:
                    pass

        df_original = pd.DataFrame()
        col_cnpj, col_simples = "CNPJ", "SIMPLES"
        caminho_salvar = None

        if arquivos:
            caminho_salvar = arquivos[0]
            try:
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
                
                if 'Emit.' not in df_original.columns: df_original['Emit.'] = ""
            except Exception as e:
                df_original = pd.DataFrame(columns=['CNPJ_Limpo', 'Simples_SN', 'Emit.'])
        else:
            caminho_salvar = self.pasta / "consulta_simples_nacional.xlsx"
            df_original = pd.DataFrame(columns=['CNPJ', 'Emit.', 'SIMPLES', 'CNPJ_Limpo', 'Simples_SN'])

        houve_alteracao = False

        # --- REVALIDAÇÃO INTELIGENTE DE CNPJs COM DIVERGÊNCIA (PLANILHA x XML OFICIAL) ---
        if mapa_simples_xml and not df_original.empty and 'CNPJ_Limpo' in df_original.columns:
            cnpjs_para_revalidar = []
            for idx, row in df_original.iterrows():
                cnpj_c = str(row.get('CNPJ_Limpo', '')).strip()
                if cnpj_c in mapa_simples_xml:
                    st_planilha = str(row.get('Simples_SN', '')).strip().upper()
                    st_xml = str(mapa_simples_xml[cnpj_c]).strip().upper()
                    if st_planilha in ['SIM', 'NÃO'] and st_xml in ['SIM', 'NÃO'] and st_planilha != st_xml:
                        cnpjs_para_revalidar.append((idx, cnpj_c, st_planilha, st_xml))
            
            if cnpjs_para_revalidar:
                print(f"[API SIMPLES] {len(cnpjs_para_revalidar)} CNPJ(s) com divergência entre a Planilha e o XML detectados! Revalidando na Receita Federal...")
                for idx, cnpj_c, st_planilha, st_xml in cnpjs_para_revalidar:
                    print(f"[API SIMPLES] Reconsultando CNPJ {cnpj_c} (Planilha={st_planilha} vs XML={st_xml})...")
                    status_api = self._consultar_api(cnpj_c)
                    if status_api in ['SIM', 'NÃO']:
                        print(f"-> Retorno da Receita Federal para {cnpj_c}: {status_api}")
                        if status_api != st_planilha:
                            print(f"-> Atualizando cadastro do CNPJ {cnpj_c} na planilha de '{st_planilha}' para '{status_api}'!")
                            df_original.at[idx, col_simples] = status_api
                            df_original.at[idx, 'Simples_SN'] = status_api
                            houve_alteracao = True
                    else:
                        # Se a API estiver temporariamente indisponível mas o XML emitido declara NÃO,
                        # respeita a declaração do documento fiscal emitido
                        print(f"-> Não foi possível obter resposta definitiva da API para {cnpj_c} ({status_api}).")
                    time.sleep(2)

        if cnpjs_alvo:
            cnpjs_existentes = set()
            if not df_original.empty and 'CNPJ_Limpo' in df_original.columns:
                df_original = df_original[df_original['Simples_SN'] != 'PENDENTE_CONSULTA'].copy()
                cnpjs_existentes = set(df_original['CNPJ_Limpo'].dropna().tolist())

            cnpjs_faltantes = {c for c in cnpjs_alvo if isinstance(c, str) and len(c) == 14} - cnpjs_existentes

            if cnpjs_faltantes:
                houve_alteracao = True
                print(f"[API SIMPLES] {len(cnpjs_faltantes)} CNPJs pendentes/novos detectados! Processando com cascata otimizada...")
                novos_registros = []
                
                for idx, cnpj in enumerate(cnpjs_faltantes, 1):
                    # 1. Otimização rápida: se o XML oficial recente da NFS-e já traz o status, aproveita de imediato (sem gastar cota externa)
                    if mapa_simples_xml and cnpj in mapa_simples_xml and mapa_simples_xml[cnpj] in ['SIM', 'NÃO']:
                        status = mapa_simples_xml[cnpj]
                        print(f"[API SIMPLES] ({idx}/{len(cnpjs_faltantes)}) CNPJ {cnpj}: Obtido diretamente do XML oficial -> {status}")
                    else:
                        # 2. Caso contrário, consulta a cascata de APIs com intervalo suave de 2s
                        status = self._consultar_api(cnpj)
                        print(f"[API SIMPLES] ({idx}/{len(cnpjs_faltantes)}) CNPJ {cnpj}: Consulta externa concluída -> {status}")
                        time.sleep(2)
                        
                    emit_novo = str(mapa_emitente.get(cnpj, "")) if mapa_emitente else ""
                    nova_linha = {
                        col_cnpj: cnpj, 'Emit.': emit_novo, col_simples: status,
                        'CNPJ_Limpo': cnpj, 'Simples_SN': status
                    }
                    novos_registros.append(nova_linha)

                df_novos = pd.DataFrame(novos_registros)
                df_original = pd.concat([df_original, df_novos], ignore_index=True)
                
        if houve_alteracao:
            try:
                df_salvar = df_original.drop(columns=['CNPJ_Limpo', 'Simples_SN'], errors='ignore')
                df_salvar[col_cnpj] = df_salvar[col_cnpj].apply(lambda x: str(x).zfill(14) if pd.notna(x) else x)
                if caminho_salvar.suffix.lower() == '.xlsx':
                    df_salvar.to_excel(caminho_salvar, index=False)
                else:
                    df_salvar.to_csv(caminho_salvar, sep=';', index=False, encoding='utf-8-sig')
                print(f"[API SIMPLES] Planilha de Simples Nacional atualizada com sucesso: {caminho_salvar.name}")
            except Exception as e_salvar:
                print(f"[AVISO] Não foi possível salvar atualizações no arquivo principal: {e_salvar}")
                # Fallback: tentar salvar na pasta local se a remota falhar
                try:
                    caminho_local = self.caminhos.diretorio_base / "dados" / "consulta_simples_nacional" / "consulta_simples_nacional.xlsx"
                    caminho_local.parent.mkdir(parents=True, exist_ok=True)
                    df_salvar.to_excel(caminho_local, index=False)
                    print(f"[API SIMPLES] Atualizações salvas com sucesso no backup local: {caminho_local}")
                except Exception:
                    pass

        if not df_original.empty and 'CNPJ_Limpo' in df_original.columns:
            df_retorno = df_original.drop_duplicates(subset=['CNPJ_Limpo'], keep='last')
            return df_retorno[['CNPJ_Limpo', 'Simples_SN']]
        else:
            return pd.DataFrame(columns=['CNPJ_Limpo', 'Simples_SN'])


class LeitorResumosNacionais:
    def __init__(self, caminhos: CaminhosProjeto):
        self.pasta_raiz = caminhos.xml_servicos_nacional

    def _obter_coluna(self, df: pd.DataFrame, nomes_possiveis: list, default=""):
        for nome in nomes_possiveis:
            matches = [c for c in df.columns if str(c).endswith(nome)]
            if matches: return df[matches[0]].fillna(default)
        return pd.Series(default, index=df.index)

    def _obter_coluna_float(self, df: pd.DataFrame, nomes_possiveis: list):
        col_str = self._obter_coluna(df, nomes_possiveis, default="0")
        return pd.to_numeric(col_str, errors='coerce').fillna(0.0)

    def carregar(self) -> pd.DataFrame:
        print("[ETL XML] Buscando planilhas Resumo_Notas_Nacional nas subpastas das empresas...")
        arquivos = list(self.pasta_raiz.rglob("Resumo_Notas_Nacional_*.xlsx"))
        if not arquivos: return pd.DataFrame()

        lista_dfs = []
        for caminho in arquivos:
            try:
                df_temp = pd.read_excel(caminho, dtype=str)
                df_temp['Origem_Arquivo'] = caminho.name
                lista_dfs.append(df_temp)
            except Exception: pass

        if not lista_dfs: return pd.DataFrame()

        df_bruto = pd.concat(lista_dfs, ignore_index=True)
        df_out = pd.DataFrame()
        df_out['NSU_XML'] = self._obter_coluna(df_bruto, ['NSU_API'])
        df_out['Origem_XML'] = self._obter_coluna(df_bruto, ['Origem_Arquivo'])
        df_out['Chave_Acesso_NFSe'] = self._obter_coluna(df_bruto, ['Chave_Acesso_NFSe'])
        
        cnpj_raw = self._obter_coluna(df_bruto, ['emit_CNPJ', 'PrestadorServico_CNPJ', 'prest_CNPJ'])
        df_out['CNPJ_XML'] = cnpj_raw.astype(str).str.replace('.0', '', regex=False).str.extract(r'(\d+)')[0].str.zfill(14)
        
        nota_raw = self._obter_coluna(df_bruto, ['nNFSe', 'Numero', 'nNFS-e'])
        df_out['Nota_XML'] = nota_raw.astype(str).str.replace('.0', '', regex=False).str.extract(r'(\d+)')[0]
        
        dh_emi = self._obter_coluna(df_bruto, ['dhEmi', 'DataEmissao', 'Competencia'])
        df_out['DataEmissao_XML'] = dh_emi.astype(str).str[:10]
        df_out['MesAno_XML'] = df_out['DataEmissao_XML'].apply(lambda x: str(x)[:7] if len(str(x)) >= 7 else "0000-00")
        df_out['Chave'] = df_out['CNPJ_XML'] + "_" + df_out['Nota_XML'] + "_" + df_out['MesAno_XML']
        
        df_out['Nome_XML'] = self._obter_coluna(df_bruto, ['emit_xNome', 'PrestadorServico_RazaoSocial', 'xNome'], "Fornecedor Desconhecido")
        
        v_serv = self._obter_coluna_float(df_bruto, ['vServ', 'ValorServicos'])
        v_bc = self._obter_coluna_float(df_bruto, ['vBC', 'BaseCalculo'])
        
        df_out['Valor_XML'] = np.where(v_serv == 0, v_bc, v_serv)
        df_out['BaseCalculo_XML'] = v_bc
        df_out['ISS_XML'] = self._obter_coluna_float(df_bruto, ['vISSQN', 'ValorIss'])
        df_out['IRRF_XML'] = self._obter_coluna_float(df_bruto, ['vRetIRRF', 'ValorIr'])
        df_out['INSS_XML'] = self._obter_coluna_float(df_bruto, ['vRetCP', 'ValorInss'])
        df_out['PIS_XML'] = self._obter_coluna_float(df_bruto, ['vRetPIS', 'ValorPis'])
        df_out['COFINS_XML'] = self._obter_coluna_float(df_bruto, ['vRetCOFINS', 'ValorCofins'])
        df_out['CSLL_XML'] = self._obter_coluna_float(df_bruto, ['vRetCSLL', 'ValorCsll'])
        df_out['Flag_Ret_ISS'] = self._obter_coluna(df_bruto, ['tpRetISSQN', 'IssRetido'], '1')
        
        op_simp_nac = self._obter_coluna(df_bruto, ['opSimpNac'])
        optante_abrasf = self._obter_coluna(df_bruto, ['OptanteSimplesNacional'])
        
        def resolver_simples(row):
            val_nac = str(row['op_simp_nac'])
            val_abr = str(row['optante_abrasf'])
            if val_nac in ['2', '3', '4']: return "SIM"
            if val_nac == '1': return "NÃO"
            if val_abr == '1': return "SIM"
            if val_abr == '2': return "NÃO"
            return "NÃO"
            
        df_temp_simples = pd.DataFrame({'op_simp_nac': op_simp_nac, 'optante_abrasf': optante_abrasf})
        df_out['Simples_XML_Resolvido'] = df_temp_simples.apply(resolver_simples, axis=1)
        
        df_out = df_out.dropna(subset=['CNPJ_XML', 'Nota_XML'])
        print(f"[ETL XML] Carregados {len(df_out)} registros cruzados diretamente das planilhas do Governo!")
        return df_out


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
        if nome_coluna not in df.columns: return 0.0
        def limpar_moeda(valor):
            if pd.isna(valor): return 0.0
            v_str = str(valor).strip()
            if ',' in v_str: v_str = v_str.replace('.', '').replace(',', '.')
            try: return float(v_str)
            except ValueError: return 0.0
        return df[nome_coluna].apply(limpar_moeda)

    def processar_relatorio(self) -> pd.DataFrame:
        print(f"[ETL TOTVS] A iniciar a leitura do ERP na pasta: {self.pasta.name}...")
        arquivos_totvs = list(self.pasta.rglob("*.xlsx")) + list(self.pasta.rglob("*.csv"))
        if not arquivos_totvs: return pd.DataFrame()

        lista_dfs = []
        for caminho_relatorio in arquivos_totvs:
            try:
                if caminho_relatorio.suffix.lower() == '.xlsx':
                    df_temp = pd.read_excel(caminho_relatorio, dtype=str)
                else:
                    try:
                        df_temp = pd.read_csv(caminho_relatorio, sep=';', encoding='utf-8', dtype=str)
                    except UnicodeDecodeError:
                        df_temp = pd.read_csv(caminho_relatorio, sep=';', encoding='latin1', dtype=str)
                lista_dfs.append(df_temp)
            except Exception: pass

        if not lista_dfs: return pd.DataFrame()
        df_totvs = pd.concat(lista_dfs, ignore_index=True)

        try:
            df_totvs.dropna(how='all', inplace=True)
            df_totvs.columns = df_totvs.columns.str.strip().str.replace('\n', ' ').str.replace('\r', '')

            if 'Docto' in df_totvs.columns and 'CNPJ emit' in df_totvs.columns:
                df_totvs['Nota_TOTVS'] = df_totvs['Docto'].apply(self._limpar_numero_nota)
                df_totvs['CNPJ_TOTVS'] = df_totvs['CNPJ emit'].apply(self._limpar_cnpj)
                
                def extract_mes_ano(dt_str):
                    if pd.isna(dt_str) or not dt_str: return "0000-00"
                    dt_str = str(dt_str).strip()
                    if len(dt_str) >= 10 and dt_str[2] == '/' and dt_str[5] == '/':
                        return f"{dt_str[6:10]}-{dt_str[3:5]}"
                    if len(dt_str) >= 10 and dt_str[4] == '-':
                        return dt_str[:7]
                    return "0000-00"
                
                df_totvs['MesAno_TOTVS'] = df_totvs.get('Dt.Emiss', pd.Series(dtype=str)).apply(extract_mes_ano)
                df_totvs['Chave'] = df_totvs.apply(
                    lambda row: f"{row['CNPJ_TOTVS']}_{row['Nota_TOTVS']}_{row['MesAno_TOTVS']}" 
                    if row['CNPJ_TOTVS'] and row['Nota_TOTVS'] else None, axis=1
                )
                df_totvs = df_totvs.dropna(subset=['Chave']).copy()
                
                df_totvs['Valor_TOTVS'] = self._converter_moeda(df_totvs, 'Tot. Nota')
                df_totvs['BaseCalc_TOTVS'] = self._converter_moeda(df_totvs, 'Base Calc.')
                df_totvs['ISS_TOTVS'] = self._converter_moeda(df_totvs, 'ISS')
                df_totvs['IRRF_TOTVS'] = self._converter_moeda(df_totvs, 'IRRF')
                df_totvs['INSS_TOTVS'] = self._converter_moeda(df_totvs, 'INSS')
                df_totvs['CSRF_TOTVS'] = self._converter_moeda(df_totvs, 'CSRF')
                df_totvs['Aliq_IRRF_TOTVS'] = self._converter_moeda(df_totvs, 'Aliq IRRF')
                df_totvs['Aliq_INSS_TOTVS'] = self._converter_moeda(df_totvs, 'Aliq INSS')
                df_totvs['Aliq_ISS_TOTVS'] = self._converter_moeda(df_totvs, 'Aliq ISS')
                df_totvs['Aliq_CSRF_TOTVS'] = self._converter_moeda(df_totvs, 'Aliq CSRF')
                
                cols_numericas = ['Valor_TOTVS', 'BaseCalc_TOTVS', 'ISS_TOTVS', 'IRRF_TOTVS', 'INSS_TOTVS', 'CSRF_TOTVS']
                regras_agrupamento = {}
                for coluna in df_totvs.columns:
                    if coluna in cols_numericas: regras_agrupamento[coluna] = 'sum'
                    elif coluna != 'Chave': regras_agrupamento[coluna] = 'first'
                        
                df_totvs = df_totvs.groupby('Chave', as_index=False).agg(regras_agrupamento)
                return df_totvs
            else:
                return pd.DataFrame()
        except Exception:
            return pd.DataFrame()


class AuditorFiscal:
    def cruzar_dados(self, df_xml: pd.DataFrame, df_totvs: pd.DataFrame, df_simples: pd.DataFrame, regras_retencao: dict) -> pd.DataFrame:
        print("[AUDITORIA] A executar a Malha Fina Tributária Avançada (com regras por COD ITEM)...")
        if df_xml.empty or df_totvs.empty: return pd.DataFrame()
            
        df_cruzamento = pd.merge(df_xml, df_totvs, on='Chave', how='outer', indicator=True)
        df_cruzamento['CNPJ_Merge'] = df_cruzamento['CNPJ_TOTVS'].fillna(df_cruzamento['CNPJ_XML'])
        
        if not df_simples.empty:
            df_cruzamento = pd.merge(df_cruzamento, df_simples, left_on='CNPJ_Merge', right_on='CNPJ_Limpo', how='left')
        else:
            df_cruzamento['Simples_SN'] = 'PENDENTE_CONSULTA'
            
        for col, col_fallback in [('CNPJ emit', 'CNPJ_XML'), ('Docto', 'Nota_XML'), ('Nome Abrev', 'Nome_XML'), ('Tot. Nota', 'Valor_XML')]:
            if col in df_cruzamento.columns: df_cruzamento[col] = df_cruzamento[col].fillna(df_cruzamento[col_fallback])
        
        def formata_data_totvs(dt_str):
            if pd.isna(dt_str) or not dt_str: return ""
            dt_str = str(dt_str).strip()
            if len(dt_str) >= 10 and dt_str[2] == '/' and dt_str[5] == '/': return f"{dt_str[6:10]}-{dt_str[3:5]}-{dt_str[0:2]}"
            if len(dt_str) >= 10 and dt_str[4] == '-': return dt_str[:10]
            return dt_str

        def validar_impostos(row):
            if row['_merge'] == 'left_only': return "" 
            
            erros_encontrados = []
            dt_xml = str(row.get('DataEmissao_XML', '')).strip()[:10]
            dt_totvs = formata_data_totvs(row.get('Dt.Emiss', ''))
            
            val_simples_planilha = str(row.get('Simples_SN', 'PENDENTE_CONSULTA')).strip().upper()
            val_simples_xml = str(row.get('Simples_XML_Resolvido', 'NÃO')).strip().upper()
            
            # --- DEFINIÇÃO INTELIGENTE DE OPÇÃO PELO SIMPLES ---
            # Se o XML emitido declara explicitamente NÃO (opSimpNac = 1) e a planilha divergia com SIM,
            # a emissão oficial do documento fiscal prevalece para a apuração da nota,
            # evitando apontar indevidamente 'SIMPLES_RET_INDEVIDA' em prestadores comuns.
            if val_simples_xml == 'NÃO' and val_simples_planilha == 'SIM':
                is_simples = False
            else:
                is_simples = (val_simples_planilha == 'SIM') or (val_simples_xml == 'SIM')
            
            cod_item = str(row.get('Cód.Item', '')).strip()
            regra = regras_retencao.get(cod_item)
            
            vlr_totvs = row.get('Valor_TOTVS', 0.0)
            ir_totvs = row.get('IRRF_TOTVS', 0.0)
            cs_totvs = row.get('CSRF_TOTVS', 0.0)
            inss_totvs = row.get('INSS_TOTVS', 0.0)
            iss_totvs = row.get('ISS_TOTVS', 0.0)
            limite_minimo = 10.0
            limite_isencao = 10.0

            teve_erro_simples = False
            teve_erro_regra = False

            # --- NOVA VERIFICAÇÃO DE DIVERGÊNCIA DE CADASTRO DO SIMPLES ---
            if row['_merge'] == 'both':
                # Só checa se não estiver "PENDENTE" nem em branco
                if val_simples_planilha not in ['PENDENTE_CONSULTA', '', 'NAN'] and val_simples_xml not in ['NÃO_INFORMADO', '', 'NAN', 'PENDENTE_CONSULTA']:
                    if val_simples_planilha != val_simples_xml:
                        erros_encontrados.append("CONFLITO_CADASTRO_SIMPLES")

            # 1. Validação de Simples Nacional (Respeitando regras de Anexo IV como Construção Civil/Empreitada)
            if is_simples:
                retencoes_indevidas = []
                if ir_totvs > 0 and row.get('Aliq_IRRF_TOTVS', 0.0) > 0: retencoes_indevidas.append('IR')
                if cs_totvs > 0 and row.get('Aliq_CSRF_TOTVS', 0.0) > 0: retencoes_indevidas.append('CSRF')
                    
                permite_inss = (regra and regra.get('INSS') == 'SIM')
                permite_iss = (regra and regra.get('ISS') == 'SIM')
                
                if inss_totvs > 0 and row.get('Aliq_INSS_TOTVS', 0.0) > 0 and not permite_inss: retencoes_indevidas.append('INSS')
                if iss_totvs > 0 and row.get('Aliq_ISS_TOTVS', 0.0) > 0 and not permite_iss: retencoes_indevidas.append('ISS')
                    
                if retencoes_indevidas:
                    erros_encontrados.append("SIMPLES_RET_INDEVIDA")
                    teve_erro_simples = True

            # 2. Validação pela Planilha de Regras (COD ITEM)
            if regra:
                if not is_simples:
                    if regra['IR'] == 'NÃO' and ir_totvs > 0:
                        erros_encontrados.append('REGRA_IR_INDEVIDO')
                        teve_erro_regra = True
                    elif regra['IR'] == 'SIM' and ir_totvs == 0 and (vlr_totvs * 0.015) >= limite_minimo:
                        erros_encontrados.append('REGRA_IR_FALTA')
                        teve_erro_regra = True
                        
                    if regra['CS'] == 'NÃO' and cs_totvs > 0:
                        erros_encontrados.append('REGRA_CS_INDEVIDO')
                        teve_erro_regra = True
                    elif regra['CS'] == 'SIM' and cs_totvs == 0 and (vlr_totvs * 0.0465) >= limite_minimo:
                        erros_encontrados.append('REGRA_CS_FALTA')
                        teve_erro_regra = True
                        
                if regra['INSS'] == 'NÃO' and inss_totvs > 0:
                    erros_encontrados.append('REGRA_INSS_INDEVIDO')
                    teve_erro_regra = True
                elif regra['INSS'] == 'SIM' and inss_totvs == 0 and (vlr_totvs * 0.11) >= limite_minimo:
                    erros_encontrados.append('REGRA_INSS_FALTA')
                    teve_erro_regra = True
                    
                if regra['ISS'] == 'NÃO' and iss_totvs > 0:
                    erros_encontrados.append('REGRA_ISS_INDEVIDO')
                    teve_erro_regra = True
                elif regra['ISS'] == 'SIM' and iss_totvs == 0:
                    erros_encontrados.append('REGRA_ISS_FALTA')
                    teve_erro_regra = True
            
            if row.get('CSRF_TOTVS', 0.0) > 0:
                aliq_csrf = row.get('Aliq_CSRF_TOTVS', 0.0)
                if not (4.6 <= aliq_csrf <= 4.7):
                    erros_encontrados.append("CSRF_ALIQ_ERRADA")

            # 3. Cruzamento Direto (Quando existe o XML do Governo)
            if row['_merge'] == 'both':
                if dt_xml and dt_totvs and dt_xml != dt_totvs:
                    erros_encontrados.append("DATA_EMISSAO")

                if abs(row.get('Valor_XML', 0.0) - row.get('Valor_TOTVS', 0.0)) > 0.10:
                    erros_encontrados.append("VALOR_BRUTO")
                    
                flag_ret = str(row.get('Flag_Ret_ISS', '1'))
                iss_esperado = row.get('ISS_XML', 0.0) if flag_ret == '2' else 0.0
                if abs(iss_esperado - row.get('ISS_TOTVS', 0.0)) > limite_isencao:
                    erros_encontrados.append("ISS")
                    
                if not teve_erro_simples and not teve_erro_regra:
                    if abs(row.get('IRRF_XML', 0.0) - row.get('IRRF_TOTVS', 0.0)) > limite_isencao:
                        erros_encontrados.append('IR')
                    if abs(row.get('INSS_XML', 0.0) - row.get('INSS_TOTVS', 0.0)) > limite_isencao:
                        erros_encontrados.append('INSS')

            # 4. Fallback (Quando NÃO existe o XML do Governo)
            elif row['_merge'] == 'right_only':
                if not teve_erro_simples and not teve_erro_regra:
                    erros_totvs = []
                    ir_mandat = (regra and regra.get('IR') == 'SIM')
                    inss_mandat = (regra and regra.get('INSS') == 'SIM')
                    iss_mandat = (regra and regra.get('ISS') == 'SIM')
                    cs_mandat = (regra and regra.get('CS') == 'SIM')
                    
                    if row.get('IRRF_TOTVS', 0.0) > limite_isencao and not ir_mandat: erros_totvs.append("IR")
                    if row.get('INSS_TOTVS', 0.0) > limite_isencao and not inss_mandat: erros_totvs.append("INSS")
                    if row.get('ISS_TOTVS', 0.0) > limite_isencao and not iss_mandat: erros_totvs.append("ISS")
                    if row.get('CSRF_TOTVS', 0.0) > limite_isencao and not cs_mandat: erros_totvs.append("CSRF")
                    
                    if erros_totvs:
                        erros_encontrados.append("_".join(erros_totvs))

            return ",".join(erros_encontrados) if erros_encontrados else "OK"
            
        def status_curto(row):
            if row['_merge'] == 'left_only': return 'IGNORAR'
            elif row['Observações'] != 'OK': 
                return 'DIVERGENTE'
            else: return 'CONCILIADO'
            
        def mapear_apontamento(obs):
            if obs in ['OK', '']: return "Sem Erro"
            
            mensagens_regra = []
            if "REGRA_IR" in obs: mensagens_regra.append("IR")
            if "REGRA_CS" in obs: mensagens_regra.append("CSRF")
            if "REGRA_INSS" in obs: mensagens_regra.append("INSS")
            if "REGRA_ISS" in obs: mensagens_regra.append("ISS")
            if mensagens_regra: return "Divergência Regra: " + "/".join(mensagens_regra)

            if "SIMPLES_RET_INDEVIDA" in obs: return "Retenção Indevida (Simples Nacional)"
            if "CSRF_ALIQ_ERRADA" in obs: return "Alíquota CSRF Divergente"
            if "CONFLITO_CADASTRO_SIMPLES" in obs: return "Conflito Cadastro Simples (XML x Receita)"
            
            tem_ir, tem_inss, tem_iss = "IR" in obs, "INSS" in obs, "ISS" in obs
            if tem_ir and tem_iss and tem_inss: return "Verf. IR/ISS/INSS"
            if tem_ir and tem_inss: return "Verf. IR/INSS"
            if tem_ir and tem_iss: return "Verf. IR/ISS"
            if tem_iss and tem_inss: return "Verf. ISS/INSS"
            if tem_ir: return "Verf. IR"
            if tem_iss: return "Verf. ISS"
            if tem_inss: return "Verf. INSS"
            
            if "DATA_EMISSAO" in obs and not "VALOR_BRUTO" in obs: return "Divergência de Data"
            if "VALOR_BRUTO" in obs: return "Divergência de Valor Bruto"
            
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
            
        df_erros = df_lancadas[df_lancadas['Valid. FINAL'] != 'CONCILIADO'].copy()

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
        
        erros_contagem = df_erros['Apontamento_Mapeado'].value_counts().to_dict()
        dados_apontamentos = []
        ordem = 1
        for desc, qtd in sorted(erros_contagem.items(), key=lambda x: x[1], reverse=True):
            dados_apontamentos.append((ordem, desc, qtd))
            ordem += 1
            
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
            for col in range(col_start, col_end + 1): ws.cell(row=row, column=col).border = borda

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
                    if isinstance(val, (int, float)): c.alignment = Alignment(horizontal="center", vertical="center")

        larguras = {'B': 12, 'C': 15, 'D': 18, 'E': 3, 'F': 8, 'G': 35, 'H': 15, 'I': 3, 'J': 18, 'K': 35, 'L': 15, 'M': 3, 'N': 20, 'O': 20}
        for col, width in larguras.items(): ws.column_dimensions[col].width = width

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
            ws.cell(row=row, column=15).number_format = 'R$ #,##0.00'

        ws_val = wb.create_sheet("Validação")
        
        colunas_comparacao = [
            'CNPJ_TOTVS', 'Nota_TOTVS', 'DataEmissao_XML', 'Dt.Emiss', 'Valor_XML', 'Valor_TOTVS',
            'ISS_XML', 'ISS_TOTVS', 'IRRF_XML', 'IRRF_TOTVS', 'INSS_XML', 'INSS_TOTVS',
            'CSRF_TOTVS', 'Aliq_CSRF_TOTVS', 'Chave_Acesso_NFSe', 'NSU_XML'
        ]

        colunas_tecnicas = [
            'Chave', 'CNPJ_XML', 'Nota_XML', 'Nome_XML', 'Origem_XML', '_merge', 
            'Apontamento_Mapeado', 'CNPJ_Merge', 'CNPJ_Limpo', 'Simples_SN', 'Flag_Ret_ISS', 
            'Simples_XML_Resolvido', 'BaseCalculo_XML', 'BaseCalc_TOTVS',
            'Aliq_IRRF_TOTVS', 'Aliq_INSS_TOTVS', 'Aliq_ISS_TOTVS'
        ]
        
        colunas_destaque = []
        for c in ['Usuario', 'Dt.Trans', 'Est', 'Est Fiscal']:
            if c in df_resultado.columns and c not in colunas_destaque:
                colunas_destaque.append(c)

        colunas_originais = [col for col in df_resultado.columns if col not in colunas_tecnicas and col not in colunas_comparacao and col not in ['Observações', 'Valid. FINAL'] and col not in colunas_destaque]
        
        colunas_finais = ['Observações', 'Valid. FINAL'] + colunas_destaque + colunas_comparacao + colunas_originais
        
        ws_val.append(colunas_finais)
        for col_idx, col_name in enumerate(colunas_finais, 1):
            celula = ws_val.cell(row=1, column=col_idx)
            celula.font = font_title
            celula.fill = fill_title
            ws_val.column_dimensions[celula.column_letter].width = 20
            
        if not df_erros.empty:
            df_erros = df_erros.sort_values(by=['Apontamento_Mapeado'])
            df_export = df_erros[colunas_finais].copy().fillna("")
            
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
                
                if "CONFLITO_CADASTRO_SIMPLES" in obs:
                    pintar_celulas("Simples_SN", "Simples_XML_Resolvido")
                
                if "CSRF_ALIQ_ERRADA" in obs or "REGRA_CS" in obs:
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
        ws_val.auto_filter.ref = f"A1:{get_column_letter(ws_val.max_column)}{ws_val.max_row}"

        wb.save(arquivo_saida)
        print(f"[SUCESSO] Dashboard Pixel Perfect gerado em: {arquivo_saida}")
        return arquivo_saida

if __name__ == "__main__":
    RAIZ_PROJETO = Path(__file__).resolve().parent.parent
    config_caminhos = CaminhosProjeto(diretorio_base=RAIZ_PROJETO)
    
    leitor_xml = LeitorResumosNacionais(caminhos=config_caminhos)
    tabela_xml = leitor_xml.carregar()
    
    tabela_totvs = LeitorRelatorioTotvs(caminhos=config_caminhos).processar_relatorio()
    
    cnpjs_para_consultar = set()
    if not tabela_xml.empty and 'CNPJ_XML' in tabela_xml.columns:
        cnpjs_para_consultar.update(tabela_xml['CNPJ_XML'].dropna().tolist())
    if not tabela_totvs.empty and 'CNPJ_TOTVS' in tabela_totvs.columns:
        cnpjs_para_consultar.update(tabela_totvs['CNPJ_TOTVS'].dropna().tolist())
        
    mapa_emit = {}
    if not tabela_totvs.empty and 'CNPJ_TOTVS' in tabela_totvs.columns and 'Emit.' in tabela_totvs.columns:
        temp_map = tabela_totvs.dropna(subset=['CNPJ_TOTVS', 'Emit.']).set_index('CNPJ_TOTVS')['Emit.'].to_dict()
        mapa_emit.update(temp_map)
        
    mapa_simples_xml = {}
    if not tabela_xml.empty and 'CNPJ_XML' in tabela_xml.columns and 'Simples_XML_Resolvido' in tabela_xml.columns:
        temp_simples = tabela_xml.dropna(subset=['CNPJ_XML', 'Simples_XML_Resolvido']).set_index('CNPJ_XML')['Simples_XML_Resolvido'].to_dict()
        mapa_simples_xml.update(temp_simples)

    tabela_simples = LeitorSimplesNacional(caminhos=config_caminhos).carregar(
        cnpjs_alvo=cnpjs_para_consultar, 
        mapa_emitente=mapa_emit,
        mapa_simples_xml=mapa_simples_xml
    )
    
    leitor_regras = LeitorRegrasRetencao(caminhos=config_caminhos)
    regras_retencao = leitor_regras.carregar()
    
    auditor = AuditorFiscal()
    resultado = auditor.cruzar_dados(
        df_xml=tabela_xml, 
        df_totvs=tabela_totvs, 
        df_simples=tabela_simples,
        regras_retencao=regras_retencao
    )
    
    if not resultado.empty:
        print("\n" + "="*50)
        print("📊 RESUMO DE CRUZAMENTO DE DADOS (MATCH)")
        print("="*50)
        qtd_both = len(resultado[resultado['_merge'] == 'both'])
        qtd_so_totvs = len(resultado[resultado['_merge'] == 'right_only'])
        
        print(f"-> Notas processadas do TOTVS: {len(tabela_totvs)}")
        print(f"-> Notas processadas do Governo: {len(tabela_xml)}")
        print(f"-> Notas conciliadas perfeitamente: {qtd_both}")
        print(f"-> Notas lançadas no TOTVS (ainda sem XML no Governo): {qtd_so_totvs}")
        
        config_caminhos.saida_auditoria.mkdir(parents=True, exist_ok=True)
        arquivo_excel = auditor.exportar_relatorio_divergencias(df_resultado=resultado, caminho_saida=config_caminhos.saida_auditoria)
        print("\n[!] Painéis de Serviços Gerados e Guardados na Pasta de Auditoria.")