# src/materials_etl.py
import os
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
        logging.FileHandler(pasta_historico / "etl_materiais.log", encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Sobrescreve o print nativo do Python neste arquivo
def print(*args, **kwargs):
    sep = kwargs.get("sep", " ")
    logger.info(sep.join(map(str, args)))
# --- FIM DA CONFIGURAÇÃO DE LOG ---

load_dotenv()

def ler_arquivo_robusto(caminho: Path) -> pd.DataFrame:
    """
    Lê o CSV emulando perfeitamente o Excel para evitar colunas deslocadas (Header Shift).
    """
    if caminho.suffix.lower() == '.xlsx':
        return pd.read_excel(caminho, dtype=str)
    
    configs = [
        {'sep': ';', 'encoding': 'latin1'},
        {'sep': ';', 'encoding': 'utf-8'},
        {'sep': '\t', 'encoding': 'latin1'},
        {'sep': '\t', 'encoding': 'utf-16'},
        {'sep': ',', 'encoding': 'utf-8'},
    ]
    
    ultimo_erro = None
    for config in configs:
        try:
            # index_col=False evita o clássico problema do Pandas deslocar os cabeçalhos
            df = pd.read_csv(
                caminho, 
                sep=config['sep'], 
                encoding=config['encoding'], 
                dtype=str, 
                on_bad_lines='skip',
                engine='python',
                index_col=False 
            )
            colunas_limpas = [str(c).replace('"', '').replace("'", "").strip() for c in df.columns]
            df.columns = colunas_limpas
            
            # Se encontrou a coluna exata, o layout está perfeito.
            if 'Chave Acesso' in df.columns:
                return df
                
        except Exception as e:
            ultimo_erro = e
            continue
            
    # Fallback se não encontrar o nome exato da coluna
    for config in configs:
        try:
            df = pd.read_csv(caminho, sep=config['sep'], encoding=config['encoding'], dtype=str, on_bad_lines='skip', engine='python', index_col=False)
            if len(df.columns) > 2:
                df.columns = [str(c).replace('"', '').replace("'", "").strip() for c in df.columns]
                return df
        except Exception:
            continue
            
    raise ValueError(f"Não foi possível decodificar o arquivo. Último erro: {ultimo_erro}")


class LeitorQuiriusNFe:
    """Lê todos os relatórios do Quirius espalhados nas subpastas por Estabelecimento"""
    def __init__(self, caminhos: CaminhosProjeto):
        self.pasta_raiz = caminhos.xml_materiais_quirius

    def _limpar_chave(self, chave_bruta) -> str:
        if pd.isna(chave_bruta): return ""
        limpa = ''.join(filter(str.isdigit, str(chave_bruta)))
        return limpa if len(limpa) == 44 else ""

    def _converter_moeda(self, valor):
        if pd.isna(valor): return 0.0
        v_str = str(valor).strip()
        if ',' in v_str and '.' in v_str:
            v_str = v_str.replace('.', '').replace(',', '.')
        elif ',' in v_str:
            v_str = v_str.replace(',', '.')
        try:
            return float(v_str)
        except ValueError:
            return 0.0

    def _encontrar_coluna_valor(self, df: pd.DataFrame) -> str:
        candidatas_exatas = ['Vlr. Total da Nota', 'Vlr. Total da NF-e', 'Valor Total', 'Tot. Nota']
        for cand in candidatas_exatas:
            for col in df.columns:
                if cand.lower() == col.lower() or cand.lower() in col.lower():
                    return col
                    
        for col in df.columns:
            nome_limpo = col.lower()
            if 'vlr' in nome_limpo and 'total' in nome_limpo:
                if 'cbs' not in nome_limpo and 'ibs' not in nome_limpo and 'st' not in nome_limpo and 'produto' not in nome_limpo:
                    return col
        return None

    def carregar(self) -> pd.DataFrame:
        print("[ETL QUIRIUS] Varrendo subpastas em busca de relatórios de Materiais (NF-e)...")
        arquivos = list(self.pasta_raiz.rglob("*.xlsx")) + list(self.pasta_raiz.rglob("*.csv")) + list(self.pasta_raiz.rglob("*.txt"))
        
        if not arquivos:
            print("[AVISO] Nenhum relatório do Quirius encontrado nas pastas.")
            return pd.DataFrame()

        lista_dfs = []
        total = len(arquivos)
        for idx, caminho in enumerate(arquivos, 1):
            try:
                print(f"[ETL QUIRIUS] Lendo arquivo {idx}/{total}: {caminho.name} ...")
                df_temp = ler_arquivo_robusto(caminho)
                df_temp['Pasta_Origem'] = caminho.parent.name
                lista_dfs.append(df_temp)
            except Exception as e:
                print(f"[ERRO QUIRIUS] Falha ao ler {caminho.name}: {e}")

        if not lista_dfs:
            print("[ERRO CRÍTICO] Nenhum arquivo do Quirius pôde ser lido corretamente.")
            return pd.DataFrame()

        print("[ETL QUIRIUS] Consolidando tabelas lidas...")
        df_quirius = pd.concat(lista_dfs, ignore_index=True)
        
        # Proteção contra relatórios duplicados na pasta do Quirius
        df_quirius.drop_duplicates(inplace=True)
        
        col_chave = next((c for c in df_quirius.columns if 'Chave' in c and 'Acesso' in c), None)
        if not col_chave:
            print("[ERRO CRÍTICO] Coluna de Chave de Acesso não encontrada no Quirius!")
            return pd.DataFrame()

        df_quirius['Chave_Acesso'] = df_quirius[col_chave].apply(self._limpar_chave)
        df_quirius = df_quirius[df_quirius['Chave_Acesso'] != ""].copy()

        col_valor = self._encontrar_coluna_valor(df_quirius)
        if col_valor:
            df_quirius['Valor_QRI'] = df_quirius[col_valor].apply(self._converter_moeda)
        else:
            df_quirius['Valor_QRI'] = 0.0

        df_quirius = df_quirius.groupby('Chave_Acesso', as_index=False).agg({
            'Valor_QRI': 'first',
            'Nº NF-e': 'first',
            'CNPJ Emissor': 'first',
            'Razão Social Emissor': 'first',
            'Data Emissão': 'first',
            'Pasta_Origem': 'first'
        })
        
        df_quirius.rename(columns={'Data Emissão': 'Data Emissão Quirius'}, inplace=True)
        print(f"[ETL QUIRIUS] Carregadas e agrupadas {len(df_quirius)} NF-es únicas.")
        return df_quirius


class LeitorOFBMateriais:
    """Lê a extração bruta do TOTVS (OFB) da pasta específica OFB013"""
    def __init__(self, caminhos: CaminhosProjeto):
        self.pasta = caminhos.entrada_totvs / "OFB013"
        self.pasta.mkdir(parents=True, exist_ok=True)

    def _limpar_chave(self, chave_bruta) -> str:
        if pd.isna(chave_bruta): return ""
        # Filter tira a aspa (') e letras, pegando só os 44 números reais
        limpa = ''.join(filter(str.isdigit, str(chave_bruta)))
        return limpa if len(limpa) == 44 else ""

    def _converter_moeda(self, valor):
        if pd.isna(valor): return 0.0
        v_str = str(valor).strip()
        v_str = ''.join(c for c in v_str if c.isdigit() or c in ',.-')
        
        if ',' in v_str and '.' in v_str:
            v_str = v_str.replace('.', '').replace(',', '.')
        elif ',' in v_str:
            v_str = v_str.replace(',', '.')
            
        try: return float(v_str)
        except ValueError: return 0.0

    def carregar(self) -> pd.DataFrame:
        print(f"[ETL TOTVS] Iniciando leitura do relatório OFB na pasta: {self.pasta.name}...")
        arquivos = list(self.pasta.rglob("*.xlsx")) + list(self.pasta.rglob("*.csv")) + list(self.pasta.rglob("*.txt"))
        
        if not arquivos:
            print(f"[AVISO] Nenhum relatório do TOTVS encontrado na pasta {self.pasta.name}.")
            return pd.DataFrame()

        lista_dfs = []
        total = len(arquivos)
        for idx, caminho in enumerate(arquivos, 1):
            try:
                print(f"[ETL TOTVS] Lendo arquivo {idx}/{total}: {caminho.name} ...")
                df_temp = ler_arquivo_robusto(caminho)
                lista_dfs.append(df_temp)
            except Exception as e:
                print(f"[ERRO TOTVS] Falha ao ler {caminho.name}: {e}")

        if not lista_dfs:
            print("[ERRO CRÍTICO] Nenhum arquivo OFB pôde ser lido corretamente.")
            return pd.DataFrame()

        df_totvs = pd.concat(lista_dfs, ignore_index=True)
        
        # Proteção Blindada: Exclui linhas duplicadas oriundas de leitura de arquivos em subpastas de backup
        df_totvs.drop_duplicates(inplace=True)

        col_usuario = next((c for c in df_totvs.columns if c.lower() in ['nome do usuario', 'usuário', 'usuario']), None)
        if col_usuario:
            qtd_antes = len(df_totvs)
            df_totvs = df_totvs[~df_totvs[col_usuario].astype(str).str.upper().str.contains('RPW - SERVICO', na=False)].copy()
            qtd_depois = len(df_totvs)
            if qtd_antes != qtd_depois:
                print(f"[ETL TOTVS] {qtd_antes - qtd_depois} registros de frete (RPW - SERVICO) ignorados para evitar duplicidade.")

        # --- CAÇADOR DE CHAVE INFALÍVEL ---
        col_chave = 'Chave Acesso'
        tem_chave_valida = False
        
        if col_chave in df_totvs.columns:
            amostra = df_totvs[col_chave].dropna().astype(str).head(5000)
            for val in amostra:
                v_clean = ''.join(filter(str.isdigit, val))
                if len(v_clean) == 44:
                    tem_chave_valida = True
                    break
                    
        if not tem_chave_valida:
            print(f"[AVISO] A coluna '{col_chave}' não possui os 44 dígitos. Buscando na tabela...")
            for col in df_totvs.columns:
                amostra = df_totvs[col].dropna().astype(str).head(5000)
                for val in amostra:
                    v_clean = ''.join(filter(str.isdigit, val))
                    if len(v_clean) == 44:
                        col_chave = col
                        tem_chave_valida = True
                        print(f"-> Sucesso! A verdadeira chave de 44 dígitos foi rastreada na coluna: '{col}'")
                        break
                if tem_chave_valida:
                    break
                    
        if not tem_chave_valida:
            print("[ERRO CRÍTICO] Nenhuma coluna com Chaves de Acesso (44 dígitos) foi encontrada no relatório OFB!")
            return pd.DataFrame()

        qtd_bruta = len(df_totvs)
        df_totvs['Chave_Acesso'] = df_totvs[col_chave].apply(self._limpar_chave)
        df_totvs = df_totvs[df_totvs['Chave_Acesso'] != ""].copy()

        if len(df_totvs) == 0 and qtd_bruta > 0:
            print(f"[AVISO] ALERTA: Foram lidas {qtd_bruta} linhas, mas nenhuma chave sobreviveu à limpeza!")

        col_vlr_item = next((c for c in df_totvs.columns if c.lower() in ['vlr item', 'valor item', 'vlr. item', 'vlr.item', 'valor do item']), None)
        if col_vlr_item:
            df_totvs['Valor_Item_OFB'] = df_totvs[col_vlr_item].apply(self._converter_moeda)
        else:
            df_totvs['Valor_Item_OFB'] = 0.0

        col_emissao = next((c for c in df_totvs.columns if c.lower() in ['emissão', 'emissao', 'data emissão', 'dt emissao']), None)
        col_dt_docto = next((c for c in df_totvs.columns if c.lower() in ['dt docto', 'data docto', 'dt. docto', 'dt.docto']), None)
        
        col_docfisc = next((c for c in df_totvs.columns if c.lower() in ['doc fisc', 'docto', 'nota']), None)
        col_nome = next((c for c in df_totvs.columns if c.lower() in ['nome', 'razão social', 'nome abrev']), None)
        col_est = next((c for c in df_totvs.columns if c.lower() in ['estabelecimento', 'est', 'est fiscal']), None)

        agg_dict = {'Valor_Item_OFB': 'sum'}
        if col_docfisc: agg_dict[col_docfisc] = 'first'
        if col_nome: agg_dict[col_nome] = 'first'
        if col_est: agg_dict[col_est] = 'first'
        if col_usuario: agg_dict[col_usuario] = 'first'
        if col_emissao: agg_dict[col_emissao] = 'first'
        if col_dt_docto: agg_dict[col_dt_docto] = 'first'

        df_totvs = df_totvs.groupby('Chave_Acesso', as_index=False).agg(agg_dict)
        
        df_totvs.rename(columns={'Valor_Item_OFB': 'Valor_OFB'}, inplace=True)
        
        if col_emissao: df_totvs.rename(columns={col_emissao: 'Data Emissão TOTVS'}, inplace=True)
        else: df_totvs['Data Emissão TOTVS'] = ""
        
        if col_dt_docto: df_totvs.rename(columns={col_dt_docto: 'Dt Docto'}, inplace=True)
        else: df_totvs['Dt Docto'] = ""
        
        if col_docfisc: df_totvs.rename(columns={col_docfisc: 'Doc Fisc'}, inplace=True)
        else: df_totvs['Doc Fisc'] = ""
        
        if col_nome: df_totvs.rename(columns={col_nome: 'Nome'}, inplace=True)
        else: df_totvs['Nome'] = ""
        
        if col_est: df_totvs.rename(columns={col_est: 'Estabelecimento'}, inplace=True)
        else: df_totvs['Estabelecimento'] = ""
        
        if col_usuario: df_totvs.rename(columns={col_usuario: 'Nome do Usuario'}, inplace=True)
        else: df_totvs['Nome do Usuario'] = ""

        print(f"[ETL TOTVS] Consolidadas {len(df_totvs)} NF-es únicas vindas dos itens.")
        return df_totvs


class AuditorMateriais:
    """Motor de Malha Fina para Notas de Material (QRI x OFB)"""
    
    def cruzar_dados(self, df_quirius: pd.DataFrame, df_totvs: pd.DataFrame) -> pd.DataFrame:
        print("[AUDITORIA] Cruzando Base Quirius vs Base TOTVS pelas Chaves de 44 Dígitos...")
        
        if df_quirius.empty or df_totvs.empty:
            return pd.DataFrame()

        df_cruzamento = pd.merge(df_quirius, df_totvs, on='Chave_Acesso', how='outer', indicator=True)

        qtd_ignoradas = len(df_cruzamento[df_cruzamento['_merge'] == 'left_only'])
        df_cruzamento = df_cruzamento[df_cruzamento['_merge'] != 'left_only'].copy()
        print(f"[AUDITORIA] {qtd_ignoradas} notas presentes no Governo mas não no TOTVS foram ignoradas.")

        def formatar_data(dt_str):
            """Transforma qualquer data para o formato YYYY-MM-DD para comparação justa"""
            if pd.isna(dt_str) or not str(dt_str).strip() or str(dt_str).lower() in ['nan', 'none', 'nat']: 
                return ""
                
            dt_str = str(dt_str).strip()
            
            # Padrão TOTVS Completo (DD/MM/YYYY) -> Converte para YYYY-MM-DD
            if len(dt_str) >= 10 and dt_str[2] == '/' and dt_str[5] == '/':
                return f"{dt_str[6:10]}-{dt_str[3:5]}-{dt_str[0:2]}"
            
            # Padrão TOTVS Curto (DD/MM/YY) -> Converte para YYYY-MM-DD (Evita alarme falso)
            if len(dt_str) == 8 and dt_str[2] == '/' and dt_str[5] == '/':
                return f"20{dt_str[6:8]}-{dt_str[3:5]}-{dt_str[0:2]}"
                
            # Padrão Quirius (YYYY-MM-DD HH:MM:SS) -> Converte para YYYY-MM-DD
            if len(dt_str) >= 10 and dt_str[4] == '-':
                return dt_str[:10]
                
            return dt_str

        def validar_erros(row):
            if row['_merge'] == 'left_only':
                return 'FALTA_NO_TOTVS'
            elif row['_merge'] == 'right_only':
                return 'FALTA_NO_QUIRIUS'
            else:
                v_qri = row.get('Valor_QRI', 0.0)
                dt_qri = str(row.get('Data Emissão Quirius', '')).strip()
                
                if pd.isna(v_qri) or v_qri == 0.0 or not dt_qri or dt_qri.lower() in ['nan', 'nat', 'none']:
                    return 'INFO_FALTANTE_QUIRIUS'

                erros = []
                dif = round(v_qri - row.get('Valor_OFB', 0.0), 2)
                if abs(dif) > 0.10:
                    erros.append('VALOR')
                
                # Validação de Datas (Só compara a Emissão Quirius com Emissão TOTVS)
                dt_qri_fmt = formatar_data(dt_qri)
                dt_ofb_fmt = formatar_data(row.get('Data Emissão TOTVS', ''))
                
                if dt_qri_fmt and dt_ofb_fmt and dt_qri_fmt != dt_ofb_fmt:
                    erros.append('DATA')

                return ",".join(erros) if erros else "OK"

        def status_final(obs):
            if obs == "OK": return "CONCILIADO"
            if obs in ["FALTA_NO_TOTVS", "FALTA_NO_QUIRIUS"]: return "FALTA XML"
            return "DIVERGENTE"

        def mapear_apontamento(obs):
            if obs == "OK": return "Sem Erro"
            if "FALTA_NO_TOTVS" in obs: return "Nota lançada no Governo, não no TOTVS"
            if "FALTA_NO_QUIRIUS" in obs: return "Nota lançada no TOTVS, não no Governo"
            if "INFO_FALTANTE_QUIRIUS" in obs: return "Falta a informação no relatório Quirius"
            
            tem_data = "DATA" in obs
            tem_valor = "VALOR" in obs
            
            if tem_data and tem_valor: return "Divergência de Valor e Data"
            if tem_valor: return "Divergência de Valor"
            if tem_data: return "Divergência de Data"
            return "Outras Divergências"

        df_cruzamento['Observações Brutas'] = df_cruzamento.apply(validar_erros, axis=1)
        df_cruzamento['Valid. FINAL'] = df_cruzamento['Observações Brutas'].apply(status_final)
        df_cruzamento['Apontamento_Mapeado'] = df_cruzamento['Observações Brutas'].apply(mapear_apontamento)
        
        df_cruzamento['Dif. QRI x OFB'] = df_cruzamento['Valor_QRI'].fillna(0) - df_cruzamento['Valor_OFB'].fillna(0)
        
        if 'Nº NF-e' in df_cruzamento.columns and 'Doc Fisc' in df_cruzamento.columns:
            df_cruzamento['Nº NF-e'] = df_cruzamento['Nº NF-e'].fillna(df_cruzamento['Doc Fisc'])
            
        if 'Razão Social Emissor' in df_cruzamento.columns and 'Nome' in df_cruzamento.columns:
            df_cruzamento['Razão Social Emissor'] = df_cruzamento['Razão Social Emissor'].fillna(df_cruzamento['Nome'])

        colunas_finais = [
            'Apontamento_Mapeado', 'Valid. FINAL', 'Nome do Usuario', 'Dt Docto', 'Estabelecimento',
            'Dif. QRI x OFB', 'Valor_QRI', 'Valor_OFB',
            'Data Emissão Quirius', 'Data Emissão TOTVS',
            'Chave_Acesso', 'Nº NF-e', 'Doc Fisc', 
            'CNPJ Emissor', 'Razão Social Emissor', 'Nome', 
            'Pasta_Origem'
        ]
        
        for col in colunas_finais:
            if col not in df_cruzamento.columns:
                df_cruzamento[col] = ""
                
        return df_cruzamento[colunas_finais]

    def exportar_dashboard(self, df_resultado: pd.DataFrame, caminho_saida: Path):
        print("[EXPORTAÇÃO] Gerando Dashboard Executivo de Materiais...")
        if df_resultado.empty: return None

        df_erros = df_resultado[df_resultado['Apontamento_Mapeado'] != 'Sem Erro'].copy()

        if 'Valor_OFB' in df_resultado.columns:
            df_resultado['Valor_OFB'] = pd.to_numeric(df_resultado['Valor_OFB'], errors='coerce').fillna(0)
        
        total_nfs = len(df_resultado)
        total_verificar = len(df_erros)

        est_col = 'Estabelecimento'
        if est_col in df_resultado.columns:
            df_estab = df_resultado.groupby(est_col).size().reset_index(name='Total NFS')
            erros_estab = df_erros.groupby(est_col).size().reset_index(name='Para Verificar')
            df_estab = pd.merge(df_estab, erros_estab, on=est_col, how='left').fillna(0)
            df_estab['Para Verificar'] = df_estab['Para Verificar'].astype(int)
        else:
            df_estab = pd.DataFrame(columns=['Estabelecimento', 'Total NFS', 'Para Verificar'])

        lista_apontamentos = [
            (1, "Divergência de Valor e Data"),
            (2, "Divergência de Valor"),
            (3, "Divergência de Data"),
            (4, "Falta a informação no relatório Quirius"),
            (5, "Nota lançada no TOTVS, não no Governo"),
        ]
        
        erros_contagem = df_erros['Apontamento_Mapeado'].value_counts().to_dict()
        dados_apontamentos = []
        for ordem, desc in lista_apontamentos:
            qtd = erros_contagem.get(desc, 0)
            dados_apontamentos.append((ordem, desc, qtd))

        if 'CNPJ Emissor' in df_resultado.columns and 'Razão Social Emissor' in df_resultado.columns:
            df_forn = df_resultado.groupby(['CNPJ Emissor', 'Razão Social Emissor']).size().reset_index(name='Quantidade')
            df_forn = df_forn.sort_values(by='Quantidade', ascending=False).head(10)
        else:
            df_forn = pd.DataFrame(columns=['CNPJ Emissor', 'Razão Social Emissor', 'Quantidade'])

        if 'Nome do Usuario' in df_resultado.columns:
            df_user = df_resultado.groupby(['Nome do Usuario']).size().reset_index(name='Quantidades')
            df_user = df_user.sort_values(by='Quantidades', ascending=False).head(5)
            df_user.insert(0, 'Cód.', '-') 
        else:
            df_user = pd.DataFrame(columns=['Cód.', 'Nome do Usuario', 'Quantidades'])

        totais_gerais = [
            ("Valor Total NFS (OFB)", df_resultado['Valor_OFB'].sum())
        ]

        arquivo_saida = caminho_saida / "relatorio_materiais.xlsx"
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

        criar_titulo_secao(ws, 2, 4, 2, "RESUMO DE NFS (MATERIAIS)")
        ws.cell(row=3, column=2, value="Total de NFS").font = Font(bold=True)
        ws.cell(row=3, column=3, value=total_nfs).alignment = Alignment(horizontal="left")
        ws.cell(row=4, column=2, value="Total NFS Para Verificar").font = Font(bold=True)
        ws.cell(row=4, column=3, value=total_verificar).alignment = Alignment(horizontal="left")
        
        criar_tabela(ws, 2, 6, ["Estabelecimento", "Total NFS", "Para Verificar"], df_estab.values.tolist())
        
        criar_titulo_secao(ws, 6, 8, 2, "APONTAMENTOS")
        criar_tabela(ws, 6, 3, ["Ordem", "Descrição", "Quantidades"], dados_apontamentos)
        
        criar_titulo_secao(ws, 10, 12, 2, "10 MAIORES FORNECEDORES DE EMISSÃO")
        criar_tabela(ws, 10, 3, ["CNPJ Emit.", "Razão Social", "Quantidade"], df_forn.values.tolist())
        
        row_user = 3 + len(df_forn) + 2
        criar_titulo_secao(ws, 10, 12, row_user, "5 MAIORES USUÁRIOS (LANÇAMENTOS)")
        criar_tabela(ws, 10, row_user+1, ["-", "Nome do Usuario", "Quantidades"], df_user.values.tolist())

        criar_titulo_secao(ws, 14, 15, 2, "TOTAL GERAL")
        criar_tabela(ws, 14, 3, ["Descrição", "Valor"], totais_gerais)
        
        for row in range(4, 4 + len(totais_gerais)):
            c_val = ws.cell(row=row, column=15)
            c_val.number_format = 'R$ #,##0.00'

        # === ABA VALIDAÇÃO ===
        ws_val = wb.create_sheet("Validação")
        
        colunas_finais = list(df_resultado.columns)
        ws_val.append(colunas_finais)
        
        for col_idx, col_name in enumerate(colunas_finais, 1):
            celula = ws_val.cell(row=1, column=col_idx)
            celula.font = font_title
            celula.fill = fill_title
            ws_val.column_dimensions[celula.column_letter].width = 20
            
        df_export = df_erros.fillna("")
        df_export = df_export.sort_values(by=['Valid. FINAL', 'Apontamento_Mapeado'])

        col_idx_map = {col_name: idx for idx, col_name in enumerate(colunas_finais, 1)}
        fill_erro = PatternFill(fill_type="solid", fgColor="FFC7CE") 
        font_erro = Font(color="9C0006", bold=True)                  
        
        for r_idx, r in enumerate(df_export.itertuples(index=False), start=2):
            ws_val.append(list(r))
            
            ws_val.cell(row=r_idx, column=col_idx_map.get('Dif. QRI x OFB', 3)).number_format = 'R$ #,##0.00'
            ws_val.cell(row=r_idx, column=col_idx_map.get('Valor_QRI', 4)).number_format = 'R$ #,##0.00'
            ws_val.cell(row=r_idx, column=col_idx_map.get('Valor_OFB', 5)).number_format = 'R$ #,##0.00'

            obs = str(r[0]) 
            
            def pintar_celulas(col_qri, col_totvs):
                if col_qri in col_idx_map:
                    ws_val.cell(row=r_idx, column=col_idx_map[col_qri]).fill = fill_erro
                    ws_val.cell(row=r_idx, column=col_idx_map[col_qri]).font = font_erro
                if col_totvs in col_idx_map:
                    ws_val.cell(row=r_idx, column=col_idx_map[col_totvs]).fill = fill_erro
                    ws_val.cell(row=r_idx, column=col_idx_map[col_totvs]).font = font_erro

            if obs != 'Sem Erro':
                ws_val.cell(row=r_idx, column=col_idx_map.get('Apontamento_Mapeado', 1)).fill = fill_erro
                ws_val.cell(row=r_idx, column=col_idx_map.get('Apontamento_Mapeado', 1)).font = font_erro
                
                if "Data" in obs: pintar_celulas("Data Emissão Quirius", "Data Emissão TOTVS")
                if "Valor" in obs: pintar_celulas("Valor_QRI", "Valor_OFB")
                if "não no TOTVS" in obs: pintar_celulas("Chave_Acesso", "Nº NF-e")
                if "não no Governo" in obs: pintar_celulas("Chave_Acesso", "Doc Fisc")
                
                if "Falta a informação" in obs:
                    pintar_celulas("Valor_QRI", "Data Emissão Quirius")
                
        ws_val.freeze_panes = "A2"
        ultima_coluna = get_column_letter(ws_val.max_column)
        ws_val.auto_filter.ref = f"A1:{ultima_coluna}{ws_val.max_row}"

        try:
            wb.save(str(arquivo_saida))
            print(f"[SUCESSO] Dashboard Finalizado: {arquivo_saida}")
            return arquivo_saida
        except PermissionError:
            print(f"\n[ERRO CRÍTICO] ACESSO NEGADO!")
            print(f"O arquivo '{str(arquivo_saida)}' está ABERTO em algum lugar.")
            print("Por favor, feche o Excel e rode o script novamente.")
            return None


if __name__ == "__main__":
    RAIZ_PROJETO = Path(__file__).resolve().parent.parent
    caminhos = CaminhosProjeto(diretorio_base=RAIZ_PROJETO)
    caminhos.saida_auditoria.mkdir(parents=True, exist_ok=True)
    
    leitor_qri = LeitorQuiriusNFe(caminhos)
    df_quirius = leitor_qri.carregar()
    
    leitor_ofb = LeitorOFBMateriais(caminhos)
    df_totvs = leitor_ofb.carregar()
    
    auditor = AuditorMateriais()
    df_resultado = auditor.cruzar_dados(df_quirius, df_totvs)
    
    if not df_resultado.empty:
        auditor.exportar_dashboard(df_resultado, caminhos.saida_auditoria)
    else:
        print("\n[AVISO] O processo terminou, mas o relatório Excel não foi gerado.")
        print("        Motivo: Após aplicar a regra de descartar notas não lançadas no TOTVS, sobraram ZERO notas para comparar.")
        print("        Solução: Verifique se o seu relatório OFB013 possui chaves de acesso válidas.")
    
    print("\n" + "="*50)
    print(" PIPELINE DE MATERIAIS CONCLUÍDO COM SUCESSO ")
    print("="*50)