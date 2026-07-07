# src/materials_etl.py
import os
import pandas as pd
from pathlib import Path
from dotenv import load_dotenv

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from config import CaminhosProjeto

load_dotenv()

def ler_arquivo_robusto(caminho: Path) -> pd.DataFrame:
    """
    Função blindada para ler arquivos com diferentes separadores, encodings 
    e ignorar o "bug das aspas" em descrições de produtos.
    """
    if caminho.suffix.lower() == '.xlsx':
        return pd.read_excel(caminho, dtype=str)
    
    configs = [
        {'sep': '\t', 'encoding': 'latin1'},
        {'sep': '\t', 'encoding': 'utf-8'},
        {'sep': '\t', 'encoding': 'utf-16'},
        {'sep': ';', 'encoding': 'latin1'},
        {'sep': ';', 'encoding': 'utf-8'},
        {'sep': ',', 'encoding': 'utf-8'},
    ]
    
    ultimo_erro = None
    for config in configs:
        try:
            df = pd.read_csv(
                caminho, 
                sep=config['sep'], 
                encoding=config['encoding'], 
                dtype=str, 
                quoting=3, 
                on_bad_lines='skip'
            )
            if len(df.columns) > 2:
                df.columns = [str(c).replace('"', '').replace("'", "").strip() for c in df.columns]
                return df
        except Exception as e:
            ultimo_erro = e
            continue
            
    raise ValueError(f"Não foi possível decodificar o arquivo. Formato desconhecido. Último erro: {ultimo_erro}")


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
        # Prioridade máxima: Buscar pelo nome EXATO do relatório do Governo
        candidatas_exatas = ['Vlr. Total da Nota', 'Vlr. Total da NF-e', 'Valor Total', 'Tot. Nota']
        for cand in candidatas_exatas:
            for col in df.columns:
                if cand.lower() == col.lower() or cand.lower() in col.lower():
                    return col
                    
        # Fallback genérico blindado contra CBS/IBS/ST
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
        for caminho in arquivos:
            try:
                df_temp = ler_arquivo_robusto(caminho)
                df_temp['Pasta_Origem'] = caminho.parent.name
                lista_dfs.append(df_temp)
            except Exception as e:
                print(f"[ERRO QUIRIUS] Falha ao ler {caminho.name}: {e}")

        if not lista_dfs:
            print("[ERRO CRÍTICO] Nenhum arquivo do Quirius pôde ser lido corretamente.")
            return pd.DataFrame()

        df_quirius = pd.concat(lista_dfs, ignore_index=True)
        
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

        # CORREÇÃO DA MULTIPLICAÇÃO DE VALOR (MUDOU DE 'sum' PARA 'first')
        # A coluna "Vlr. Total da Nota" já tem o valor cheio repetido em cada item.
        df_quirius = df_quirius.groupby('Chave_Acesso', as_index=False).agg({
            'Valor_QRI': 'first',
            'Nº NF-e': 'first',
            'CNPJ Emissor': 'first',
            'Razão Social Emissor': 'first',
            'Data Emissão': 'first',
            'Pasta_Origem': 'first'
        })

        print(f"[ETL QUIRIUS] Carregadas e agrupadas {len(df_quirius)} NF-es únicas.")
        return df_quirius


class LeitorOFBMateriais:
    """Lê a extração bruta do TOTVS (OFB) da pasta específica OFB013"""
    def __init__(self, caminhos: CaminhosProjeto):
        self.pasta = caminhos.entrada_totvs / "OFB013"
        self.pasta.mkdir(parents=True, exist_ok=True)

    def _limpar_chave(self, chave_bruta) -> str:
        if pd.isna(chave_bruta): return ""
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
        arquivos = list(self.pasta.glob("*.xlsx")) + list(self.pasta.glob("*.csv")) + list(self.pasta.glob("*.txt"))
        
        if not arquivos:
            print(f"[AVISO] Nenhum relatório do TOTVS encontrado na pasta {self.pasta.name}.")
            return pd.DataFrame()

        lista_dfs = []
        for caminho in arquivos:
            try:
                df_temp = ler_arquivo_robusto(caminho)
                lista_dfs.append(df_temp)
            except Exception as e:
                print(f"[ERRO TOTVS] Falha ao ler {caminho.name}: {e}")

        if not lista_dfs:
            print("[ERRO CRÍTICO] Nenhum arquivo OFB pôde ser lido corretamente.")
            return pd.DataFrame()

        df_totvs = pd.concat(lista_dfs, ignore_index=True)

        if 'Chave Acesso' not in df_totvs.columns:
            print("[ERRO CRÍTICO] Coluna 'Chave Acesso' não encontrada no relatório OFB!")
            return pd.DataFrame()

        df_totvs['Chave_Acesso'] = df_totvs['Chave Acesso'].apply(self._limpar_chave)
        df_totvs = df_totvs[df_totvs['Chave_Acesso'] != ""].copy()

        df_totvs['Valor_Item_OFB'] = df_totvs['Vlr Item'].apply(self._converter_moeda)

        # Para o TOTVS, a nota está dividida nos itens "Vlr Item", portanto AQUI usamos o sum()!
        df_totvs = df_totvs.groupby('Chave_Acesso', as_index=False).agg({
            'Valor_Item_OFB': 'sum',
            'Doc Fisc': 'first',
            'Nome': 'first',
            'Estabelecimento': 'first',
            'Nome do Usuario': 'first'
        })
        
        df_totvs.rename(columns={'Valor_Item_OFB': 'Valor_OFB'}, inplace=True)
        print(f"[ETL TOTVS] Consolidadas {len(df_totvs)} NF-es únicas vindas dos itens.")
        return df_totvs


class AuditorMateriais:
    """Motor de Malha Fina para Notas de Material (QRI x OFB)"""
    
    def cruzar_dados(self, df_quirius: pd.DataFrame, df_totvs: pd.DataFrame) -> pd.DataFrame:
        print("[AUDITORIA] Cruzando Base Quirius vs Base TOTVS pelas Chaves de 44 Dígitos...")
        if df_quirius.empty or df_totvs.empty:
            return pd.DataFrame()

        df_cruzamento = pd.merge(df_quirius, df_totvs, on='Chave_Acesso', how='outer', indicator=True)

        def validar_status(row):
            if row['_merge'] == 'left_only':
                return 'FALTA NO TOTVS'
            elif row['_merge'] == 'right_only':
                return 'FALTA NO QUIRIUS'
            else:
                dif = round(row.get('Valor_QRI', 0.0) - row.get('Valor_OFB', 0.0), 2)
                if abs(dif) > 0.10:
                    return 'DIVERGÊNCIA DE VALOR'
                return 'CONCILIADO OK'

        df_cruzamento['Validação'] = df_cruzamento.apply(validar_status, axis=1)
        df_cruzamento['Dif. QRI x OFB'] = df_cruzamento['Valor_QRI'].fillna(0) - df_cruzamento['Valor_OFB'].fillna(0)
        
        if 'Nº NF-e' in df_cruzamento.columns and 'Doc Fisc' in df_cruzamento.columns:
            df_cruzamento['Nº NF-e'] = df_cruzamento['Nº NF-e'].fillna(df_cruzamento['Doc Fisc'])
            
        if 'Razão Social Emissor' in df_cruzamento.columns and 'Nome' in df_cruzamento.columns:
            df_cruzamento['Razão Social Emissor'] = df_cruzamento['Razão Social Emissor'].fillna(df_cruzamento['Nome'])

        colunas_finais = [
            'Validação', 'Dif. QRI x OFB', 'Valor_QRI', 'Valor_OFB',
            'Chave_Acesso', 'Nº NF-e', 'Doc Fisc', 'Data Emissão', 
            'CNPJ Emissor', 'Razão Social Emissor', 'Nome', 
            'Pasta_Origem', 'Estabelecimento', 'Nome do Usuario'
        ]
        
        for col in colunas_finais:
            if col not in df_cruzamento.columns:
                df_cruzamento[col] = ""
                
        return df_cruzamento[colunas_finais].sort_values(by='Validação')

    def exportar_dashboard(self, df_resultado: pd.DataFrame, caminho_saida: Path):
        print("[EXPORTAÇÃO] Gerando Dashboard Pixel Perfect de Materiais...")
        
        arquivo_saida = caminho_saida / "Dashboard_Materiais_Validacao.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Validação OFB x Quirius"
        
        colunas = list(df_resultado.columns)
        ws.append(colunas)
        
        fill_header = PatternFill(fill_type="solid", fgColor="0E2841")
        font_header = Font(bold=True, color="FFFFFF")
        
        for col_idx, col_name in enumerate(colunas, 1):
            celula = ws.cell(row=1, column=col_idx)
            celula.font = font_header
            celula.fill = fill_header
            ws.column_dimensions[get_column_letter(col_idx)].width = 22

        fill_erro = PatternFill(fill_type="solid", fgColor="FFC7CE") 
        font_erro = Font(color="9C0006", bold=True)
        fill_ok = PatternFill(fill_type="solid", fgColor="C6EFCE")
        font_ok = Font(color="006100", bold=True)

        for r_idx, r in enumerate(df_resultado.itertuples(index=False), start=2):
            ws.append(list(r))
            
            ws.cell(row=r_idx, column=2).number_format = 'R$ #,##0.00'
            ws.cell(row=r_idx, column=3).number_format = 'R$ #,##0.00'
            ws.cell(row=r_idx, column=4).number_format = 'R$ #,##0.00'

            status = str(r[0])
            celula_status = ws.cell(row=r_idx, column=1)
            
            if status == 'CONCILIADO OK':
                celula_status.fill = fill_ok
                celula_status.font = font_ok
            else:
                celula_status.fill = fill_erro
                celula_status.font = font_erro
                for c_idx in range(1, len(colunas) + 1):
                    ws.cell(row=r_idx, column=c_idx).fill = fill_erro
                    
        ws.freeze_panes = "A2"
        ultima_coluna = get_column_letter(ws.max_column)
        ws.auto_filter.ref = f"A1:{ultima_coluna}{ws.max_row}"

        wb.save(arquivo_saida)
        print(f"[SUCESSO] Dashboard Finalizado: {arquivo_saida}")


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