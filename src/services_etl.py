# src/services_etl.py
import pandas as pd
import xml.etree.ElementTree as ET
from pathlib import Path
from config import CaminhosProjeto

class LeitorXmlNacional:
    """
    Classe responsável por extrair dados dos XMLs Nacionais,
    tratar Namespaces e montar a tabela de Fonte da Verdade.
    """
    def __init__(self, caminhos: CaminhosProjeto):
        self.caminhos = caminhos
        self.namespaces_sped = {'ns': 'http://www.sped.fazenda.gov.br/nfse'}

    def _normalizar_numero_nota(self, numero_bruto: str) -> str:
        """Remove o prefixo de 13 dígitos do Ambiente Nacional para matching com o TOTVS."""
        if not numero_bruto:
            return None
        numero_limpo = str(numero_bruto).strip()
        if len(numero_limpo) >= 13 and numero_limpo.isdigit():
            return str(int(numero_limpo[-9:]))
        return numero_limpo

    def processar_diretorio_xml(self) -> pd.DataFrame:
        print("[ETL XML] Iniciando a extração dos XMLs Nacionais...")
        
        lista_arquivos_xml = list(self.caminhos.xml_servicos_nacional.glob("*.xml"))
        if not lista_arquivos_xml:
            print("[AVISO] Nenhum XML encontrado na pasta.")
            return pd.DataFrame()

        registros = []
        for caminho in lista_arquivos_xml:
            try:
                arvore = ET.parse(caminho)
                raiz = arvore.getroot()

                # 1. Busca e limpa o Número da Nota
                nodo_numero = raiz.find('.//ns:nNFSe', self.namespaces_sped)
                numero_bruto = nodo_numero.text.strip() if nodo_numero is not None else None
                numero_normalizado = self._normalizar_numero_nota(numero_bruto)

                # 2. Busca o CNPJ do Emissor
                nodo_emit = raiz.find('.//ns:emit', self.namespaces_sped)
                cnpj = nodo_emit.find('ns:CNPJ', self.namespaces_sped).text.strip() if nodo_emit is not None else None

                # 3. Busca o Valor Bruto (Base de Cálculo), ignorando o Líquido
                nodo_valor = raiz.find('.//ns:vServ', self.namespaces_sped)
                if nodo_valor is None:
                    nodo_valor = raiz.find('.//ns:vBC', self.namespaces_sped)
                valor = float(nodo_valor.text.strip()) if nodo_valor is not None else 0.0

                if cnpj and numero_normalizado:
                    chave = f"{cnpj}_{numero_normalizado}"
                    registros.append({
                        "Chave": chave,
                        "CNPJ_XML": cnpj,
                        "Nota_XML": numero_normalizado,
                        "Valor_XML": valor,
                        "Origem_XML": caminho.name
                    })
            except Exception as e:
                print(f"[ERRO XML] Falha no arquivo {caminho.name}: {e}")

        df = pd.DataFrame(registros)
        print(f"[SUCESSO XML] {len(df)} notas processadas.\n")
        return df


class LeitorRelatorioTotvs:
    """
    Classe responsável por ler o export do TOTVS (REB405), limpar sujeiras
    (CNPJs formatados, espaços) e montar a tabela de Lançamentos Contábeis.
    """
    def __init__(self, caminhos: CaminhosProjeto):
        self.caminhos = caminhos

    def _limpar_cnpj(self, cnpj_bruto) -> str:
        """Remove pontos, barras e traços do CNPJ exportado pelo TOTVS."""
        if pd.isna(cnpj_bruto):
            return ""
        return ''.join(filter(str.isdigit, str(cnpj_bruto)))

    def _limpar_numero_nota(self, numero_bruto) -> str:
        """Garante que o número da nota seja uma string limpa sem decimais."""
        if pd.isna(numero_bruto):
            return ""
        numero_str = str(numero_bruto).strip()
        if numero_str.endswith(".0"):
            numero_str = numero_str[:-2]
        return numero_str

    def processar_relatorio(self) -> pd.DataFrame:
        print("[ETL TOTVS] Iniciando leitura do relatório REB405...")
        
        # Aceita relatórios tanto em formato Excel (testes) quanto CSV (produção)
        arquivos_totvs = list(self.caminhos.entrada_totvs.glob("*.xlsx")) + list(self.caminhos.entrada_totvs.glob("*.csv"))
        
        if not arquivos_totvs:
            print("[AVISO] Nenhum relatório do TOTVS encontrado na pasta entrada_totvs.")
            return pd.DataFrame()

        caminho_relatorio = arquivos_totvs[0]
        extensao = caminho_relatorio.suffix.lower()
        
        try:
            print(f"[ETL TOTVS] Carregando o arquivo: {caminho_relatorio.name}")
            
            # Motor híbrido para abrir o arquivo correto sem quebrar por codificação
            if extensao == '.xlsx':
                df_totvs = pd.read_excel(caminho_relatorio, dtype=str)
            else:
                df_totvs = pd.read_csv(caminho_relatorio, sep=';', encoding='latin1', dtype=str)
            
            df_totvs.dropna(how='all', inplace=True)
            
            if 'Docto' in df_totvs.columns and 'FiscalEmit.CNPJ' in df_totvs.columns and 'Tot. Nota' in df_totvs.columns:
                
                df_totvs['Nota_TOTVS'] = df_totvs['Docto'].apply(self._limpar_numero_nota)
                df_totvs['CNPJ_TOTVS'] = df_totvs['FiscalEmit.CNPJ'].apply(self._limpar_cnpj)
                
                # Conversão de moeda (TOTVS Brasileiro: tira os pontos de milhar e troca vírgula por ponto decimal)
                df_totvs['Valor_TOTVS'] = df_totvs['Tot. Nota'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
                
                # Montagem da Chave Composta
                df_totvs['Chave'] = df_totvs.apply(
                    lambda row: f"{row['CNPJ_TOTVS']}_{row['Nota_TOTVS']}" if row['CNPJ_TOTVS'] and row['Nota_TOTVS'] else None, 
                    axis=1
                )
                
                df_totvs = df_totvs.dropna(subset=['Chave']).copy()
                df_auditoria = df_totvs[['Chave', 'CNPJ_TOTVS', 'Nota_TOTVS', 'Valor_TOTVS', 'emitNome']]
                
                print(f"[SUCESSO TOTVS] {len(df_auditoria)} lançamentos contábeis processados.\n")
                return df_auditoria
            
            else:
                print(f"[ERRO TOTVS] As colunas obrigatórias ('Docto', 'FiscalEmit.CNPJ', 'Tot. Nota') não foram encontradas.")
                return pd.DataFrame()
                
        except Exception as e:
            print(f"[ERRO TOTVS] Falha crítica ao processar o relatório: {e}")
            return pd.DataFrame()


class AuditorFiscal:
    """
    Classe final que cruza as tabelas (XML x TOTVS) e gera o diagnóstico.
    """
    def cruzar_dados(self, df_xml: pd.DataFrame, df_totvs: pd.DataFrame) -> pd.DataFrame:
        print("[AUDITORIA] Cruzando Base Nacional com o Lançamento do TOTVS...")
        
        if df_xml.empty or df_totvs.empty:
            print("[AVISO] Auditoria cancelada. Faltam dados do XML ou do TOTVS.")
            return pd.DataFrame()
            
        # O MERGE (O coração da aplicação)
        df_cruzamento = pd.merge(df_xml, df_totvs, on='Chave', how='outer', indicator=True)
        
        def definir_status(row):
            if row['_merge'] == 'left_only':
                return "ALERTA: Documento no Governo, mas NÃO LANÇADO no TOTVS."
            elif row['_merge'] == 'right_only':
                return "ALERTA: Documento no TOTVS, mas NÃO ENCONTRADO na Base Nacional."
            else:
                diferenca_valor = abs(row['Valor_XML'] - row['Valor_TOTVS'])
                if diferenca_valor > 0.10: # Margem de segurança contra arredondamentos bobos
                    return f"DIVERGÊNCIA DE VALOR: XML (R${row['Valor_XML']}) / TOTVS (R${row['Valor_TOTVS']})."
                return "OK: Lançamento Conciliado."
                
        df_cruzamento['Status_Auditoria'] = df_cruzamento.apply(definir_status, axis=1)
        
        print("[SUCESSO] Cruzamento Finalizado!")
        return df_cruzamento

# ==============================================================================
# BLOCO DE EXECUÇÃO
# ==============================================================================
if __name__ == "__main__":
    RAIZ_PROJETO = Path(__file__).resolve().parent.parent
    config_caminhos = CaminhosProjeto(diretorio_base=RAIZ_PROJETO)
    
    # 1. Extrai a Fonte da Verdade (Governo)
    leitor_xml = LeitorXmlNacional(caminhos=config_caminhos)
    tabela_xml = leitor_xml.processar_diretorio_xml()
    
    # 2. Extrai o Fato Lançado (ERP)
    leitor_totvs = LeitorRelatorioTotvs(caminhos=config_caminhos)
    tabela_totvs = leitor_totvs.processar_relatorio()
    
    # 3. Cruza os mundos
    auditor = AuditorFiscal()
    resultado = auditor.cruzar_dados(df_xml=tabela_xml, df_totvs=tabela_totvs)
    
    if not resultado.empty:
        colunas_visuais = ['Chave', 'Valor_XML', 'Valor_TOTVS', 'Status_Auditoria']
        print("\n--- RESULTADO DA AUDITORIA ---")
        print(resultado[colunas_visuais].to_string(index=False))