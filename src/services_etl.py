# src/services_etl.py
import pandas as pd
import xml.etree.ElementTree as ET
from pathlib import Path
from config import CaminhosProjeto

class LeitorXmlNacional:
    """
    Extrai dados do Ambiente Nacional (SPED), incluindo todas as tags de impostos e retenções.
    """
    def __init__(self, caminhos: CaminhosProjeto):
        self.caminhos = caminhos
        self.namespaces_sped = {'ns': 'http://www.sped.fazenda.gov.br/nfse'}

    def _normalizar_numero_nota(self, numero_bruto: str) -> str:
        if not numero_bruto: return None
        numero_limpo = str(numero_bruto).strip()
        if len(numero_limpo) >= 13 and numero_limpo.isdigit():
            return str(int(numero_limpo[-9:]))
        return numero_limpo

    def _extrair_valor(self, raiz, tag):
        """Função auxiliar para buscar tags financeiras e converter para float de forma segura."""
        nodo = raiz.find(f'.//ns:{tag}', self.namespaces_sped)
        if nodo is not None and nodo.text:
            try:
                return float(nodo.text.strip())
            except ValueError:
                return 0.0
        return 0.0

    def processar_diretorio_xml(self) -> pd.DataFrame:
        print("[ETL XML] Iniciando extração do Fato Gerador (Governo)...")
        
        lista_arquivos_xml = list(self.caminhos.xml_servicos_nacional.glob("*.xml"))
        if not lista_arquivos_xml:
            print("[AVISO] Nenhum XML encontrado.")
            return pd.DataFrame()

        registros = []
        for caminho in lista_arquivos_xml:
            try:
                arvore = ET.parse(caminho)
                raiz = arvore.getroot()

                # Identificação
                nodo_numero = raiz.find('.//ns:nNFSe', self.namespaces_sped)
                numero_normalizado = self._normalizar_numero_nota(nodo_numero.text.strip() if nodo_numero is not None else None)

                nodo_emit = raiz.find('.//ns:emit', self.namespaces_sped)
                cnpj = nodo_emit.find('ns:CNPJ', self.namespaces_sped).text.strip() if nodo_emit is not None else None

                # Valor Bruto Principal
                valor_bruto = self._extrair_valor(raiz, 'vServ')
                if valor_bruto == 0.0:
                    valor_bruto = self._extrair_valor(raiz, 'vBC')

                # Puxa a flag que diz se a Britânia deve ou não reter o imposto
                nodo_ret_iss = raiz.find('.//ns:tpRetISSQN', self.namespaces_sped)
                flag_ret_iss = str(nodo_ret_iss.text.strip()) if nodo_ret_iss is not None and nodo_ret_iss.text else '1'

                if cnpj and numero_normalizado:
                    chave = f"{cnpj}_{numero_normalizado}"
                    registros.append({
                        "Chave": chave,
                        "CNPJ_XML": cnpj,
                        "Nota_XML": numero_normalizado,
                        "Valor_XML": valor_bruto,
                        "ISS_XML": self._extrair_valor(raiz, 'vISSQN'),
                        "Flag_Ret_ISS": flag_ret_iss, # 1 = Não Retido, 2 = Retido
                        "IRRF_XML": self._extrair_valor(raiz, 'vRetIRRF'),
                        "INSS_XML": self._extrair_valor(raiz, 'vRetCP'),
                        "PIS_XML": self._extrair_valor(raiz, 'vRetPIS'),
                        "COFINS_XML": self._extrair_valor(raiz, 'vRetCOFINS'),
                        "CSLL_XML": self._extrair_valor(raiz, 'vRetCSLL'),
                        "Origem_XML": caminho.name
                    })
            except Exception as e:
                print(f"[ERRO XML] Falha no arquivo {caminho.name}: {e}")

        df = pd.DataFrame(registros)
        print(f"[SUCESSO XML] {len(df)} notas processadas com mapeamento tributário.\n")
        return df


class LeitorRelatorioTotvs:
    """
    Lê o relatório do ERP, limpando caracteres e capturando o lançamento humano dos impostos.
    """
    def __init__(self, caminhos: CaminhosProjeto):
        self.caminhos = caminhos

    def _limpar_cnpj(self, cnpj_bruto) -> str:
        if pd.isna(cnpj_bruto): return ""
        return ''.join(filter(str.isdigit, str(cnpj_bruto)))

    def _limpar_numero_nota(self, numero_bruto) -> str:
        if pd.isna(numero_bruto): return ""
        numero_str = str(numero_bruto).strip()
        if numero_str.endswith(".0"): numero_str = numero_str[:-2]
        return numero_str

    def _converter_moeda_totvs(self, df, nome_coluna):
        """Troca padrão brasileiro de milhar e vírgula para float americano."""
        if nome_coluna in df.columns:
            return df[nome_coluna].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False).apply(pd.to_numeric, errors='coerce').fillna(0.0)
        return 0.0

    def processar_relatorio(self) -> pd.DataFrame:
        print("[ETL TOTVS] Iniciando leitura do ERP...")
        
        arquivos_totvs = list(self.caminhos.entrada_totvs.glob("*.xlsx")) + list(self.caminhos.entrada_totvs.glob("*.csv"))
        if not arquivos_totvs:
            print("[AVISO] Nenhum relatório do TOTVS encontrado.")
            return pd.DataFrame()

        caminho_relatorio = arquivos_totvs[0]
        try:
            if caminho_relatorio.suffix.lower() == '.xlsx':
                df_totvs = pd.read_excel(caminho_relatorio, dtype=str)
            else:
                df_totvs = pd.read_csv(caminho_relatorio, sep=';', encoding='latin1', dtype=str)
            
            df_totvs.dropna(how='all', inplace=True)
            df_totvs.columns = df_totvs.columns.str.strip().str.replace('\n', ' ').str.replace('\r', '')
            
            if 'Docto' in df_totvs.columns and 'CNPJ emit' in df_totvs.columns:
                
                df_totvs['Nota_TOTVS'] = df_totvs['Docto'].apply(self._limpar_numero_nota)
                df_totvs['CNPJ_TOTVS'] = df_totvs['CNPJ emit'].apply(self._limpar_cnpj)
                
                # Conversão de todas as colunas financeiras
                df_totvs['Valor_TOTVS'] = self._converter_moeda_totvs(df_totvs, 'Tot. Nota')
                df_totvs['ISS_TOTVS'] = self._converter_moeda_totvs(df_totvs, 'ISS')
                df_totvs['IRRF_TOTVS'] = self._converter_moeda_totvs(df_totvs, 'IRRF')
                df_totvs['INSS_TOTVS'] = self._converter_moeda_totvs(df_totvs, 'INSS')
                df_totvs['PIS_TOTVS'] = self._converter_moeda_totvs(df_totvs, 'PIS')
                df_totvs['COFINS_TOTVS'] = self._converter_moeda_totvs(df_totvs, 'COFINS')
                df_totvs['CSLL_TOTVS'] = self._converter_moeda_totvs(df_totvs, 'Valor CS')
                
                df_totvs['Chave'] = df_totvs.apply(lambda row: f"{row['CNPJ_TOTVS']}_{row['Nota_TOTVS']}" if row['CNPJ_TOTVS'] and row['Nota_TOTVS'] else None, axis=1)
                df_totvs = df_totvs.dropna(subset=['Chave']).copy()
                
                colunas_export = ['Chave', 'Nome Abrev', 'Valor_TOTVS', 'ISS_TOTVS', 'IRRF_TOTVS', 'INSS_TOTVS', 'PIS_TOTVS', 'COFINS_TOTVS', 'CSLL_TOTVS']
                print(f"[SUCESSO TOTVS] {len(df_totvs)} lançamentos processados.\n")
                return df_totvs[colunas_export]
            else:
                print("[ERRO TOTVS] Colunas obrigatórias não encontradas.")
                return pd.DataFrame()
        except Exception as e:
            print(f"[ERRO TOTVS] Falha ao processar relatório: {e}")
            return pd.DataFrame()


class AuditorFiscal:
    """
    O Motor de Malha Fina. Compara Fato Gerador x Lançamento Contábil centavo por centavo.
    """
    def cruzar_dados(self, df_xml: pd.DataFrame, df_totvs: pd.DataFrame) -> pd.DataFrame:
        print("[AUDITORIA] Rodando Malha Fina Tributária...")
        
        if df_xml.empty or df_totvs.empty:
            return pd.DataFrame()
            
        df_cruzamento = pd.merge(df_xml, df_totvs, on='Chave', how='outer', indicator=True)
        
        def validar_impostos(row):
            if row['_merge'] == 'left_only':
                return "ALERTA CRÍTICO: Nota no Governo, NÃO LANÇADA no ERP."
            elif row['_merge'] == 'right_only':
                return "ALERTA: Documento no TOTVS, não achado no Governo."
            
            erros_encontrados = []
            
            # Validação do Valor Principal
            if abs(row['Valor_XML'] - row['Valor_TOTVS']) > 0.10:
                erros_encontrados.append(f"Bruto(XML: {row['Valor_XML']:.2f} | ERP: {row['Valor_TOTVS']:.2f})")
                
            # Validação INTELIGENTE do ISS (Olhando para a Flag de Retenção)
            flag_ret = str(row.get('Flag_Ret_ISS', '1'))
            # Só espera que o TOTVS tenha valor de ISS se a flag oficial for '2' (Retido)
            iss_esperado_totvs = row.get('ISS_XML', 0.0) if flag_ret == '2' else 0.0
            val_iss_totvs = row.get('ISS_TOTVS', 0.0)
            
            if abs(iss_esperado_totvs - val_iss_totvs) > 0.10:
                texto_regra = f"{iss_esperado_totvs:.2f} (Retido)" if flag_ret == '2' else "0.00 (Não Retido)"
                erros_encontrados.append(f"ISS(XML Esperava: {texto_regra} | ERP Lançou: {val_iss_totvs:.2f})")
                
            # Validação dos Impostos Federais (Já são exclusivos de retenção nativamente no XML)
            matriz_federais = [
                ('IRRF', 'IRRF_XML', 'IRRF_TOTVS'),
                ('INSS', 'INSS_XML', 'INSS_TOTVS'),
                ('PIS', 'PIS_XML', 'PIS_TOTVS'),
                ('COFINS', 'COFINS_XML', 'COFINS_TOTVS'),
                ('CSLL', 'CSLL_XML', 'CSLL_TOTVS')
            ]
            
            for nome_imposto, col_xml, col_erp in matriz_federais:
                val_xml = row.get(col_xml, 0.0)
                val_erp = row.get(col_erp, 0.0)
                if abs(val_xml - val_erp) > 0.10:
                    erros_encontrados.append(f"{nome_imposto}(XML: {val_xml:.2f} | ERP: {val_erp:.2f})")
            
            if erros_encontrados:
                return "DIVERGÊNCIA TRIBUTÁRIA -> " + " ; ".join(erros_encontrados)
            else:
                return "OK: 100% Conciliado (Bruto e Impostos Bateram)."
                
        df_cruzamento['Status_Auditoria'] = df_cruzamento.apply(validar_impostos, axis=1)
        print("[SUCESSO] Auditoria Concluída!")
        return df_cruzamento

if __name__ == "__main__":
    RAIZ_PROJETO = Path(__file__).resolve().parent.parent
    config_caminhos = CaminhosProjeto(diretorio_base=RAIZ_PROJETO)
    
    leitor_xml = LeitorXmlNacional(caminhos=config_caminhos)
    tabela_xml = leitor_xml.processar_diretorio_xml()
    
    leitor_totvs = LeitorRelatorioTotvs(caminhos=config_caminhos)
    tabela_totvs = leitor_totvs.processar_relatorio()
    
    auditor = AuditorFiscal()
    resultado = auditor.cruzar_dados(df_xml=tabela_xml, df_totvs=tabela_totvs)
    
    if not resultado.empty:
        pd.set_option('display.max_colwidth', None)
        colunas_visuais = ['Chave', 'Nome Abrev', 'Status_Auditoria']
        
        print("\n" + "="*90)
        print("                        RELATÓRIO DE MALHA FINA - BRITÂNIA")
        print("="*90)
        print(resultado[colunas_visuais].to_string(index=False))