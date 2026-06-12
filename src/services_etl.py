# src/services_etl.py
import pandas as pd
import xml.etree.ElementTree as ET
from pathlib import Path
from config import CaminhosProjeto

class LeitorXmlNacional:
    """
    Classe responsável por escanear o diretório local, interpretar a árvore do XML,
    tratar as barreiras dos Namespaces do SPED e estruturar os dados no Pandas.
    """
    def __init__(self, caminhos: CaminhosProjeto):
        self.caminhos = caminhos
        # Mapeamento do Namespace padrão exigido pela Receita Federal para leitura das tags
        self.namespaces_sped = {'ns': 'http://www.sped.fazenda.gov.br/nfse'}

    def processar_diretorio_xml(self) -> pd.DataFrame:
        print("[ETL] Iniciando a varredura e extração dos arquivos XML locais...")
        
        lista_arquivos_xml = list(self.caminhos.xml_servicos_nacional.glob("*.xml"))
        
        if not lista_arquivos_xml:
            print("[AVISO] Nenhum arquivo XML encontrado. O motor de auditoria operará 100% via contingência da planilha.")
            return pd.DataFrame()

        registros_notas_fiscais = []

        for caminho_arquivo in lista_arquivos_xml:
            try:
                arvore_xml = ET.parse(caminho_arquivo)
                raiz_xml = arvore_xml.getroot()

                # Captura o Número da Nota Fiscal de Serviço
                nodo_numero_nota = raiz_xml.find('.//ns:nNFSe', self.namespaces_sped)
                numero_nota = nodo_numero_nota.text.strip() if nodo_numero_nota is not None else None

                # Captura o CNPJ do Prestador/Emitente dentro do bloco exclusivo <emit>
                nodo_emitente = raiz_xml.find('.//ns:emit', self.namespaces_sped)
                cnpj_prestador = nodo_emitente.find('ns:CNPJ', self.namespaces_sped).text.strip() if nodo_emitente is not None else None

                # Captura o Valor do Serviço Prestado de forma tolerante a layouts
                nodo_valor_servico = raiz_xml.find('.//ns:vServ', self.namespaces_sped)
                if nodo_valor_servico is None:
                    # Fallback para capturar a tag de valor base do bloco principal caso vServ não esteja explícito
                    nodo_valor_servico = raiz_xml.find('.//ns:vBC', self.namespaces_sped)
                
                valor_total_nota = float(nodo_valor_servico.text.strip()) if nodo_valor_servico is not None else 0.0

                if cnpj_prestador and numero_nota:
                    # Montagem da Chave Composta estrutural para execução simétrica do merge com o TOTVS
                    chave_composta_relacionamento = f"{cnpj_prestador}_{numero_nota}"
                    
                    registros_notas_fiscais.append({
                        "Chave": chave_composta_relacionamento,
                        "CNPJ_Prestador": cnpj_prestador,
                        "Numero_Nota": numero_nota,
                        "Valor_Total_XML": valor_total_nota,
                        "Nome_Arquivo_Origem": caminho_arquivo.name
                    })
                else:
                    print(f"[AVISO] Tags estruturais em falta no arquivo: {caminho_arquivo.name}")

            except ET.ParseError:
                print(f"[ERRO] Falha de sintaxe na leitura da estrutura do arquivo: {caminho_arquivo.name}")
            except Exception as erro_leitura:
                print(f"[ERRO] Inconsistência inesperada ao processar o arquivo {caminho_arquivo.name}: {erro_leitura}")

        # Transforma os dicionários carregados em memória em uma tabela do Pandas
        dataframe_xml_consolidado = pd.DataFrame(registros_notas_fiscais)
        print(f"[SUCESSO] Mineração de dados concluída: {len(dataframe_xml_consolidado)} notas fiscais organizadas para o Pandas.\n")
        return dataframe_xml_consolidado


if __name__ == "__main__":
    RAIZ_PROJETO = Path(__file__).resolve().parent.parent
    config_caminhos = CaminhosProjeto(diretorio_base=RAIZ_PROJETO)
    
    analisador_pandas = LeitorXmlNacional(caminhos=config_caminhos)
    tabela_final_pandas = analisador_pandas.processar_diretorio_xml()
    
    if not tabela_final_pandas.empty:
        print(tabela_final_pandas.to_string(index=False))