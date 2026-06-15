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
        self.namespaces_sped = {'ns': 'http://www.sped.fazenda.gov.br/nfse'}

    def _normalizar_numero_nota(self, numero_bruto: str) -> str:
        """
        Remove prefixos do Emissor Nacional (13 dígitos) para garantir
        o match com a digitação no TOTVS.
        Exemplo: '2000000010781' vira '10781'.
        """
        if not numero_bruto:
            return None
            
        numero_limpo = str(numero_bruto).strip()
        
        # Se for uma nota gerada no Emissor Nacional Web/App (13 dígitos)
        if len(numero_limpo) >= 13 and numero_limpo.isdigit():
            # Converte os últimos 9 dígitos para inteiro (matando os zeros) e volta para string
            return str(int(numero_limpo[-9:]))
            
        # Se for uma nota normal de prefeitura (ex: 5152), devolve como está
        return numero_limpo

    def processar_diretorio_xml(self) -> pd.DataFrame:
        print("[ETL] Iniciando a varredura e extração dos arquivos XML locais...")
        
        lista_arquivos_xml = list(self.caminhos.xml_servicos_nacional.glob("*.xml"))
        
        if not lista_arquivos_xml:
            print("[AVISO] Nenhum arquivo XML encontrado na pasta.")
            return pd.DataFrame()

        registros_notas_fiscais = []

        for caminho_arquivo in lista_arquivos_xml:
            try:
                arvore_xml = ET.parse(caminho_arquivo)
                raiz_xml = arvore_xml.getroot()

                # Captura e normaliza o Número da Nota
                nodo_numero = raiz_xml.find('.//ns:nNFSe', self.namespaces_sped)
                numero_nota_bruto = nodo_numero.text.strip() if nodo_numero is not None else None
                numero_nota_normalizado = self._normalizar_numero_nota(numero_nota_bruto)

                # Captura o CNPJ do Prestador
                nodo_emitente = raiz_xml.find('.//ns:emit', self.namespaces_sped)
                cnpj_prestador = nodo_emitente.find('ns:CNPJ', self.namespaces_sped).text.strip() if nodo_emitente is not None else None

                # Captura o Valor Bruto
                nodo_valor_servico = raiz_xml.find('.//ns:vServ', self.namespaces_sped)
                if nodo_valor_servico is None:
                    nodo_valor_servico = raiz_xml.find('.//ns:vBC', self.namespaces_sped)
                
                valor_total_nota = float(nodo_valor_servico.text.strip()) if nodo_valor_servico is not None else 0.0

                if cnpj_prestador and numero_nota_normalizado:
                    # Chave Composta alinhada com o TOTVS
                    chave_composta_relacionamento = f"{cnpj_prestador}_{numero_nota_normalizado}"
                    
                    registros_notas_fiscais.append({
                        "Chave": chave_composta_relacionamento,
                        "CNPJ_Prestador": cnpj_prestador,
                        "Numero_Nota_Original": numero_nota_bruto,
                        "Numero_Nota_TOTVS": numero_nota_normalizado,
                        "Valor_Total_XML": valor_total_nota,
                        "Nome_Arquivo_Origem": caminho_arquivo.name
                    })

            except Exception as erro_leitura:
                print(f"[ERRO] Inconsistência ao processar {caminho_arquivo.name}: {erro_leitura}")

        dataframe_xml_consolidado = pd.DataFrame(registros_notas_fiscais)
        print(f"[SUCESSO] Mineração concluída: {len(dataframe_xml_consolidado)} notas organizadas.\n")
        return dataframe_xml_consolidado

if __name__ == "__main__":
    RAIZ_PROJETO = Path(__file__).resolve().parent.parent
    config_caminhos = CaminhosProjeto(diretorio_base=RAIZ_PROJETO)
    
    analisador_pandas = LeitorXmlNacional(caminhos=config_caminhos)
    tabela_final = analisador_pandas.processar_diretorio_xml()
    
    if not tabela_final.empty:
        print(tabela_final.to_string(index=False))