import os
import time
import requests
import base64
import gzip
import xml.etree.ElementTree as ET
import xml.dom.minidom
import logging
import pandas as pd
from pathlib import Path
from config import settings

class SincronizadorGoverno:
    def __init__(self):
        self.cnpj = settings.cnpj_empresa
        self.url_base = "https://adn.nfse.gov.br/contribuintes/DFe" 
        
        if settings.cert_pub_path and settings.cert_key_path:
            self.cert_path = (settings.cert_pub_path, settings.cert_key_path)
        else:
            self.cert_path = None
        
        # Configuração de Pastas e Caminhos
        raiz_projeto = Path(__file__).resolve().parent.parent
        self.pasta_historico = raiz_projeto / "dados" / "historico"
        self.pasta_historico.mkdir(parents=True, exist_ok=True)
        self.arquivo_nsu = self.pasta_historico / "ultimo_nsu.txt"
        
        self.pasta_destino = raiz_projeto / "dados" / "xml_servicos_nacional"
        self.pasta_destino.mkdir(parents=True, exist_ok=True)
        
        # Arquivo Excel onde os dados serão consolidados
        self.arquivo_excel = self.pasta_destino / "Resumo_Notas_Nacional.xlsx"
        
        # Configuração do Logger
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(self.pasta_historico / "sincronizador.log", encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger("SincronizadorNacional")

    def ler_ultimo_nsu(self) -> int:
        if self.arquivo_nsu.exists():
            try:
                with open(self.arquivo_nsu, "r") as f:
                    return int(f.read().strip())
            except ValueError:
                self.logger.warning("Arquivo NSU corrompido. Reiniciando do 0.")
        return 0

    def salvar_ultimo_nsu(self, nsu: int):
        with open(self.arquivo_nsu, "w") as f:
            f.write(str(nsu))

    def _descompactar_e_formatar(self, conteudo_base64: str) -> str:
        """Descompacta o GZIP em Base64 e formata o XML para ficar legível."""
        try:
            dados_binarios = base64.b64decode(conteudo_base64)
            xml_bruto = gzip.decompress(dados_binarios).decode('utf-8')
            
            dom = xml.dom.minidom.parseString(xml_bruto)
            xml_formatado = dom.toprettyxml(indent="  ")
            
            linhas = [linha for linha in xml_formatado.split('\n') if linha.strip()]
            return '\n'.join(linhas)
        except Exception as e:
            self.logger.error(f"Erro ao descompactar documento: {e}")
            return conteudo_base64
            
    def _converter_xml_para_dict(self, xml_string: str, nsu: int) -> dict:
        """Achata (flatten) as tags do XML para transformá-las em colunas do Excel"""
        try:
            raiz = ET.fromstring(xml_string)
            dados = {'NSU_API': str(nsu)}

            def achatar(nodo, prefixo=''):
                for filho in nodo:
                    # Remove o namespace da tag (ex: {http://www.sped.fazenda.gov.br/nfse}infNFSe)
                    tag = filho.tag.split('}')[-1]
                    chave = f"{prefixo}_{tag}" if prefixo else tag
                    
                    if len(filho) > 0:
                        achatar(filho, chave)
                    elif filho.text and filho.text.strip():
                        dados[chave] = filho.text.strip()

            achatar(raiz)
            return dados
        except Exception as e:
            self.logger.error(f"Erro ao converter XML do NSU {nsu} para dicionário: {e}")
            return None

    def _atualizar_excel(self, notas_extraidas: list):
        """Salva a lista de notas num arquivo Excel, preservando o histórico existente."""
        if not notas_extraidas:
            return
            
        try:
            df_novos = pd.DataFrame(notas_extraidas)
            
            if self.arquivo_excel.exists():
                df_existente = pd.read_excel(self.arquivo_excel, dtype=str)
                df_final = pd.concat([df_existente, df_novos], ignore_index=True)
                
                # Remove possíveis duplicatas baseadas no NSU
                if 'NSU_API' in df_final.columns:
                    df_final = df_final.drop_duplicates(subset=['NSU_API'], keep='last')
            else:
                df_final = df_novos
                
            df_final.to_excel(self.arquivo_excel, index=False)
            self.logger.info(f"Sucesso! {len(notas_extraidas)} nota(s) processada(s) e salva(s) em {self.arquivo_excel.name}.")
        except Exception as e:
            self.logger.error(f"Falha ao salvar os dados no Excel: {e}")

    def _sincronizar_xmls_locais(self):
        """Lê a pasta de XMLs e joga no Excel aqueles que já foram baixados mas não estão na planilha."""
        self.logger.info("Verificando se há XMLs locais pendentes de integração com o Excel...")
        
        nsus_no_excel = set()
        if self.arquivo_excel.exists():
            try:
                df_existente = pd.read_excel(self.arquivo_excel, usecols=['NSU_API'], dtype=str)
                nsus_no_excel = set(df_existente['NSU_API'].dropna().tolist())
            except Exception:
                pass # Se o excel estiver vazio ou sem a coluna ainda
                
        arquivos_xml = list(self.pasta_destino.glob("nota_nsu_*.xml"))
        notas_pendentes = []

        for caminho_xml in arquivos_xml:
            # Extrai o NSU do nome do arquivo (ex: "nota_nsu_125.xml" -> "125")
            nome_arquivo = caminho_xml.stem
            nsu_str = nome_arquivo.replace("nota_nsu_", "")
            
            if nsu_str not in nsus_no_excel:
                try:
                    with open(caminho_xml, 'r', encoding='utf-8') as f:
                        conteudo = f.read()
                    
                    dados_nota = self._converter_xml_para_dict(conteudo, int(nsu_str))
                    if dados_nota:
                        notas_pendentes.append(dados_nota)
                except Exception as e:
                    self.logger.error(f"Erro ao ler XML local {nome_arquivo}: {e}")

        if notas_pendentes:
            self.logger.info(f"Encontrados {len(notas_pendentes)} XML(s) na pasta que não estavam no Excel. Atualizando...")
            self._atualizar_excel(notas_pendentes)
        else:
            self.logger.info("O arquivo Excel já está perfeitamente sincronizado com a pasta de XMLs.")


    def buscar_notas(self):
        # Passo 1: Auto-cura. Varre a pasta para ter certeza de que tudo baixado está no Excel
        self._sincronizar_xmls_locais()
        
        # Passo 2: Continua a busca normal na API
        ult_nsu = self.ler_ultimo_nsu()
        self.logger.info(f"Iniciando busca a partir do NSU: {ult_nsu}")
        
        notas_para_excel = []

        while True:
            url = f"{self.url_base}/{self.cnpj}/NSU/{ult_nsu}"
            
            tentativas = 0
            sucesso = False
            
            while tentativas < 3:
                try:
                    resp = requests.get(url, cert=self.cert_path, timeout=30)
                    
                    if resp.status_code == 200:
                        conteudo = bytearray()
                        for chunk in resp.iter_content(chunk_size=8192):
                            if chunk: conteudo += chunk
                            
                        xml_final = self._descompactar_e_formatar(conteudo.decode('utf-8'))
                        
                        # 1. Salva o XML em arquivo físico
                        with open(self.pasta_destino / f"nota_nsu_{ult_nsu}.xml", "w", encoding="utf-8") as f:
                            f.write(xml_final)
                        
                        # 2. Converte o XML em dicionário para o Excel
                        dados_nota = self._converter_xml_para_dict(xml_final, ult_nsu)
                        if dados_nota:
                            notas_para_excel.append(dados_nota)
                        
                        self.logger.info(f"Nota NSU {ult_nsu} processada e guardada na memória!")
                        ult_nsu += 1
                        self.salvar_ultimo_nsu(ult_nsu)
                        sucesso = True
                        break
                    
                    elif resp.status_code in [404, 204]:
                        self.logger.info("Fim da fila. Nenhuma nota nova encontrada.")
                        self._atualizar_excel(notas_para_excel)
                        return 
                    
                    else:
                        self.logger.error(f"Erro inesperado: {resp.status_code} - {resp.text}")
                        self._atualizar_excel(notas_para_excel)
                        return

                except Exception as e:
                    self.logger.error(f"Falha na conexão: {e}")
                    time.sleep(5)
                    tentativas += 1
            
            if not sucesso:
                self.logger.critical("Falha persistente na comunicação. Robô pausado.")
                self._atualizar_excel(notas_para_excel)
                break
            
            time.sleep(3) # Delay entre cada nota para respeitar o servidor

if __name__ == "__main__":
    SincronizadorGoverno().buscar_notas()