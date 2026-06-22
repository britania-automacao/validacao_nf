import os
import time
import requests
import base64
import gzip
import xml.etree.ElementTree as ET
import xml.dom.minidom
import logging
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

class SincronizadorGoverno:
    def __init__(self):
        self.cnpj = os.getenv('CNPJ_EMPRESA', '76492701000157')
        self.url_base = "https://adn.nfse.gov.br/contribuintes/DFe" 
        self.cert_path = (os.getenv('CERT_PUB_PATH'), os.getenv('CERT_KEY_PATH'))
        
        # Configuração de Pastas e Caminhos
        raiz_projeto = Path(__file__).resolve().parent.parent
        self.pasta_historico = raiz_projeto / "dados" / "historico"
        self.pasta_historico.mkdir(parents=True, exist_ok=True)
        self.arquivo_nsu = self.pasta_historico / "ultimo_nsu.txt"
        
        self.pasta_destino = raiz_projeto / "dados" / "xml_servicos_nacional"
        self.pasta_destino.mkdir(parents=True, exist_ok=True)
        
        # Configuração do Logger
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(self.pasta_historico / "fetcher.log", encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)

    def ler_ultimo_nsu(self) -> int:
        if self.arquivo_nsu.exists():
            with open(self.arquivo_nsu, "r") as f:
                try: return int(f.read().strip())
                except ValueError: return 0
        return 0

    def salvar_ultimo_nsu(self, nsu: int):
        with open(self.arquivo_nsu, "w") as f:
            f.write(str(nsu))

    def _descompactar_e_formatar(self, xml_resposta: str) -> str:
        try:
            raiz_envelope = ET.fromstring(xml_resposta)
            nodo_arquivo = next((e for e in raiz_envelope.iter() if 'ArquivoXml' in e.tag), None)
            
            if nodo_arquivo is not None and nodo_arquivo.text:
                dados_zipados = base64.b64decode(nodo_arquivo.text)
                xml_verdadeiro = gzip.decompress(dados_zipados).decode('utf-8')
                
                # Embelezador de XML (formatação visual)
                dom = xml.dom.minidom.parseString(xml_verdadeiro.replace("\n", "").replace("\r", ""))
                return dom.toprettyxml(indent="    ")
            return xml_resposta
        except Exception as e:
            self.logger.error(f"Falha ao processar/formatar XML: {e}")
            return xml_resposta

    def buscar_notas(self):
        ult_nsu = self.ler_ultimo_nsu()
        if ult_nsu == 0: ult_nsu = 1
            
        self.logger.info(f"Sincronização iniciada. CNPJ: {self.cnpj} | NSU Inicial: {ult_nsu}")

        while True:
            url_chamada = f"{self.url_base}/{ult_nsu}"
            tentativas = 0
            sucesso = False
            
            # Lógica de Retentativa para erros de rede/servidor (502, 503, 504 e 429)
            while tentativas < 3:
                try:
                    resp = requests.get(
                        url_chamada,
                        params={'cnpjConsulta': self.cnpj},
                        headers={'Accept': 'application/xml'},
                        cert=self.cert_path,
                        timeout=90,
                        stream=True # Evita erros de IncompleteRead
                    )
                    
                    # Trata limite de requisições
                    if resp.status_code == 429:
                        wait = 30 * (tentativas + 1)
                        self.logger.warning(f"Limite excedido (429). Pausando por {wait}s...")
                        time.sleep(wait)
                        tentativas += 1
                        continue

                    # Trata instabilidade no servidor
                    if resp.status_code in [502, 503, 504]:
                        tentativas += 1
                        self.logger.warning(f"Instabilidade no Governo (Erro {resp.status_code}). Tentativa {tentativas}/3...")
                        time.sleep(10 * tentativas)
                        continue
                    
                    if resp.status_code == 200:
                        # Leitura em pedaços (streaming)
                        conteudo = b""
                        for chunk in resp.iter_content(chunk_size=8192):
                            if chunk: conteudo += chunk
                            
                        xml_final = self._descompactar_e_formatar(conteudo.decode('utf-8'))
                        
                        with open(self.pasta_destino / f"nota_nsu_{ult_nsu}.xml", "w", encoding="utf-8") as f:
                            f.write(xml_final)
                        
                        self.logger.info(f"Nota NSU {ult_nsu} processada!")
                        ult_nsu += 1
                        self.salvar_ultimo_nsu(ult_nsu)
                        sucesso = True
                        break
                    
                    elif resp.status_code in [404, 204]:
                        self.logger.info("Fim da fila. Nenhuma nota nova encontrada.")
                        return 
                    
                    else:
                        self.logger.error(f"Erro inesperado: {resp.status_code} - {resp.text}")
                        return

                except Exception as e:
                    self.logger.error(f"Falha na conexão: {e}")
                    time.sleep(5)
                    tentativas += 1
            
            if not sucesso:
                self.logger.critical("Falha persistente na comunicação. Robô pausado.")
                break
            
            time.sleep(3) # Delay entre cada nota para respeitar o servidor

if __name__ == "__main__":
    SincronizadorGoverno().buscar_notas()