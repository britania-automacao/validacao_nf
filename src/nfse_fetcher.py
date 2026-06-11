import os
import base64
import gzip
import io
from pathlib import Path
from dotenv import load_dotenv
from requests_pkcs12 import get
from config import ProjectPaths

load_dotenv()

class NationalNfseFetcher:
    """
    Classe responsável por conectar-se à API da Base Nacional da NFS-e (ADN),
    autenticar via Certificado Digital A1 e descarregar os XMLs por NSU.
    """
    def __init__(self, paths: ProjectPaths):
        self.paths = paths
        self.cert_name = os.getenv("CERTIFICADO_NOME")
        self.cert_password = os.getenv("CERTIFICADO_SENHA")
        self.cnpj_empresa = os.getenv("CNPJ_EMPRESA")
        
        # URL Base do Ambiente de Produção Restrita (Testes)
        self.api_base_url = "https://adn.producaorestrita.nfse.gov.br/contribuintes"

    def run_fetch_cycle(self) -> bool:
        """Orquestra o processo de autenticação e download."""
        print(f"[FETCHER] A iniciar sincronização para o CNPJ: {self.cnpj_empresa}...")
        
        caminho_cert = self._get_certificate_path()
        if not caminho_cert:
            return False

        # Para este tentar puxar o NSU número 1 da fila da empresa
        nsu_teste = 1
        sucesso = self._request_national_xml_by_nsu(caminho_cert, nsu_teste)
        return sucesso

    def _get_certificate_path(self):
        """Valida e retorna o caminho absoluto do certificado A1."""
        if not self.cert_name:
            print("[CRITICAL] Nome do certificado não configurado no ficheiro .env")
            return None
            
        cert_path = self.paths.base_dir / "dados" / "certificados" / self.cert_name
        
        if not cert_path.exists():
            print(f"[CRITICAL] Certificado não encontrado no caminho: {cert_path}")
            return None
            
        return str(cert_path)

    def _request_national_xml_by_nsu(self, cert_path: str, nsu: int) -> bool:
        """Executa a chamada HTTP GET utilizando o Certificado Digital A1."""
        url_endpoint = f"{self.api_base_url}/DFe/{nsu}"
        
        # CNPJ de consulta como parâmetro
        parametros = {
            "cnpjConsulta": self.cnpj_empresa
        }

        print(f"[FETCHER] A consultar o NSU {nsu} no portal do Governo...")

        try:
            # Envia o pedido GET assinando com o certificado A1
            response = get(
                url=url_endpoint,
                params=parametros,
                pkcs12_filename=cert_path,
                pkcs12_password=self.cert_password,
                timeout=30
            )

            if response.status_code == 200:
                print(f"[SUCCESS] Documento NSU {nsu} recebido com sucesso!")
                
                # O governo devolve um JSON contendo a lista de documentos. 
                # decodificar e guardar cada um deles.
                dados_json = response.json()
                lote = dados_json.get("LoteDFe", [])
                
                if not lote:
                    print(f"[INFO] Lote vazio recebido para o NSU {nsu}.")
                    return False
                    
                for documento in lote:
                    self._decode_and_save_xml(documento)
                    
                return True
                
            elif response.status_code == 404:
                print(f"[INFO] NSU {nsu} não encontrado ou fila vazia.")
                return False
            elif response.status_code == 403:
                print("[ERRO] Acesso Negado (403). O Certificado não tem permissão para este CNPJ.")
                return False
            else:
                print(f"[ERRO API] O Servidor retornou o código: {response.status_code}")
                print(f"Detalhe: {response.text}")
                return False

        except Exception as e:
            print(f"[CRITICAL] Falha na conexão segura com o Governo: {e}")
            return False

    def _decode_and_save_xml(self, documento: dict):
        """Descompacta o Base64/GZIP do governo e guarda o ficheiro XML em texto limpo."""
        nsu_doc = documento.get("NSU", "Desconhecido")
        chave = documento.get("ChaveAcesso", f"SemChave_{nsu_doc}")
        arquivo_base64 = documento.get("ArquivoXml")

        if not arquivo_base64:
            print(f"[ALERTA] Documento NSU {nsu_doc} não contém o campo ArquivoXml.")
            return

        try:
            # Transforma o Base64 de volta para Bytes comprimidos
            bytes_comprimidos = base64.b64decode(arquivo_base64)
            
            # Descomprime o formato GZIP para obter o XML original
            xml_texto_puro = gzip.decompress(bytes_comprimidos).decode('utf-8')

            # Guarda no disco para o Pandas o poder ler
            caminho_arquivo = self.paths.xml_servicos / f"NFSe_{chave}.xml"
            
            with open(caminho_arquivo, 'w', encoding='utf-8') as f:
                f.write(xml_texto_puro)
                
            print(f"[FETCHER] XML descompactado e guardado: {caminho_arquivo.name}")

        except Exception as e:
            print(f"[ERRO] Falha ao descompactar e guardar o NSU {nsu_doc}: {e}")

if __name__ == "__main__":
    BASE_DIR = Path(__file__).resolve().parent.parent
    caminhos = ProjectPaths(base_dir=BASE_DIR)
    
    fetcher = NationalNfseFetcher(paths=caminhos)
    fetcher.run_fetch_cycle()