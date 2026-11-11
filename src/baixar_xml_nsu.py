# src/baixar_xml_nsu.py
import requests
import base64
import gzip
import xml.etree.ElementTree as ET
import xml.dom.minidom
from pathlib import Path
import urllib3

from config import get_empresa_config

try:
    from cryptography.hazmat.primitives.serialization import pkcs12
    from cryptography.hazmat.primitives import serialization
except ImportError:
    raise ImportError("A biblioteca 'cryptography' é necessária para converter PFX. Execute: pip install cryptography")

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


class BaixadorXML:
    def __init__(self, empresa_id):
        self.empresa_id = str(empresa_id)
        cfg = get_empresa_config(self.empresa_id)
        self.cnpj = cfg['cnpj']
        self.url_base = "https://adn.nfse.gov.br/contribuintes/DFe" 
        
        self.raiz_projeto = Path(__file__).resolve().parent.parent
        
        # Apenas a pasta de destino do XML
        self.pasta_destino = self.raiz_projeto / "dados" / "xml_servicos_nacional" / self.empresa_id
        self.pasta_destino.mkdir(parents=True, exist_ok=True)
        
        # Configuração de Certificados
        pasta_cert = self.raiz_projeto / "dados" / "certificados" / self.empresa_id
        pasta_cert.mkdir(parents=True, exist_ok=True)
        self.cert_path = self._garantir_certificados_pem(pasta_cert, cfg['cert_nome'], cfg['cert_senha'])

    def _garantir_certificados_pem(self, pasta_cert: Path, cert_nome: str, cert_senha: str):
        cert_pub_path = pasta_cert / "cert_publico.pem"
        cert_key_path = pasta_cert / "chave.pem"
        
        if not cert_pub_path.exists() or not cert_key_path.exists():
            pfx_path = pasta_cert / cert_nome
            if not pfx_path.exists():
                raise FileNotFoundError(f"Arquivo de certificado PFX não encontrado: {pfx_path}")
            
            print(f"Extraindo chaves .pem do certificado PFX {cert_nome}...")
            with open(pfx_path, "rb") as f:
                pfx_data = f.read()
                
            private_key, certificate, _ = pkcs12.load_key_and_certificates(
                pfx_data, cert_senha.encode() if cert_senha else None
            )
            
            with open(cert_key_path, "wb") as f:
                f.write(private_key.private_bytes(
                    encoding=serialization.Encoding.PEM,
                    format=serialization.PrivateFormat.TraditionalOpenSSL,
                    encryption_algorithm=serialization.NoEncryption()
                ))
                
            with open(cert_pub_path, "wb") as f:
                f.write(certificate.public_bytes(serialization.Encoding.PEM))
                
        return (str(cert_pub_path), str(cert_key_path))

    def _descompactar_e_formatar(self, conteudo_resposta: str) -> tuple[bool, str]:
        try:
            try:
                raiz = ET.fromstring(conteudo_resposta)
            except ET.ParseError:
                return False, "XML Inválido ou Malformado"

            for elem in raiz.iter():
                if elem.tag.endswith('cStat'):
                    if elem.text == '137': return False, "137 - Nenhum documento localizado"
                    if elem.text == '656': return False, "656 - Consumo Indevido (Aguarde alguns minutos)"

            base64_str = None
            for elem in raiz.iter():
                if elem.tag.endswith('docZip') or elem.tag.endswith('nfseZip') or elem.tag.endswith('ArquivoXml'):
                    base64_str = elem.text
                    break
            
            if base64_str:
                dados_binarios = base64.b64decode(base64_str)
                xml_bruto = gzip.decompress(dados_binarios).decode('utf-8')
                xml_final = xml_bruto
            else:
                xml_final = conteudo_resposta
            
            dom = xml.dom.minidom.parseString(xml_final)
            xml_formatado = dom.toprettyxml(indent="  ")
            linhas = [linha for linha in xml_formatado.split('\n') if linha.strip()]
            return True, '\n'.join(linhas)
            
        except Exception as e:
            if "<" in conteudo_resposta and ">" in conteudo_resposta:
                 return True, conteudo_resposta
            return False, f"ERRO_GERAL: {e}"

    def baixar(self, nsu_alvo: int):
        print(f"\nConectando na Receita Federal para buscar NSU: {nsu_alvo} (Empresa {self.empresa_id})...")
        
        url = f"{self.url_base}/{nsu_alvo}"
        
        try:
            resp = requests.get(
                url, 
                params={'cnpjConsulta': self.cnpj}, 
                headers={'Accept': 'application/xml'}, 
                cert=self.cert_path, 
                timeout=30, 
                stream=True
            )
            
            if resp.status_code == 200:
                conteudo = bytearray()
                for chunk in resp.iter_content(chunk_size=8192):
                    if chunk: conteudo += chunk
                    
                sucesso_descomp, resultado = self._descompactar_e_formatar(conteudo.decode('utf-8'))
                
                if not sucesso_descomp:
                    print(f"[X] Falha ou aviso da Receita: {resultado}")
                    return
                
                xml_final = resultado
                caminho_xml = self.pasta_destino / f"nota_nsu_{nsu_alvo}.xml"
                
                with open(caminho_xml, "w", encoding="utf-8") as f:
                    f.write(xml_final)
                print(f"[+] SUCESSO! XML baixado e salvo em: {caminho_xml}")
                
            elif resp.status_code in [404, 204]:
                print(f"[X] NSU {nsu_alvo} não existe na Receita Federal (Status {resp.status_code}).")
            elif resp.status_code == 429:
                print(f"[X] Você está bloqueado por excesso de requisições (Status 429). Tente novamente em alguns minutos.")
            else:
                print(f"[X] Erro HTTP desconhecido: {resp.status_code}")

        except Exception as e:
            print(f"[X] Falha na conexão ou processamento: {e}")


def main():
    print("="*50)
    print(" DOWNLOAD DIRETO DE XML POR NSU")
    print("="*50)
    
    # ========================================================
    # ⚠️ INSIRA OS DADOS MANUALMENTE AQUI ANTES DE RODAR ⚠️
    # ========================================================
    EMPRESA_ID = "15"        # Coloque o ID da Empresa (ex: "1", "15", "30", "35")
    NSU_ALVO = "5086"       # Coloque o número do NSU desejado
    # ========================================================
    
    if not EMPRESA_ID or not NSU_ALVO.isdigit():
        print("Valores inválidos. Verifique as variáveis EMPRESA_ID e NSU_ALVO no código.")
        return
        
    try:
        baixador = BaixadorXML(empresa_id=EMPRESA_ID)
        baixador.baixar(int(NSU_ALVO))
    except Exception as e:
        print(f"\nErro Crítico: {e}")
        
    print("\nProcesso finalizado.")


if __name__ == "__main__":
    main()