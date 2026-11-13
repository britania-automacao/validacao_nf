# src/descobrir_nsu.py
import os
import time
import requests
import urllib3
from pathlib import Path

from config import get_empresa_config

try:
    from cryptography.hazmat.primitives.serialization import pkcs12
    from cryptography.hazmat.primitives import serialization
except ImportError:
    raise ImportError("Instale a biblioteca cryptography: pip install cryptography")

# Oculta os alertas de segurança de conexão HTTPS
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def garantir_certificados_pem(empresa_id: str) -> tuple[str, str]:
    """Garante que as chaves .pem existam para fazer a requisição."""
    cfg = get_empresa_config(empresa_id)
    raiz_projeto = Path(__file__).resolve().parent.parent
    pasta_cert = raiz_projeto / "dados" / "certificados" / empresa_id
    pasta_cert.mkdir(parents=True, exist_ok=True)
    
    cert_pub_path = pasta_cert / "cert_publico.pem"
    cert_key_path = pasta_cert / "chave.pem"
    
    if not cert_pub_path.exists() or not cert_key_path.exists():
        pfx_path = pasta_cert / cfg['cert_nome']
        if not pfx_path.exists():
            raise FileNotFoundError(f"Certificado PFX não encontrado: {pfx_path}")
            
        print("[*] Extraindo chaves do arquivo PFX (Primeira vez)...")
        with open(pfx_path, "rb") as f:
            pfx_data = f.read()
            
        private_key, certificate, _ = pkcs12.load_key_and_certificates(
            pfx_data, cfg['cert_senha'].encode() if cfg['cert_senha'] else None
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

def buscar_ultimo_nsu_regressivo():
    print("="*50)
    print(" BUSCADOR REGRESSIVO DE NSU - RECEITA FEDERAL")
    print("="*50)
    
    empresa_id = input("Digite o ID da Empresa (ex: 1, 15, 30...): ").strip()
    try:
        cfg = get_empresa_config(empresa_id)
        cnpj = cfg['cnpj']
    except Exception as e:
        print(f"\n[ERRO] Empresa não configurada: {e}")
        return

    nsu_str = input("Digite o NSU inicial para começar a descer (ex: 5000): ").strip()
    if not nsu_str.isdigit():
        print("[ERRO] O NSU precisa ser um número inteiro.")
        return
        
    nsu_atual = int(nsu_str)
    
    try:
        caminho_pub, caminho_key = garantir_certificados_pem(empresa_id)
    except Exception as e:
        print(f"[ERRO] Falha com certificados: {e}")
        return
        
    url_base = "https://adn.nfse.gov.br/contribuintes/DFe"
    
    print(f"\nIniciando busca regressiva no CNPJ {cnpj} a partir do NSU {nsu_atual}...\n")

    while nsu_atual > 0:
        url = f"{url_base}/{nsu_atual}"
        params = {'cnpjConsulta': cnpj}
        
        try:
            # Faz a requisição para a API com TLS validado e fallback resiliente
            req_kwargs = dict(
                params=params, 
                cert=(caminho_pub, caminho_key), 
                headers={'Accept': 'application/xml'},
                timeout=15
            )
            try:
                resp = requests.get(url, verify=True, **req_kwargs)
            except requests.exceptions.SSLError:
                resp = requests.get(url, verify=False, **req_kwargs)
            
            if resp.status_code == 200:
                texto_resp = resp.text
                
                # O Governo retorna 137 quando o NSU ainda não foi gerado
                if "<cStat>137</cStat>" in texto_resp:
                    print(f"[-] NSU {nsu_atual}: Não existe (137). Descendo para {nsu_atual - 1}...")
                    nsu_atual -= 1
                    time.sleep(1) # Pausa de 1s para não sobrecarregar a API
                    
                # O Governo retorna 656 quando você atinge o limite de requisições de documentos inexistentes
                elif "<cStat>656</cStat>" in texto_resp:
                    print(f"\n[BLOQUEIO] A Receita retornou Erro 656 no NSU {nsu_atual}.")
                    print("Você atingiu o limite de requisições sem sucesso seguidas.")
                    print("Aguarde alguns minutos antes de tentar novamente.")
                    break
                    
                else:
                    # Se não tem erro 137 nem 656, e o status é 200, achamos uma nota!
                    print("\n" + "="*50)
                    print(f"🎉 BINGO! O ÚLTIMO NSU VÁLIDO ENCONTRADO FOI: {nsu_atual} 🎉")
                    print("="*50)
                    
                    # Salva automaticamente o número no arquivo para facilitar a vida do robô
                    raiz = Path(__file__).resolve().parent.parent
                    arquivo_historico = raiz / "dados" / "historico" / empresa_id / "ultimo_nsu.txt"
                    arquivo_historico.parent.mkdir(parents=True, exist_ok=True)
                    
                    with open(arquivo_historico, "w") as f:
                        f.write(str(nsu_atual))
                        
                    print(f"O arquivo 'ultimo_nsu.txt' da Empresa {empresa_id} já foi atualizado para {nsu_atual}!")
                    break
                    
            # NOVA LÓGICA: Se retornar 404 (Not Found) ou 204 (No Content), significa que também não existe.
            elif resp.status_code in [404, 204]:
                print(f"[-] NSU {nsu_atual}: Não existe (Status {resp.status_code}). Descendo para {nsu_atual - 1}...")
                nsu_atual -= 1
                time.sleep(1)
                
            elif resp.status_code == 429:
                print(f"[!] Limite de taxa (429) atingido no NSU {nsu_atual}. Aguardando 10 segundos...")
                time.sleep(10)
                
            else:
                print(f"[!] Erro HTTP {resp.status_code} no NSU {nsu_atual}. Tentando novamente...")
                time.sleep(3)
                
        except requests.exceptions.RequestException as e:
            print(f"[!] Falha na conexão: {e}. Tentando novamente em 3 segundos...")
            time.sleep(3)

if __name__ == "__main__":
    buscar_ultimo_nsu_regressivo()