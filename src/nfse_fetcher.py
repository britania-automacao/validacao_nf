# src/nfse_fetcher.py
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
import urllib3

from config import get_empresa_config

try:
    from cryptography.hazmat.primitives.serialization import pkcs12
    from cryptography.hazmat.primitives import serialization
except ImportError:
    raise ImportError("A biblioteca 'cryptography' é necessária para converter PFX. Execute: pip install cryptography")

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


class SincronizadorGoverno:
    def __init__(self, empresa_id):
        self.empresa_id = str(empresa_id)
        cfg = get_empresa_config(self.empresa_id)
        self.cnpj = cfg['cnpj']
        self.url_base = "https://adn.nfse.gov.br/contribuintes/DFe" 
        
        raiz_projeto = Path(__file__).resolve().parent.parent
        
        # 1. Configurar Pastas Base e Criar o LOGGER primeiro!
        self.pasta_historico = raiz_projeto / "dados" / "historico" / self.empresa_id
        self.pasta_historico.mkdir(parents=True, exist_ok=True)
        
        self.logger = logging.getLogger(f"Sincronizador_{self.empresa_id}")
        self.logger.setLevel(logging.INFO)
        if self.logger.hasHandlers():
            self.logger.handlers.clear()
            
        fmt = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        fh = logging.FileHandler(self.pasta_historico / "sincronizador.log", encoding='utf-8')
        sh = logging.StreamHandler()
        fh.setFormatter(fmt)
        sh.setFormatter(fmt)
        self.logger.addHandler(fh)
        self.logger.addHandler(sh)

        # 2. Configuração de Certificados (Agora o logger já existe para ser usado)
        pasta_cert = raiz_projeto / "dados" / "certificados" / self.empresa_id
        pasta_cert.mkdir(parents=True, exist_ok=True)
        self.cert_path = self._garantir_certificados_pem(pasta_cert, cfg['cert_nome'], cfg['cert_senha'])
        
        # 3. Configurar restante dos arquivos locais
        self.arquivo_nsu = self.pasta_historico / "ultimo_nsu.txt"
        
        self.pasta_destino = raiz_projeto / "dados" / "xml_servicos_nacional" / self.empresa_id
        self.pasta_destino.mkdir(parents=True, exist_ok=True)
        
        self.arquivo_excel = self.pasta_destino / f"Resumo_Notas_Nacional_{self.empresa_id}.xlsx"
        
        # 4. Configurar caminho para CÓPIA NA REDE
        # Aponta para P:\FISCAL\GERAL\05 - KPIs RECEBIMENTO FISCAL\8. Automacao Validacao NF\xml_servicos_nacional
        self.pasta_rede = Path(r"P:\FISCAL\GERAL\05 - KPIs RECEBIMENTO FISCAL\8. Automacao Validacao NF\xml_servicos_nacional") / self.empresa_id
        self.arquivo_excel_rede = self.pasta_rede / f"Resumo_Notas_Nacional_{self.empresa_id}.xlsx"

    def _garantir_certificados_pem(self, pasta_cert: Path, cert_nome: str, cert_senha: str):
        """Transforma o certificado .pfx em arquivos .pem para a API em Python"""
        cert_pub_path = pasta_cert / "cert_publico.pem"
        cert_key_path = pasta_cert / "chave.pem"
        
        if not cert_pub_path.exists() or not cert_key_path.exists():
            pfx_path = pasta_cert / cert_nome
            if not pfx_path.exists():
                raise FileNotFoundError(f"Arquivo de certificado PFX não encontrado: {pfx_path}")
            
            self.logger.info(f"Extraindo chaves .pem do certificado PFX {cert_nome}...")
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

    def _descompactar_e_formatar(self, conteudo_resposta: str) -> tuple[bool, str]:
        try:
            try:
                raiz = ET.fromstring(conteudo_resposta)
            except ET.ParseError:
                return False, "XML_INVALIDO"

            for elem in raiz.iter():
                if elem.tag.endswith('cStat'):
                    if elem.text == '137': return False, "137"
                    if elem.text == '656': return False, "656"

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
            
    def _converter_xml_para_dict(self, xml_string: str, nsu: int) -> dict:
        try:
            raiz = ET.fromstring(xml_string)
            dados = {'NSU_API': str(nsu)}

            def achatar(nodo, prefixo=''):
                for filho in nodo:
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
        if not notas_extraidas: return
            
        try:
            df_novos = pd.DataFrame(notas_extraidas)
            
            if self.arquivo_excel.exists():
                df_existente = pd.read_excel(self.arquivo_excel, dtype=str)
                df_final = pd.concat([df_existente, df_novos], ignore_index=True)
                if 'NSU_API' in df_final.columns:
                    df_final = df_final.drop_duplicates(subset=['NSU_API'], keep='last')
            else:
                df_final = df_novos
                
            # Salva na pasta LOCAL (C:)
            df_final.to_excel(self.arquivo_excel, index=False)
            
            # Garante que a pasta na REDE existe e salva a CÓPIA nela
            try:
                self.pasta_rede.mkdir(parents=True, exist_ok=True)
                df_final.to_excel(self.arquivo_excel_rede, index=False)
                self.logger.info(f"Sucesso! {len(notas_extraidas)} nota(s) salva(s) no Resumo Local e na Cópia da Rede.")
            except Exception as e_rede:
                self.logger.warning(f"Resumo Local salvo, mas falha ao salvar cópia na Rede (P:): {e_rede}")

        except Exception as e:
            self.logger.error(f"Falha ao salvar os dados no Excel: {e}")

    def _sincronizar_xmls_locais(self):
        self.logger.info("Verificando se há XMLs locais pendentes...")
        nsus_no_excel = set()
        if self.arquivo_excel.exists():
            try:
                df_existente = pd.read_excel(self.arquivo_excel, usecols=['NSU_API'], dtype=str)
                nsus_no_excel = set(df_existente['NSU_API'].dropna().tolist())
            except Exception: pass
                
        arquivos_xml = list(self.pasta_destino.glob("nota_nsu_*.xml"))
        notas_pendentes = []

        for caminho_xml in arquivos_xml:
            nsu_str = caminho_xml.stem.replace("nota_nsu_", "")
            if nsu_str not in nsus_no_excel:
                try:
                    with open(caminho_xml, 'r', encoding='utf-8') as f:
                        conteudo = f.read()
                    dados_nota = self._converter_xml_para_dict(conteudo, int(nsu_str))
                    if dados_nota: notas_pendentes.append(dados_nota)
                except Exception: pass

        if notas_pendentes:
            self.logger.info(f"Atualizando Excel com {len(notas_pendentes)} XML(s) que já existiam na pasta...")
            self._atualizar_excel(notas_pendentes)

    def buscar_notas(self):
        self._sincronizar_xmls_locais()
        ult_nsu = self.ler_ultimo_nsu()
        if ult_nsu == 0: ult_nsu = 1
        
        self.logger.info(f"Iniciando busca no Governo a partir do NSU: {ult_nsu}")
        notas_para_excel = []

        while True:
            url = f"{self.url_base}/{ult_nsu}"
            tentativas = 0
            sucesso = False
            
            while tentativas < 3:
                try:
                    resp = requests.get(url, params={'cnpjConsulta': self.cnpj}, headers={'Accept': 'application/xml'}, cert=self.cert_path, timeout=30, stream=True)
                    
                    if resp.status_code == 200:
                        conteudo = bytearray()
                        for chunk in resp.iter_content(chunk_size=8192):
                            if chunk: conteudo += chunk
                            
                        sucesso_descomp, resultado = self._descompactar_e_formatar(conteudo.decode('utf-8'))
                        
                        if not sucesso_descomp:
                            if resultado == "137":
                                self.logger.info("Fim da fila (cStat 137). Nenhuma nota nova.")
                                self._atualizar_excel(notas_para_excel)
                                return 
                            elif resultado == "656":
                                self.logger.warning("Bloqueio 656. O robô vai parar.")
                                self._atualizar_excel(notas_para_excel)
                                return 
                            else:
                                self.logger.error(f"Erro crítico no NSU {ult_nsu}: {resultado}")
                                return 
                        
                        xml_final = resultado
                        with open(self.pasta_destino / f"nota_nsu_{ult_nsu}.xml", "w", encoding="utf-8") as f:
                            f.write(xml_final)
                        
                        dados_nota = self._converter_xml_para_dict(xml_final, ult_nsu)
                        if dados_nota: notas_para_excel.append(dados_nota)
                        
                        self.logger.info(f"Nota NSU {ult_nsu} processada!")
                        ult_nsu += 1
                        self.salvar_ultimo_nsu(ult_nsu)
                        sucesso = True
                        break
                    
                    elif resp.status_code in [404, 204]:
                        self.logger.info("Fim da fila (Status HTTP).")
                        self._atualizar_excel(notas_para_excel)
                        return 
                    
                    elif resp.status_code == 429:
                        time.sleep(30 * (tentativas + 1))
                        tentativas += 1
                        continue
                    else:
                        self.logger.error(f"Erro HTTP {resp.status_code}")
                        return

                except Exception as e:
                    self.logger.error(f"Falha na conexão: {e}")
                    time.sleep(5)
                    tentativas += 1
            
            if not sucesso: break
            time.sleep(3)


if __name__ == "__main__":
    empresas = ["1", "15", "30", "35"]
    
    for emp in empresas:
        print(f"\n{'='*50}")
        print(f" Iniciando sincronização do Governo - Empresa {emp}")
        print(f"{'='*50}")
        try:
            SincronizadorGoverno(empresa_id=emp).buscar_notas()
        except Exception as e:
            print(f"[ERRO CRÍTICO] Falha ao processar a Empresa {emp}: {e}")