# src/nfse_fetcher.py
import os
import base64
import gzip
from pathlib import Path
from dotenv import load_dotenv
from requests_pkcs12 import get
from config import CaminhosProjeto

load_dotenv()

class BuscadorNfseNacional:
    """
    Classe responsável por conectar-se à API do Ambiente de Dados Nacional (ADN),
    autenticar via Certificado Digital A1 e baixar os arquivos XML de forma sequencial por NSU.
    """
    def __init__(self, caminhos: CaminhosProjeto):
        self.caminhos = caminhos
        self.nome_certificado = os.getenv("CERTIFICADO_NOME")
        self.senha_certificado = os.getenv("CERTIFICADO_SENHA")
        self.cnpj_empresa = os.getenv("CNPJ_EMPRESA")
        
        # URL Base do Ambiente de Produção Restrita (Testes), conforme o manual oficial da API
        self.url_base_api = "https://adn.producaorestrita.nfse.gov.br/contribuintes"

    def executar_ciclo_busca(self) -> bool:
        """Orquestra o download contínuo dos XMLs consumindo a fila até encontrar o fim dela (Erro 404)."""
        print(f"[FETCHER] Iniciando a sincronização de notas para o CNPJ: {self.cnpj_empresa}...")
        
        caminho_absoluto_certificado = self._obter_caminho_certificado()
        if not caminho_absoluto_certificado:
            return False

        # Define o arquivo de texto que serve como memória de progresso
        arquivo_memoria_nsu = self.caminhos.historico_controle / "ultimo_nsu.txt"
        
        # Recupera o último NSU processado ou inicia do primeiro se o arquivo não existir
        if arquivo_memoria_nsu.exists():
            nsu_atual = int(arquivo_memoria_nsu.read_text().strip())
        else:
            nsu_atual = 1

        print(f"[FETCHER] Retomando os downloads a partir do NSU: {nsu_atual}...")
        total_notas_baixadas = 0

        # Loop de consumo da fila que só encerra quando o governo retornar 404 (fim da fila)
        while True:
            sucesso_requisicao = self._requisitar_xml_por_nsu(caminho_absoluto_certificado, nsu_atual)
            
            if sucesso_requisicao:
                total_notas_baixadas += 1
                nsu_atual += 1
                # Atualiza a memória do robô no disco imediatamente para evitar perda de tracking
                arquivo_memoria_nsu.write_text(str(nsu_atual))
            else:
                print(f"[FETCHER] Fila de transmissão esgotada. Nenhuma nota nova encontrada após o NSU {nsu_atual - 1}.")
                break
                
        print(f"[SUCESSO] Ciclo finalizado! Total de novos lotes salvos nesta sessão: {total_notas_baixadas}\n")
        return True

    def _obter_caminho_certificado(self) -> str:
        """Valida as configurações do ambiente e retorna o caminho físico do certificado A1."""
        if not self.nome_certificado:
            print("[CRITICAL] Nome do certificado não configurado no arquivo .env")
            return None
            
        caminho_certificado = self.caminhos.diretorio_base / "dados" / "certificados" / self.nome_certificado
        
        if not caminho_certificado.exists():
            print(f"[CRITICAL] Arquivo do certificado não encontrado em: {caminho_certificado}")
            return None
            
        return str(caminho_certificado)

    def _requisitar_xml_por_nsu(self, caminho_cert: str, nsu_consulta: int) -> bool:
        """Executa a chamada HTTP GET assinando digitalmente a requisição com o certificado PFX."""
        url_endpoint_completo = f"{self.url_base_api}/DFe/{nsu_consulta}"
        
        parametros_consulta = {
            "cnpjConsulta": self.cnpj_empresa
        }

        print(f"[FETCHER] Consultando a posição de NSU {nsu_consulta} na Base Nacional...")

        try:
            resposta_servidor = get(
                url=url_endpoint_completo,
                params=parametros_consulta,
                pkcs12_filename=caminho_cert,
                pkcs12_password=self.senha_certificado,
                timeout=30
            )

            # Status 200: Encontrou dados na fila
            if resposta_servidor.status_code == 200:
                dados_resposta_json = resposta_servidor.json()
                lista_documentos_lote = dados_resposta_json.get("LoteDFe", [])
                
                if not lista_documentos_lote:
                    return False
                    
                for documento_fiscal in lista_documentos_lote:
                    self._decodificar_e_salvar_arquivo_xml(documento_fiscal)
                    
                return True
                
            # Status 404: Fim da fila ou índice não gerado ainda pelo fisco
            elif resposta_servidor.status_code == 404:
                print(f"[INFO] NSU {nsu_consulta} não localizado (Fim da fila alcançado).")
                return False
            elif resposta_servidor.status_code == 403:
                print("[ERRO] Acesso Negado (403). O certificado não possui permissão para este CNPJ.")
                return False
            else:
                print(f"[ERRO API] O servidor retornou o status {resposta_servidor.status_code}")
                return False

        except Exception as erro_conexao:
            print(f"[CRITICAL] Falha na conexão segura SSL com o ambiente governamental: {erro_conexao}")
            return False

    def _decodificar_e_salvar_arquivo_xml(self, dados_documento: dict):
        """Descompacta a string criptografada em Base64/GZIP gerando o arquivo XML legível em UTF-8."""
        nsu_documento = dados_documento.get("NSU", "Desconhecido")
        chave_acesso = dados_documento.get("ChaveAcesso", f"SemChave_{nsu_documento}")
        string_base64_comprimida = dados_documento.get("ArquivoXml")

        if not string_base64_comprimida:
            print(f"[AVISO] O documento de NSU {nsu_documento} veio com o bloco XML vazio.")
            return

        try:
            # Decodifica a base64 transformando em bytes binários compactados
            bytes_comprimidos = base64.b64decode(string_base64_comprimida)
            
            # Descomprime o GZIP recuperando a string estruturada do XML original
            texto_xml_puro = gzip.decompress(bytes_comprimidos).decode('utf-8')

            # Grava o arquivo de texto limpo no diretório mapeado pelo ecossistema
            caminho_final_arquivo = self.caminhos.xml_servicos_nacional / f"NFSe_{chave_acesso}.xml"
            
            with open(caminho_final_arquivo, 'w', encoding='utf-8') as arquivo_escrita:
                arquivo_escrita.write(texto_xml_puro)
                
            print(f"[FETCHER] Arquivo gravado com sucesso: {caminho_final_arquivo.name}")

        except Exception as erro_decodificacao:
            print(f"[ERRO] Falha ao processar o desmembramento criptográfico do NSU {nsu_documento}: {erro_decodificacao}")


if __name__ == "__main__":
    RAIZ_PROJETO = Path(__file__).resolve().parent.parent
    config_caminhos = CaminhosProjeto(diretorio_base=RAIZ_PROJETO)
    
    gerenciador_fetcher = BuscadorNfseNacional(caminhos=config_caminhos)
    gerenciador_fetcher.executar_ciclo_busca()