# src/email_sender.py
import os
import base64
import requests
import urllib3
from pathlib import Path
from dotenv import load_dotenv
from config import CaminhosProjeto

import logging
import sys

# Desativa os avisos de requisição insegura (SSL) no terminal
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- INÍCIO DA CONFIGURAÇÃO DE LOG ---
pasta_historico = Path(__file__).resolve().parent.parent / "dados" / "historico"
pasta_historico.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(message)s',
    handlers=[
        logging.FileHandler(pasta_historico / "email.log", encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

def print(*args, **kwargs):
    sep = kwargs.get("sep", " ")
    logger.info(sep.join(map(str, args)))
# --- FIM DA CONFIGURAÇÃO DE LOG ---

load_dotenv()

class NotificadorGlobal:
    def __init__(self, caminhos: CaminhosProjeto):
        self.pasta_auditoria = caminhos.saida_auditoria

    def _obter_token_graph(self, tenant_id: str, client_id: str, client_secret: str) -> str:
        """Solicita um token de acesso de aplicativo (Client Credentials) para a Microsoft."""
        url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        payload = {
            'grant_type': 'client_credentials',
            'client_id': client_id,
            'client_secret': client_secret,
            'scope': 'https://graph.microsoft.com/.default'
        }
        
        # Inserido verify=False para ignorar o proxy corporativo
        resposta = requests.post(url, data=payload, verify=False)
        resposta.raise_for_status()
        return resposta.json().get('access_token')

    def enviar_relatorios_unificados(self):
        print("[EMAIL] Preparando o envio unificado via Microsoft Graph API...")
        
        remetente = os.getenv('EMAIL_REMETENTE')
        destinatario = os.getenv('EMAIL_DESTINATARIO')
        tenant_id = os.getenv('TENANT_ID')
        client_id = os.getenv('CLIENT_ID')
        client_secret = os.getenv('CLIENT_SECRET')

        if not all([remetente, destinatario, tenant_id, client_id, client_secret]):
            print("[AVISO] Credenciais do Azure/Graph ausentes no .env. Ignorando envio.")
            return

        if not self.pasta_auditoria.exists():
            print("[AVISO] Pasta de auditoria não existe. Nenhum relatório para enviar.")
            return

        arquivos_para_enviar = [
            f for f in self.pasta_auditoria.glob("*.xlsx")
            if "Acao_Necessaria" not in f.name and "Simples_Nacional" not in f.name
        ]

        if not arquivos_para_enviar:
            print("[AVISO] Nenhum arquivo Excel válido encontrado para envio.")
            return

        try:
            # 1. Obter o Token de Acesso
            access_token = self._obter_token_graph(tenant_id, client_id, client_secret)

            # 2. Preparar os anexos em formato Base64
            attachments_json = []
            for caminho in arquivos_para_enviar:
                with open(caminho, 'rb') as f:
                    conteudo_bytes = f.read()
                    conteudo_base64 = base64.b64encode(conteudo_bytes).decode('utf-8')
                    
                    attachments_json.append({
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": caminho.name,
                        "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "contentBytes": conteudo_base64
                    })

            # 3. Tratar múltiplos destinatários (caso seja separado por vírgula ou ponto e vírgula)
            lista_destinatarios = [
                {"emailAddress": {"address": email.strip()}} 
                for email in destinatario.replace(';', ',').split(',') if email.strip()
            ]

            # 4. Montar a estrutura da mensagem (JSON Payload)
            email_payload = {
                "message": {
                    "subject": "📊 Relatório Diário Unificado: Malha Fina (Serviços e Materiais)",
                    "body": {
                        "contentType": "Text",
                        "content": "Bom dia,\n\nA auditoria fiscal diária foi concluída com sucesso.\nEm anexo, você encontrará os Dashboards consolidados e os arquivos auxiliares com as divergências tanto para as Notas de SERVIÇOS quanto para as Notas de MATERIAIS.\n\nCom os melhores cumprimentos,\nAutomação Fiscal"
                    },
                    "toRecipients": lista_destinatarios,
                    "attachments": attachments_json
                },
                "saveToSentItems": "true"
            }

            # 5. Enviar o e-mail chamando a Graph API
            url_send = f"https://graph.microsoft.com/v1.0/users/{remetente}/sendMail"
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json'
            }
            
            # Inserido verify=False para ignorar o proxy corporativo
            resposta_envio = requests.post(url_send, headers=headers, json=email_payload, verify=False)
            resposta_envio.raise_for_status()
                
            print(f"[SUCESSO] E-mail disparado! {len(arquivos_para_enviar)} relatório(s) anexado(s) com sucesso.")
            
        except requests.exceptions.HTTPError as err_http:
            print(f"[ERRO EMAIL] Falha na comunicação com a API: {err_http.response.text}")
        except Exception as e:
            print(f"[ERRO EMAIL] Erro interno ao preparar o envio: {e}")

if __name__ == "__main__":
    RAIZ_PROJETO = Path(__file__).resolve().parent.parent
    config_caminhos = CaminhosProjeto(diretorio_base=RAIZ_PROJETO)
    
    notificador = NotificadorGlobal(caminhos=config_caminhos)
    notificador.enviar_relatorios_unificados()