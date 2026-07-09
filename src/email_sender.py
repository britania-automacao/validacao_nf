# src/email_sender.py
import os
import smtplib
from email.message import EmailMessage
from pathlib import Path
from dotenv import load_dotenv

from config import CaminhosProjeto

load_dotenv()

class NotificadorGlobal:
    def __init__(self, caminhos: CaminhosProjeto):
        self.pasta_auditoria = caminhos.saida_auditoria

    def enviar_relatorios_unificados(self):
        print("[EMAIL] Preparando o envio unificado de relatórios...")
        
        remetente = os.getenv('EMAIL_REMETENTE')
        senha = os.getenv('EMAIL_SENHA')
        destinatario = os.getenv('EMAIL_DESTINATARIO')
        smtp_server = os.getenv('SMTP_SERVER')
        smtp_port = int(os.getenv('SMTP_PORT', 587))

        if not all([remetente, senha, destinatario, smtp_server]):
            print("[AVISO] Credenciais de e-mail ausentes no .env. Ignorando envio.")
            return

        if not self.pasta_auditoria.exists():
            print("[AVISO] Pasta de auditoria não existe. Nenhum relatório para enviar.")
            return

        # Busca todos os arquivos Excel, ignorando expressamente qualquer arquivo isolado do Simples Nacional
        arquivos_para_enviar = [
            f for f in self.pasta_auditoria.glob("*.xlsx")
            if "Acao_Necessaria" not in f.name and "Simples_Nacional" not in f.name
        ]

        if not arquivos_para_enviar:
            print("[AVISO] Nenhum arquivo Excel válido encontrado para envio.")
            return

        try:
            msg = EmailMessage()
            msg['Subject'] = '📊 Relatório Diário Unificado: Malha Fina (Serviços e Materiais)'
            msg['From'] = remetente
            msg['To'] = destinatario
            
            texto_email = (
                "Bom dia,\n\n"
                "A auditoria fiscal diária foi concluída com sucesso.\n"
                "Em anexo, você encontrará os Dashboards consolidados e os arquivos auxiliares com as divergências "
                "tanto para as Notas de SERVIÇOS quanto para as Notas de MATERIAIS.\n\n"
                "Com os melhores cumprimentos,\nAutomação Fiscal"
            )
            
            msg.set_content(texto_email)

            # Anexa todos os arquivos encontrados e filtrados
            for caminho in arquivos_para_enviar:
                with open(caminho, 'rb') as f:
                    msg.add_attachment(f.read(), maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=caminho.name)

            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(remetente, senha)
                server.send_message(msg)
                
            print(f"[SUCESSO] E-mail disparado! {len(arquivos_para_enviar)} relatório(s) anexado(s) com sucesso.")
            
        except Exception as e:
            print(f"[ERRO EMAIL] Falha ao enviar o e-mail: {e}")

if __name__ == "__main__":
    RAIZ_PROJETO = Path(__file__).resolve().parent.parent
    config_caminhos = CaminhosProjeto(diretorio_base=RAIZ_PROJETO)
    
    notificador = NotificadorGlobal(caminhos=config_caminhos)
    notificador.enviar_relatorios_unificados()