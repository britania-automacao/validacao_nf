# src/config.py
import os
from pathlib import Path
from pydantic_settings import BaseSettings, SettingsConfigDict
from pydantic import Field
from dotenv import load_dotenv

# Força o carregamento do .env
load_dotenv(Path(__file__).resolve().parent.parent / '.env', override=True)

class ConfigSettings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=str(Path(__file__).resolve().parent.parent / '.env'),
        env_file_encoding='utf-8',
        extra='ignore'
    )

    # Quirius Configuration
    quirius_user: str = Field(default="", validation_alias="QUIRIUS_USER")
    quirius_pass: str = Field(default="", validation_alias="QUIRIUS_PASS")
    quirius_url: str = Field(default="http://10.0.10.57:600", validation_alias="QUIRIUS_URL")

    # Email Configuration
    email_remetente: str | None = Field(default=None, validation_alias="EMAIL_REMETENTE")
    email_senha: str | None = Field(default=None, validation_alias="EMAIL_SENHA")
    email_destinatario: str | None = Field(default=None, validation_alias="EMAIL_DESTINATARIO")
    smtp_server: str | None = Field(default=None, validation_alias="SMTP_SERVER")
    smtp_port: int = Field(default=587, validation_alias="SMTP_PORT")

    # SMB Storage Configuration
    smb_user: str | None = Field(default=None, validation_alias="SMB_USER")
    smb_pass: str | None = Field(default=None, validation_alias="SMB_PASS")

settings = ConfigSettings()

def get_empresa_config(empresa_id: str) -> dict:
    """Busca as credenciais de uma empresa específica no .env."""
    cnpj = os.getenv(f"EMPRESA_{empresa_id}_CNPJ")
    senha = os.getenv(f"EMPRESA_{empresa_id}_CERT_SENHA")
    nome = os.getenv(f"EMPRESA_{empresa_id}_CERT_NOME")
    
    if not cnpj or not nome:
        raise ValueError(f"Configurações incompletas no .env para a Empresa {empresa_id}.")
        
    return {
        "cnpj": cnpj,
        "cert_senha": senha,
        "cert_nome": nome
    }

def get_storage():
    from unistorage import LocalStorage, SMBStorage
    from unistorage.router import StorageRouter
    
    local = LocalStorage()
    smb = None
    if settings.smb_user and settings.smb_pass:
        smb = SMBStorage(username=settings.smb_user, password=settings.smb_pass)
    return StorageRouter(local=local, smb=smb)

class CaminhosProjeto:
    """
    Classe responsável por centralizar, organizar e gerenciar todos os caminhos
    de diretórios e arquivos do ecossistema de validação fiscal.
    """
    def __init__(self, diretorio_base: Path):
        self.diretorio_base = diretorio_base
        
        self.entrada_totvs = self.diretorio_base / "dados" / "entrada_totvs"
        self.saida_auditoria = diretorio_base / "dados" / "saida_auditoria"
        self.xml_materiais_quirius = self.diretorio_base / "dados" / "xml_materiais_quirius"
        self.xml_servicos_nacional = self.diretorio_base / "dados" / "xml_servicos_nacional"
        self.certificados = self.diretorio_base / "dados" / "certificados"
        
        self.historico_controle = self.diretorio_base / "dados" / "historico"
        self.planilha_mestre = self.diretorio_base / "planilha_mestre" / "planilha_mestre.xlsm"

    def obter_todos_diretorios_obrigatorios(self) -> list:
        return [
            self.entrada_totvs,
            self.xml_materiais_quirius,
            self.xml_servicos_nacional,
            self.certificados,
            self.historico_controle
        ]

class InicializadorAmbiente:
    def __init__(self, caminhos: CaminhosProjeto):
        self.caminhos = caminhos

    def executar_configuracao_ciclo_vida(self) -> bool:
        print("[INFO] Iniciando a validação da infraestrutura local...")
        estrutura_pastas_criada = self._criar_arvore_diretorios()
        planilha_mestre_verificada = self._verificar_presenca_planilha_mestre()
        
        if estrutura_pastas_criada and planilha_mestre_verificada:
            print("[SUCESSO] Fase 1 concluída. Ambiente pronto para o processamento.\n")
            return True
        else:
            print("[ERRO] Falha crítica na preparação do ambiente.\n")
            return False

    def _criar_arvore_diretorios(self) -> bool:
        try:
            for diretorio in self.caminhos.obter_todos_diretorios_obrigatorios():
                if not diretorio.exists():
                    diretorio.mkdir(parents=True, exist_ok=True)
            return True
        except Exception as erro:
            print(f"[CRITICAL] Erro ao instanciar a árvore de diretórios: {erro}")
            return False

    def _verificar_presenca_planilha_mestre(self) -> bool:
        if self.caminhos.planilha_mestre.exists():
            return True
        else:
            print(f"\n[AVISO] ALERTA: Arquivo '{self.caminhos.planilha_mestre.name}' ausente!")
            return False