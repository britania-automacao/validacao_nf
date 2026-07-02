# src/config.py
from pathlib import Path
from pydantic_settings import BaseSettings, SettingsConfigDict
from pydantic import Field

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

    # National NFSe / Government Configuration
    cnpj_empresa: str = Field(default="76492701000157", validation_alias="CNPJ_EMPRESA")
    cert_pub_path: str | None = Field(default=None, validation_alias="CERT_PUB_PATH")
    cert_key_path: str | None = Field(default=None, validation_alias="CERT_KEY_PATH")

    # Email Configuration
    email_remetente: str | None = Field(default=None, validation_alias="EMAIL_REMETENTE")
    email_senha: str | None = Field(default=None, validation_alias="EMAIL_SENHA")
    email_destinatario: str | None = Field(default=None, validation_alias="EMAIL_DESTINATARIO")
    smtp_server: str | None = Field(default=None, validation_alias="SMTP_SERVER")
    smtp_port: int = Field(default=587, validation_alias="SMTP_PORT")

    # Certificates (for pfx extraction if needed)
    certificado_nome: str = Field(default="certificado.pfx", validation_alias="CERTIFICADO_NOME")
    certificado_senha: str = Field(default="", validation_alias="CERTIFICADO_SENHA")

    # SMB Storage Configuration
    smb_user: str | None = Field(default=None, validation_alias="SMB_USER")
    smb_pass: str | None = Field(default=None, validation_alias="SMB_PASS")


# Singleton instance
settings = ConfigSettings()


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
        
        # Diretórios de entrada para os relatórios agendados do ERP TOTVS
        self.entrada_totvs = self.diretorio_base / "dados" / "entrada_totvs"

        # Diretório de saída para conferência
        self.saida_auditoria = diretorio_base / "dados" / "saida_auditoria"
        
        # Diretórios de entrada para os arquivos XML originais (Fontes de Verdade)
        self.xml_materiais_quirius = self.diretorio_base / "dados" / "xml_materiais_quirius"
        self.xml_servicos_nacional = self.diretorio_base / "dados" / "xml_servicos_nacional"

        # O caminho direto para o seu arquivo mestre com as macros
        self.planilha_mestre = diretorio_base / "dados" / "planilha_mestre.xlsm"
        
        # Diretório para armazenamento de histórico e arquivos de controle (ex: memória do NSU)
        self.historico_controle = self.diretorio_base / "dados" / "historico"
        
        # Caminho físico do Painel de Controle Excel (Planilha Mestre)
        self.planilha_mestre = self.diretorio_base / "planilha_mestre" / "planilha_mestre.xlsm"

    def obter_todos_diretorios_obrigatorios(self) -> list:
        """Retorna a lista com todos os caminhos de pastas necessários para a aplicação."""
        return [
            self.entrada_totvs,
            self.xml_materiais_quirius,
            self.xml_servicos_nacional,
            self.historico_controle
        ]


class InicializadorAmbiente:
    """
    Classe de infraestrutura responsável por validar pré-requisitos,
    criar a árvore de diretórios e garantir a integridade do ambiente antes do processamento.
    """
    def __init__(self, caminhos: CaminhosProjeto):
        self.caminhos = caminhos

    def executar_configuracao_ciclo_vida(self) -> bool:
        """Executa a preparação e validação completa do ambiente local de execução."""
        print("[INFO] Iniciando a validação da infraestrutura local...")
        
        estrutura_pastas_criada = self._criar_arvore_diretorios()
        planilha_mestre_verificada = self._verificar_presenca_planilha_mestre()
        
        if estrutura_pastas_criada and planilha_mestre_verificada:
            print("[SUCESSO] Fase 1 concluída. Ambiente pronto para o processamento.\n")
            return True
        else:
            print("[ERRO] Falha crítica na preparação do ambiente. Verifique as mensagens acima.\n")
            return False

    def _criar_arvore_diretorios(self) -> bool:
        """Varre os caminhos configurados e cria de forma automática as pastas ausentes."""
        try:
            for diretorio in self.caminhos.obter_todos_diretorios_obrigatorios():
                if not diretorio.exists():
                    diretorio.mkdir(parents=True, exist_ok=True)
                    print(f"[SETUP] Diretório criado com sucesso: dados/{diretorio.name}/")
                else:
                    print(f"[INFO] Diretório validado e existente: dados/{diretorio.name}/")
            return True
        except Exception as erro:
            print(f"[CRITICAL] Erro ao instanciar a árvore de diretórios: {erro}")
            return False

    def _verificar_presenca_planilha_mestre(self) -> bool:
        """Garante que a planilha mestre (interface com o usuário) está posicionada corretamente."""
        if self.caminhos.planilha_mestre.exists():
            print(f"[INFO] Interface do usuário identificada: {self.caminhos.planilha_mestre.name}")
            return True
        else:
            print(f"\n[AVISO] ALERTA: Arquivo '{self.caminhos.planilha_mestre.name}' ausente!")
            print(f"[AVISO] Certifique-se de salvá-lo em: {self.caminhos.planilha_mestre.parent}\n")
            return False


if __name__ == "__main__":
    # Captura a raiz do projeto de forma dinâmica (2 níveis acima deste arquivo)
    RAIZ_PROJETO = Path(__file__).resolve().parent.parent
    
    # Instanciação dos objetos com injeção de dependência
    config_caminhos = CaminhosProjeto(diretorio_base=RAIZ_PROJETO)
    inicializador = InicializadorAmbiente(caminhos=config_caminhos)
    
    # Executa o setup inicial
    inicializador.executar_configuracao_ciclo_vida()