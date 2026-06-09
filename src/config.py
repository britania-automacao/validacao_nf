# src/config.py
from pathlib import Path

class ProjectPaths:
    """
    Classe responsável por encapsular, centralizar e gerenciar todos os caminhos
    de diretórios e arquivos do ecossistema de validação fiscal.
    """
    def __init__(self, base_dir: Path):
        self.base_dir = base_dir
        
        # Diretórios de entrada para os relatórios agendados do ERP TOTVS
        self.entrada_totvs = self.base_dir / "dados" / "entrada_totvs"
        
        # Diretórios de entrada para os XMLs originais (Fontes de Verdade)
        self.xml_materiais = self.base_dir / "dados" / "xml_materiais_quirius"
        self.xml_servicos = self.base_dir / "dados" / "xml_servicos_nacional"
        
        # Diretório para armazenamento de logs e histórico de processamento
        self.historico = self.base_dir / "dados" / "historico"
        
        # Caminho físico do Painel de Controle Excel (Planilha Mestre)
        self.planilha_mestre = self.base_dir / "planilha_mestre" / "Analise de notas de serviços_V03.07.2025.xlsm"

    def get_all_required_directories(self) -> list:
        """Retorna a lista de todas as pastas que o sistema precisa para funcionar."""
        return [
            self.entrada_totvs,
            self.xml_materiais,
            self.xml_servicos,
            self.historico
        ]


class EnvironmentInitializer:
    """
    Classe de infraestrutura responsável por validar pré-requisitos,
    criar árvores de diretórios e garantir a integridade do ambiente antes do ETL.
    """
    def __init__(self, paths: ProjectPaths):
        self.paths = paths

    def run_lifecycle_setup(self) -> bool:
        """Executa o ciclo de vida de preparação do ambiente de execução."""
        print("[INFO] Iniciando validação de infraestrutura local...")
        
        folders_created = self._build_directory_tree()
        master_sheet_verified = self._check_master_sheet_presence()
        
        if folders_created and master_sheet_verified:
            print("[SUCCESS] Fase 1 concluída. Ambiente Orientado a Objetos pronto.\n")
            return True
        else:
            print("[ERROR] Falha crítica na preparação do ambiente. Verifique os logs.\n")
            return False

    def _build_directory_tree(self) -> bool:
        """Varre os caminhos encapsulados e cria as pastas ausentes de forma segura."""
        try:
            for directory in self.paths.get_all_required_directories():
                if not directory.exists():
                    directory.mkdir(parents=True, exist_ok=True)
                    print(f"[SETUP] Diretório criado: dados/{directory.name}/")
                else:
                    print(f"[INFO] Estrutura validada: dados/{directory.name}/")
            return True
        except Exception as e:
            print(f"[CRITICAL] Erro ao instanciar árvore de diretórios: {e}")
            return False

    def _check_master_sheet_presence(self) -> bool:
        """Garante que o arquivo Excel (UI/Painel de Controle) está posicionado."""
        if self.paths.planilha_mestre.exists():
            print(f"[INFO] Interface do usuário identificada: {self.paths.planilha_mestre.name}")
            return True
        else:
            print(f"\n[WARNING] ALERTA: Arquivo '{self.paths.planilha_mestre.name}' ausente!")
            print(f"[WARNING] Certifique-se de salvá-lo em: {self.paths.planilha_mestre.parent}\n")
            return False


if __name__ == "__main__":
    # Ponto de entrada isolado para validação da Fase 1
    # Captura a raiz do projeto de forma dinâmica (2 níveis acima deste arquivo)
    PROJECT_ROOT = Path(__file__).resolve().parent.parent
    
    # Injeção de Dependência clássica da POO
    app_paths = ProjectPaths(base_dir=PROJECT_ROOT)
    environment = EnvironmentInitializer(paths=app_paths)
    
    # Execução do setup do ecossistema
    environment.run_lifecycle_setup()