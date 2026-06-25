# src/quirius_fetcher.py
import os
import time
import shutil
import zipfile
from pathlib import Path
from datetime import datetime
from dateutil.relativedelta import relativedelta
from dotenv import load_dotenv

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

from config import CaminhosProjeto

load_dotenv()

class ExtratorQuirius:
    def __init__(self, caminhos: CaminhosProjeto):
        self.caminhos = caminhos
        
        # Credenciais e URLs (Ajuste o IP conforme o servidor Quirius da Britânia)
        self.url_login = "http://10.0.10.57:600/UIM/Admin/Login.aspx" 
        self.usuario = os.getenv("QUIRIUS_USER")
        self.senha = os.getenv("QUIRIUS_PASS")
        
        # Gestão de Pastas
        self.pasta_temp = self.caminhos.diretorio_base / "dados" / "temp_downloads"
        self.pasta_destino = self.caminhos.xml_materiais_quirius
        
        self.pasta_temp.mkdir(parents=True, exist_ok=True)
        self.pasta_destino.mkdir(parents=True, exist_ok=True)
        
        self.driver = None
        self.wait = None

    def _limpar_pasta_temporaria(self):
        """Garante que a pasta de downloads está vazia antes de começar."""
        for arquivo in self.pasta_temp.iterdir():
            try:
                arquivo.unlink()
            except Exception as e:
                print(f"[AVISO] Não foi possível apagar arquivo temporário: {e}")

    def iniciar_navegador(self):
        print("[QUIRIUS] Iniciando o Chrome WebDriver...")
        options = Options()
        
        # Força o Chrome a baixar arquivos para a nossa pasta temporária
        prefs = {
            "download.default_directory": str(self.pasta_temp),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        options.add_experimental_option("prefs", prefs)
        options.add_argument("--window-size=1920,1080")
        
        # Descomente a linha abaixo quando quiser que o robô rode "invisível" (em background)
        # options.add_argument("--headless=new") 

        servico = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=servico, options=options)
        self.wait = WebDriverWait(self.driver, 30)

    def fazer_login_e_selecionar_empresa(self, id_empresa: str):
        print("[QUIRIUS] Acessando tela de login...")
        self.driver.get(self.url_login)
        
        # Preenche Usuário e Senha
        self.wait.until(EC.presence_of_element_located((By.ID, "ctl00_txtLogin"))).send_keys(self.usuario)
        self.driver.find_element(By.ID, "ctl00_txtSenha").send_keys(self.senha)
        self.driver.find_element(By.ID, "ctl00_btnEntrar").click()
        
        print(f"[QUIRIUS] Login efetuado. Selecionando a empresa {id_empresa}...")
        # Lógica para forçar a seleção da empresa no dropdown
        script_js = """
        const select = document.getElementById('ddlEmpresa');
        if (select) {
            select.value = arguments[0];
            select.dispatchEvent(new Event('change', { bubbles: true }));
        }
        """
        self.driver.execute_script(script_js, str(id_empresa))
        time.sleep(2)
        
        # Clica em Confirmar/Enviar (Ajuste o ID se o botão confirmar chamar outra coisa)
        btn_confirmar = self.wait.until(EC.element_to_be_clickable((By.ID, "ButtonSelecionarEmpresa1")))
        btn_confirmar.click()

    def navegar_para_relatorios(self):
        print("[QUIRIUS] Navegando para DF-e > Relatórios > Resumo NF-e...")
        # 1. Clica no Módulo DF-e (Ajuste o XPath conforme o ícone do seu painel)
        btn_dfe = self.wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl12")))
        btn_dfe.click()
        
        # 2. Clica no Menu Relatórios
        btn_relatorios = self.wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_ctl00_ctl00_GeneralContent_BaseAjaxContent_BaseDefaultAjaxContent_MenuGuideControl1_Relatórios")))
        btn_relatorios.click()
        
        # 3. Clica no Submenu "Resumo NF-e"
        btn_resumo = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Resumo NF-e')]")))
        btn_resumo.click()

    def aplicar_filtros_e_exportar(self):
        print("[QUIRIUS] Aplicando filtro salvo e datas (D-90 a D-1)...")
        
        # --- Cálculo de Datas (Últimos 3 meses a partir de ontem) ---
        data_fim = datetime.now() - relativedelta(days=1)
        data_inicio = data_fim - relativedelta(months=3)
        str_inicio = data_inicio.strftime("%d/%m/%Y")
        str_fim = data_fim.strftime("%d/%m/%Y")
        print(f"          -> Período: {str_inicio} até {str_fim}")

        # 1. Seleciona o filtro "validacao_NF_material" no Dropdown de filtros salvos
        # (Ajuste o ID do select de filtros se necessário)
        dropdown_filtro = self.wait.until(EC.presence_of_element_located((By.ID, "id_do_dropdown_filtros_salvos")))
        dropdown_filtro.send_keys("validacao_NF_material")
        time.sleep(2)

        # 2. Preenche as Datas (Ajuste os IDs dos campos de data)
        input_dt_ini = self.driver.find_element(By.ID, "id_campo_data_inicial")
        input_dt_ini.clear()
        input_dt_ini.send_keys(str_inicio)
        
        input_dt_fim = self.driver.find_element(By.ID, "id_campo_data_final")
        input_dt_fim.clear()
        input_dt_fim.send_keys(str_fim)
        
        # 3. Clica em Exportar NF-e Detalhado
        print("[QUIRIUS] Solicitando exportação...")
        btn_exportar = self.driver.find_element(By.XPATH, "//span[contains(text(), 'Exportar NF-e Detalhado')]/..")
        btn_exportar.click()

    def aguardar_e_extrair_download(self):
        print("[QUIRIUS] Aguardando download do ZIP finalizar...")
        tempo_decorrido = 0
        timeout = 300 # Espera até 5 minutos pelo relatório pesado
        arquivo_zip = None

        while tempo_decorrido < timeout:
            arquivos = list(self.pasta_temp.iterdir())
            baixando = any(f.suffix == '.crdownload' for f in arquivos)
            
            if not baixando:
                zips = [f for f in arquivos if f.suffix == '.zip']
                if zips:
                    arquivo_zip = zips[0]
                    break
                    
            time.sleep(2)
            tempo_decorrido += 2

        if not arquivo_zip:
            raise TimeoutError("O download do Quirius demorou muito e falhou.")

        print(f"[QUIRIUS] Download concluído: {arquivo_zip.name}. Extraindo Excel...")
        
        with zipfile.ZipFile(arquivo_zip, 'r') as zip_ref:
            lista_interna = zip_ref.namelist()
            arquivo_excel = next((f for f in lista_interna if f.endswith(('.xlsx', '.csv'))), None)
            
            if arquivo_excel:
                zip_ref.extract(arquivo_excel, self.pasta_temp)
                caminho_extraido = self.pasta_temp / arquivo_excel
                
                # Move o Excel para a pasta definitiva de Materiais
                nome_final = f"NFe_Materiais_Quirius_{int(time.time())}.xlsx"
                caminho_final = self.pasta_destino / nome_final
                shutil.move(str(caminho_extraido), str(caminho_final))
                print(f"[SUCESSO] Relatório de Materiais pronto para o ETL em: {caminho_final.name}")
            else:
                print("[ERRO] O ZIP baixado não continha nenhum arquivo Excel.")

    def executar(self, id_empresa: str):
        try:
            self._limpar_pasta_temporaria()
            self.iniciar_navegador()
            self.fazer_login_e_selecionar_empresa(id_empresa)
            self.navegar_para_relatorios()
            self.aplicar_filtros_e_exportar()
            self.aguardar_e_extrair_download()
            
        except Exception as e:
            print(f"[FALHA CRÍTICA] Erro na automação do Quirius: {e}")
        finally:
            if self.driver:
                print("[QUIRIUS] Fechando o navegador.")
                self.driver.quit()
                self._limpar_pasta_temporaria()

if __name__ == "__main__":
    RAIZ_PROJETO = Path(__file__).resolve().parent.parent
    config = CaminhosProjeto(diretorio_base=RAIZ_PROJETO)
    
    # 1 representa o ID da Empresa Britânia (Ajuste conforme o Quirius)
    robo = ExtratorQuirius(caminhos=config)
    robo.executar(id_empresa="1")