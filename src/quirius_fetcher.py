# src/quirius_fetcher.py
import os
import time
import shutil
import zipfile
from pathlib import Path
from datetime import datetime
from dateutil.relativedelta import relativedelta
from dotenv import load_dotenv

from utils import scroll_and_click

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException

from config import CaminhosProjeto

load_dotenv()

def get_driver(headless: bool = False) -> WebDriver:
    print("Driver: Instantiating...")
    
    options = Options()
    options.add_argument("--guest")
    options.add_argument("--window-size=1280,720")
    options.add_argument("--incognito")
    options.add_argument("--disable-plugins-discovery")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")

    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

    # driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager(driver_version=Constants.Chrome.VERSION).install()), options=options)
    driver = webdriver.Chrome(options=options)
    if headless:
        driver.set_window_size(1920, 1080) # type: ignore[reportUnknownMemberType]
    else:
        driver.maximize_window()
    print("Driver: instantiated.")
    return driver

class QuiriusActions:
    URL_LOGIN = "http://10.0.10.57:600/UIM/Admin/Login.aspx" 

    class Elements:
        INPUT_LOGIN = (By.ID, "ctl00_txtLogin")
        INPUT_SENHA = (By.ID, "ctl00_txtSenha")
        BTN_ENTRAR = (By.ID, "ctl00_btnEntrar")
        BTN_CONFIRMAR_SELECAO_EMPRESA = (By.ID, "ButtonSelecionarEmpresa1")
        DROPDOWN_FILTER = (By.XPATH, '//*[@id="ctl00_ContentPlaceHolder_SearchControlBO_ddlFiltersSaved"]')
        BUTTON_EXPORT_NFE = (By.XPATH, '//*[@id="ctl00_ContentPlaceHolder_btnExportarNFe"]')
        

    def __init__(self, driver: WebDriver):
        self.driver = driver
        self.wait = WebDriverWait(driver, 30)

    def switch_to_main_iframe(self):
        print("Switching to iframe...")
        iframe = self.wait.until(
            EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'Forms/ReceiverForms/Crud/frm')]"))
        )
        self.driver.switch_to.frame(iframe)

    def login(self, username: str, password: str):
        print("[QUIRIUS] Acessando tela de login...")
        # Preenche Usuário e Senha
        self.driver.get(self.URL_LOGIN)
        self.wait.until(EC.presence_of_element_located(self.Elements.INPUT_LOGIN)).send_keys(username)
        self.wait.until(EC.presence_of_element_located(self.Elements.INPUT_SENHA)).send_keys(password)
        self.wait.until(EC.presence_of_element_located(self.Elements.BTN_ENTRAR)).click()

    def selecionar_empresa(self, id_empresa: str):        
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
        btn_confirmar = self.wait.until(EC.element_to_be_clickable(self.Elements.BTN_CONFIRMAR_SELECAO_EMPRESA))
        btn_confirmar.click()

    def selecionar_modulo(self, module: str) -> None:
        print(f"[QUIRIUS] Selecionando módulo {module}...")

        module_btn = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, f"//input[@type='image' and @title='{module}']"))
        )

        module_btn.click()

    def selecionar_tab(self, tab: str):
        print(f"[QUIRIUS] Selecionando tab {tab}...")
        try:
            locator = (By.XPATH, f"//input[(contains(@class, 'menuAba1') or contains(@class, 'menuAba2')) and @type='submit' and @value='{tab}']")
            scroll_and_click(self.driver, self.wait, locator, scroll=True, to_be_clickable=True, js_click=True)
        except TimeoutException:
            raise TimeoutException(f"Couldn't find/click '{tab}' tab.") from None

    def selecionar_submenu(self, parent: str = 'Resumo NFS-e', submenu: str | None = None):
        print(f"[QUIRIUS] Selecionando submenu {parent} -> {submenu}...")
        try:
            menu_locator = (By.XPATH, f"//span[text()[normalize-space()='{parent}']]")
            scroll_and_click(self.driver, self.wait, menu_locator, scroll=True, to_be_clickable=True, js_click=True)
        except TimeoutException:
            raise TimeoutException(f"Couldn't find/click '{parent}' menu.") from None

        if not submenu: return

        try:
            submenu_locator = (By.XPATH, f"//tr[@id='{parent}']//span[normalize-space(text())='{submenu}']")
            scroll_and_click(self.driver, self.wait, submenu_locator, scroll=True, to_be_clickable=True, js_click=True)
        except TimeoutException:
            raise TimeoutException(f"Couldn't find/click '{submenu}' submenu.") from None

    def selecionar_filtro(self, name: str) -> None:
        print(f"[QUIRIUS] Selecionando filtro '{name}'...")
        filter_select_element = self.wait.until(
            EC.presence_of_element_located(self.Elements.DROPDOWN_FILTER),
            f"Filter select '{name}' not found"
        )

        try:
            select = Select(filter_select_element)
            select.select_by_visible_text(name)
        except Exception as e:
            print(f"[QUIRIUS] Erro ao selecionar filtro {name}: {e}")
            raise

    def selecionar_data_emissao(self, de: str, ate: str):
        """
        Altera o valor da data no filtro 'Data Emissão'.
        Args:
            de (str): data no formato DD/MM/AAAA (ex: "30/07/2025").
            ate (str): data no formato DD/MM/AAAA (ex: "30/07/2025").
        """
        print(f"[QUIRIUS] Setando data de emissão para {de} até {ate}...")


        try:
            print(f"[QUIRIUS] Encontrando campo de data...")
            field_select = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//select[contains(@id, 'ddlFieldsLine')]/option[@selected='selected' and contains(text(), 'Data Emissão')]/parent::select"))
            )

            # Sobe para o <tr> e desce até o <input> do campo de data
            #print("Finding row from: 'Data Emissão'...")
            #container_row = field_select.find_element(By.XPATH, "./ancestor::tr")
            #print(f"Finding input from: 'Data Emissão'...")
            #date_input = container_row.find_element(By.XPATH, ".//input[contains(@id, 'tbxData1Line3')]")
            #date_input_2 = container_row.find_element(By.XPATH, ".//input[contains(@id, 'tbxData2Line3')]")

            date_input = self.wait.until(EC.presence_of_element_located((By.XPATH, ".//input[contains(@id, 'tbxData1Line3')]")))
            date_input_2 = self.wait.until(EC.presence_of_element_located((By.XPATH, ".//input[contains(@id, 'tbxData2Line3')]")))

            # Limpa e insere nova data
            print(f"Setting field value from: 'Data Emissão' to '{de} até {ate}'...")
            date_input.clear()
            date_input.send_keys(de)
            date_input_2.clear()
            date_input_2.send_keys(ate)
            print("[QUIRIUS] Data de emissão setada.")

        except TimeoutException:
            raise TimeoutException("Couldn't find/set 'Data Emissão' filter.") from None

    def pesquisar(self) -> None:
        try:
            print("[QUIRIUS] Encontrando botão de pesquisar...")
            locator = (By.XPATH, "//input[@type='submit' and @title='Pesquisar']")
            print("[QUIRIUS] Clicando em pesquisar...")
            scroll_and_click(self.driver, self.wait, locator)
        except TimeoutException:
            raise TimeoutException("[QUIRIUS] Não foi possível encontrar/clicar no botão pesquisar.") from None

    def pesquisa_tem_retorno(self, timeout: int = 30) -> bool:
        wait = WebDriverWait(self.driver, timeout)
        try:
            table = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder_GridViewBO"]')))

            if "Não foram encontrados registros que atendam as condições." in table.text:
                return False
            else:
                return True
        except Exception as e:
            print(f"Error: {e}")
            return False

    def wait_for_download(self, download_dir: Path, timeout: int = 60):
        """
        Waits until Chrome finishes downloading files in the folder, while skipping
        the files in the ignore_files list if any.
        
        :param timeout: Maximum wait time in seconds.
        :param ignore_files: List of files to ignore (optional).
        :return: List of completed files.
        """
        ignore_files = os.listdir(download_dir)
        print(f"Ignoring {len(ignore_files)} files.")

        print("Waiting for download...")
        end_time = time.time() + timeout

        while time.time() < end_time:
            files = os.listdir(download_dir)
            # Chrome marks in-progress downloads as *.crdownload
            in_progress = [f for f in files if f.endswith((".crdownload", '.htm', '.tmp')) and f not in ignore_files]
            completed = [f for f in files if not f.endswith((".crdownload", '.htm', '.tmp')) and f not in ignore_files]

            # If no in-progress downloads and some completed ones, return completed files
            if completed and not in_progress:
                return [os.path.join(download_dir, filename) for filename in completed]

            time.sleep(1)

        raise TimeoutException("Download did not complete in time.")

    def exportar_nfe_detalhado(self, download_dir: Path, timeout: int = 1800):
        print("Finding export button...")
        button_export = self.wait.until(
            EC.element_to_be_clickable(self.Elements.BUTTON_EXPORT_NFE),
            "Couldn't find export button"
        )
        print("Pressing export")
        button_export.click()

        input()

        #downloaded = self.wait_for_download(download_dir=download_dir, timeout=timeout)
        #time.sleep(3)
        print("Download complete.")
        return []


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

    def aguardar_e_extrair_download(self, pasta_temp: Path, pasta_dest: Path):
        print("[QUIRIUS] Aguardando download do ZIP finalizar...")
        tempo_decorrido = 0
        timeout = 300 # Espera até 5 minutos pelo relatório pesado
        arquivo_zip = None

        while tempo_decorrido < timeout:
            arquivos = list(pasta_temp.iterdir())
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
                zip_ref.extract(arquivo_excel, pasta_temp)
                caminho_extraido = pasta_temp / arquivo_excel
                
                # Move o Excel para a pasta definitiva de Materiais
                nome_final = f"NFe_Materiais_Quirius_{int(time.time())}.xlsx"
                caminho_final = pasta_dest / nome_final
                shutil.move(str(caminho_extraido), str(caminho_final))
                print(f"[SUCESSO] Relatório de Materiais pronto para o ETL em: {caminho_final.name}")
            else:
                print("[ERRO] O ZIP baixado não continha nenhum arquivo Excel.")

class QuiriusExtractor:
    def __init__(self, caminhos: CaminhosProjeto):
        self.caminhos = caminhos
        
        # Credenciais e URLs (Ajuste o IP conforme o servidor Quirius da Britânia)
        self.usuario = os.getenv("QUIRIUS_USER") or ""
        self.senha = os.getenv("QUIRIUS_PASS") or ""
        
        # Gestão de Pastas
        self.pasta_temp = self.caminhos.diretorio_base / "dados" / "temp_downloads"
        self.pasta_destino = self.caminhos.xml_materiais_quirius
        
        self.pasta_temp.mkdir(parents=True, exist_ok=True)
        self.pasta_destino.mkdir(parents=True, exist_ok=True)
    

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
        driver = webdriver.Chrome(service=servico, options=options)
        return driver

    def executar(self, id_empresa: str):
        self._limpar_pasta_temporaria()
        self.iniciar_navegador()
        #self.fazer_login_e_selecionar_empresa(id_empresa)
        #self.navegar_para_relatorios()
        #self.aplicar_filtros_e_exportar()
        #self.aguardar_e_extrair_download()
            

    def run(self, run_date: datetime | None = None):
        print("[QUIRIUS] Run")
        run_date = run_date or datetime.now()
        usuario = os.getenv("QUIRIUS_USER") or ""
        senha = os.getenv("QUIRIUS_PASS") or ""

        data_fim = run_date - relativedelta(days=1)
        data_inicio = data_fim - relativedelta(months=3)
        str_inicio = data_inicio.strftime("%d/%m/%Y")
        str_fim = data_fim.strftime("%d/%m/%Y")
        print(f"[QUIRIUS] Run date: {run_date.strftime('%d/%m/%Y')} | Periodo: {str_inicio} até {str_fim}")

        print("[QUIRIUS] Iniciando navegador...")
        driver: WebDriver | None = None
        try:    
            driver = get_driver()
            actions = QuiriusActions(driver)
            actions.login(username=usuario, password=senha)
            actions.selecionar_empresa("1")
            actions.selecionar_modulo("DF-e")
            actions.selecionar_tab("Relatórios")
            actions.selecionar_submenu("Resumo NF-e")
            actions.switch_to_main_iframe()
            actions.selecionar_filtro("validacao_nf_material")
            actions.selecionar_data_emissao(str_inicio, str_fim)
            actions.pesquisar()
            if actions.pesquisa_tem_retorno(timeout=1200):
                downloaded = actions.exportar_nfe_detalhado(Path("./downloads"), timeout=1200)
                print(downloaded)
                

            input("Continuar?")
        except Exception as e:
            print(f"[FALHA CRÍTICA] Erro na automação do Quirius: {e}")
            import traceback
            traceback.print_exc()
        finally:
            if driver:
                print("[QUIRIUS] Fechando o navegador.")
                driver.quit()



def main():
    RAIZ_PROJETO = Path(__file__).resolve().parent.parent
    config = CaminhosProjeto(diretorio_base=RAIZ_PROJETO)
    
    # 1 representa o ID da Empresa Britânia (Ajuste conforme o Quirius)
    robo = QuiriusExtractor(caminhos=config)
    #robo.executar(id_empresa="1")

    robo.run()

if __name__ == "__main__":
    #driver = get_driver()
    #driver.get("")
    main()
