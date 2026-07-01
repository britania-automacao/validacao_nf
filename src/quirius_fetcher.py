# src/quirius_fetcher.py
import os
import time
import shutil
import zipfile
import logging
import argparse
from pathlib import Path
from datetime import datetime
from dataclasses import dataclass
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

logger = logging.getLogger("quirius_fetcher")


def setup_logger(diretorio_base: Path) -> logging.Logger:
    historico_dir = diretorio_base / "dados" / "historico"
    historico_dir.mkdir(parents=True, exist_ok=True)
    
    logger_obj = logging.getLogger("quirius_fetcher")
    logger_obj.setLevel(logging.INFO)
    
    if not logger_obj.handlers:
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - [%(name)s] %(message)s')
        
        # File handler
        file_handler = logging.FileHandler(historico_dir / "quirius.log", encoding='utf-8')
        file_handler.setFormatter(formatter)
        logger_obj.addHandler(file_handler)
        
        # Console stream handler
        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(formatter)
        logger_obj.addHandler(stream_handler)
        
    return logger_obj


def get_driver(headless: bool = False, download_dir: Path | None = None, insecure_origin: str | None = None) -> WebDriver:
    logger.info("Initializing Chrome WebDriver...")
    
    options = Options()
    # Commented out to allow dynamic download directory preferences to work without prompt dialogs
    # options.add_argument("--guest")
    # options.add_argument("--incognito")
    options.add_argument("--window-size=1280,720")
    options.add_argument("--disable-plugins-discovery")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    
    if insecure_origin:
        options.add_argument(f"--unsafely-treat-insecure-origin-as-secure={insecure_origin}")

    if download_dir:
        download_dir = Path(download_dir).resolve()
        download_dir.mkdir(parents=True, exist_ok=True)
        prefs = {
            "download.default_directory": str(download_dir),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": False,
            "safebrowsing.disable_download_protection": True,
            "download.restrictions": 0,
            "profile.default_content_settings.popups": 0,
            "profile.default_content_setting_values.automatic_downloads": 1
        }
        options.add_experimental_option("prefs", prefs)

    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(options=options)
    
    if headless and download_dir:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": str(download_dir)
        })

    if headless:
        driver.set_window_size(1920, 1080)
    else:
        driver.maximize_window()
        
    logger.info("Driver successfully instantiated.")
    return driver


@dataclass
class QuiriusExtractionTask:
    company_id: str
    filter_name: str
    download_dir: Path
    output_dir: Path
    username: str = ""
    password: str = ""
    insecure_origin: str = ""
    date_range_days: int = 0  # 0 means only 1 day (yesterday)

    def run(self, run_date: datetime, headless: bool) -> Path | None:
        # 1. Resolve configuration values
        username = self.username or os.getenv("QUIRIUS_USER") or ""
        password = self.password or os.getenv("QUIRIUS_PASS") or ""
        insecure_origin = self.insecure_origin or os.getenv("QUIRIUS_URL") or "http://10.0.10.57:600"
        
        if not username or not password:
            logger.error(f"Credentials missing for company ID '{self.company_id}'. Skipping task.")
            return None

        pasta_temp = Path(self.download_dir).resolve()
        pasta_destino = Path(self.output_dir).resolve()
        
        pasta_temp.mkdir(parents=True, exist_ok=True)
        pasta_destino.mkdir(parents=True, exist_ok=True)
        
        # Clean temporary folder before starting
        for file_path in pasta_temp.iterdir():
            try:
                if file_path.is_file():
                    file_path.unlink()
            except Exception as e:
                logger.warning(f"Could not clean up temporary file {file_path.name}: {e}")
                
        # Calculate dates
        data_fim = run_date - relativedelta(days=1)
        data_inicio = data_fim - relativedelta(days=self.date_range_days)
        str_inicio = data_inicio.strftime("%d/%m/%Y")
        str_fim = data_fim.strftime("%d/%m/%Y")
        
        logger.info(f"=== Starting task for Company ID: {self.company_id} ===")
        logger.info(f"Query Period: {str_inicio} to {str_fim} | Headless: {headless}")
        
        driver: WebDriver | None = None
        try:
            driver = get_driver(headless=headless, download_dir=pasta_temp, insecure_origin=insecure_origin)
            actions = QuiriusActions(driver)
            
            # Log in and navigate to the reports page
            actions.login(username=username, password=password)
            actions.selecionar_empresa(self.company_id)
            actions.selecionar_modulo("DF-e")
            actions.selecionar_tab("Relatórios")
            actions.selecionar_submenu("Resumo NF-e")
            actions.switch_to_main_iframe()
            
            # Apply filters and run query
            actions.selecionar_filtro(self.filter_name)
            actions.selecionar_data_emissao(str_inicio, str_fim)
            actions.pesquisar()
            
            if actions.pesquisa_tem_retorno(timeout=1200):
                downloaded_file = actions.exportar_nfe(pasta_temp, timeout=1200)
                logger.info(f"Download complete: {downloaded_file}")
                
                if downloaded_file and downloaded_file.suffix == '.zip':
                    logger.info(f"Extracting Excel report from {downloaded_file.name}...")
                    with zipfile.ZipFile(downloaded_file, 'r') as zip_ref:
                        lista_interna = zip_ref.namelist()
                        arquivo_excel = next((f for f in lista_interna if f.endswith(('.xlsx', '.csv'))), None)
                        
                        if arquivo_excel:
                            zip_ref.extract(arquivo_excel, pasta_temp)
                            caminho_extraido = pasta_temp / arquivo_excel
                            
                            nome_final = f"NFe_Materiais_Quirius_{self.company_id}_{int(time.time())}.xlsx"
                            caminho_final = pasta_destino / nome_final
                            shutil.move(str(caminho_extraido), str(caminho_final))
                            logger.info(f"[SUCCESS] Materials report ready for ETL at: {caminho_final.name}")
                            return caminho_final
                        else:
                            logger.error("ZIP package did not contain any valid spreadsheet file (.xlsx/.csv).")
                else:
                    logger.warning(f"Downloaded file {downloaded_file.name} is not a ZIP. Leaving it in temp.")
                    return downloaded_file
            else:
                logger.info("Search query returned zero results.")
                return None
                
        except Exception as e:
            logger.exception(f"[CRITICAL ERROR] Quirius extraction failed: {e}")
            return None
        finally:
            if driver:
                logger.info("Tearing down Chrome browser.")
                driver.quit()
            # Post-execution temporary files cleanup
            for file_path in pasta_temp.iterdir():
                try:
                    if file_path.is_file():
                        file_path.unlink()
                except Exception:
                    pass


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
        logger.info("Switching to form iframe...")
        iframe = self.wait.until(
            EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'Forms/ReceiverForms/Crud/frm')]"))
        )
        self.driver.switch_to.frame(iframe)

    def login(self, username: str, password: str):
        logger.info(f"Navigating to login page: {self.URL_LOGIN}")
        self.driver.get(self.URL_LOGIN)
        self.wait.until(EC.presence_of_element_located(self.Elements.INPUT_LOGIN)).send_keys(username)
        self.wait.until(EC.presence_of_element_located(self.Elements.INPUT_SENHA)).send_keys(password)
        scroll_and_click(self.driver, self.wait, self.Elements.BTN_ENTRAR)

    def selecionar_empresa(self, id_empresa: str):        
        logger.info(f"Selecting company with ID: {id_empresa}")
        script_js = """
        const select = document.getElementById('ddlEmpresa');
        if (select) {
            select.value = arguments[0];
            select.dispatchEvent(new Event('change', { bubbles: true }));
        }
        """
        self.driver.execute_script(script_js, str(id_empresa))
        time.sleep(2)
        scroll_and_click(self.driver, self.wait, self.Elements.BTN_CONFIRMAR_SELECAO_EMPRESA)

    def selecionar_modulo(self, module: str) -> None:
        logger.info(f"Navigating to module: {module}")
        scroll_and_click(self.driver, self.wait, (By.XPATH, f"//input[@type='image' and @title='{module}']"))

    def selecionar_tab(self, tab: str):
        logger.info(f"Selecting menu tab: {tab}")
        try:
            locator = (By.XPATH, f"//input[(contains(@class, 'menuAba1') or contains(@class, 'menuAba2')) and @type='submit' and @value='{tab}']")
            scroll_and_click(self.driver, self.wait, locator, scroll=True, to_be_clickable=True, js_click=True)
        except TimeoutException:
            raise TimeoutException(f"Couldn't find/click '{tab}' tab.") from None

    def selecionar_submenu(self, parent: str = 'Resumo NFS-e', submenu: str | None = None):
        logger.info(f"Selecting submenu: {parent} -> {submenu}")
        try:
            menu_locator = (By.XPATH, f"//span[text()[normalize-space()='{parent}']]")
            scroll_and_click(self.driver, self.wait, menu_locator, scroll=True, to_be_clickable=True, js_click=True)
        except TimeoutException:
            raise TimeoutException(f"Couldn't find/click '{parent}' menu.") from None

        if not submenu:
            return

        try:
            submenu_locator = (By.XPATH, f"//tr[@id='{parent}']//span[normalize-space(text())='{submenu}']")
            scroll_and_click(self.driver, self.wait, submenu_locator, scroll=True, to_be_clickable=True, js_click=True)
        except TimeoutException:
            raise TimeoutException(f"Couldn't find/click '{submenu}' submenu.") from None

    def selecionar_filtro(self, name: str) -> None:
        logger.info(f"Applying filter: '{name}'")
        filter_select_element = self.wait.until(
            EC.presence_of_element_located(self.Elements.DROPDOWN_FILTER),
            f"Filter select '{name}' not found"
        )
        try:
            select = Select(filter_select_element)
            select.select_by_visible_text(name)
        except Exception as e:
            logger.error(f"Failed to select filter '{name}': {e}")
            raise

    def selecionar_data_emissao(self, de: str, ate: str):
        logger.info(f"Setting emission date filter to: {de} to {ate}")
        try:
            date_input = self.wait.until(EC.presence_of_element_located((By.XPATH, ".//input[contains(@id, 'tbxData1Line3')]")))
            date_input_2 = self.wait.until(EC.presence_of_element_located((By.XPATH, ".//input[contains(@id, 'tbxData2Line3')]")))

            # Clear and enter new dates
            date_input.clear()
            date_input.send_keys(de)
            date_input_2.clear()
            date_input_2.send_keys(ate)
            logger.info("Emission dates set successfully.")
        except TimeoutException:
            raise TimeoutException("Couldn't find/set 'Data Emissão' filter.") from None

    def pesquisar(self) -> None:
        try:
            logger.info("Executing search query...")
            locator = (By.XPATH, "//input[@type='submit' and @title='Pesquisar']")
            scroll_and_click(self.driver, self.wait, locator)
        except TimeoutException:
            raise TimeoutException("Could not find/click the search button.") from None

    def pesquisa_tem_retorno(self, timeout: int = 30) -> bool:
        wait = WebDriverWait(self.driver, timeout)
        try:
            table = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder_GridViewBO"]')))
            if "Não foram encontrados registros que atendam as condições." in table.text:
                return False
            return True
        except Exception as e:
            logger.warning(f"Error checking search results: {e}")
            return False

    def exportar_nfe(self, download_dir: Path, timeout: int = 1800) -> Path:
        download_dir = Path(download_dir).resolve()
        download_dir.mkdir(parents=True, exist_ok=True)
        
        # Capture files that exist before clicking export
        initial_files = set(os.listdir(download_dir))
        
        logger.info("Clicking the Export button...")
        scroll_and_click(self.driver, self.wait, self.Elements.BUTTON_EXPORT_NFE)

        logger.info("Waiting for file download to start and complete...")
        end_time = time.time() + timeout
        while time.time() < end_time:
            current_files = set(os.listdir(download_dir))
            new_files = current_files - initial_files
            
            # Filter in-progress and complete downloads
            completed_new_files = [
                f for f in new_files 
                if not f.endswith((".crdownload", ".tmp", ".download", ".part"))
            ]
            in_progress_new_files = [
                f for f in new_files 
                if f.endswith((".crdownload", ".tmp", ".download", ".part"))
            ]
            
            if completed_new_files and not in_progress_new_files:
                file_path = download_dir / completed_new_files[0]
                
                # Check for file stability
                try:
                    initial_size = file_path.stat().st_size
                    time.sleep(1)
                    if file_path.stat().st_size == initial_size and initial_size > 0:
                        logger.info(f"Download complete: {file_path}")
                        return file_path
                except OSError:
                    pass
                    
            time.sleep(1)

        raise TimeoutException(f"Download did not complete within {timeout} seconds.")


def main():
    # Resolve base paths
    raiz_projeto = Path(__file__).resolve().parent.parent
    caminhos = CaminhosProjeto(diretorio_base=raiz_projeto)
    
    # Initialize logging configuration
    setup_logger(raiz_projeto)
    
    # CLI setup
    parser = argparse.ArgumentParser(description="Task-based RPA extractor for Quirius Portal reports.")
    parser.add_argument(
        "--company", "-c",
        type=str,
        required=True,
        help="Target company ID in the dropdown (e.g. '1')."
    )
    parser.add_argument(
        "--filter", "-f",
        type=str,
        required=True,
        help="Saved search filter name in Quirius (e.g. 'validacao_nf_material')."
    )
    parser.add_argument(
        "--date", "-d",
        type=str,
        default=None,
        help="Reference run date in YYYY-MM-DD format. Default is today. Period ends at (date - 1 day)."
    )
    parser.add_argument(
        "--date-range-days", "-r",
        type=int,
        default=0,
        help="Number of days before the end date to query (default: 0, which queries yesterday only)."
    )
    parser.add_argument(
        "--headless", "-hl",
        action="store_true",
        help="Run browser in headless mode."
    )
    
    args = parser.parse_args()
    
    # Parse run date
    if args.date:
        try:
            run_date = datetime.strptime(args.date, "%Y-%m-%d")
        except ValueError:
            logger.error("Invalid date format. Please use YYYY-MM-DD.")
            return
    else:
        run_date = datetime.now()
        
    # Construct task directly from CLI arguments
    task = QuiriusExtractionTask(
        company_id=args.company,
        filter_name=args.filter,
        download_dir=caminhos.diretorio_base / "dados" / "temp_downloads",
        output_dir=caminhos.xml_materiais_quirius,
        date_range_days=args.date_range_days
    )
    
    logger.info(f"Starting execution of task for company ID: {task.company_id}...")
    
    result = task.run(run_date, args.headless)
    if result:
        logger.info(f"Task for company ID '{task.company_id}' succeeded.")
    else:
        logger.error(f"Task for company ID '{task.company_id}' failed.")


if __name__ == "__main__":
    main()
