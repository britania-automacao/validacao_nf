# src/quirius/actions.py
import os
import time
import logging
from pathlib import Path
from utils import scroll_and_click

from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

logger = logging.getLogger("quirius_fetcher.actions")


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
