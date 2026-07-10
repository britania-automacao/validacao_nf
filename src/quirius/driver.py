# src/quirius/driver.py
import logging
from pathlib import Path
from contextlib import contextmanager
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.remote.webdriver import WebDriver

logger = logging.getLogger("quirius_fetcher.driver")


def get_driver(headless: bool = False, download_dir: Path | None = None, insecure_origins: list[str] | None = None) -> WebDriver:
    logger.info("Initializing Chrome WebDriver...")
    
    options = Options()
    options.add_argument("--window-size=1280,720")
    options.add_argument("--disable-plugins-discovery")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--allow-running-insecure-content")
    options.add_argument("--disable-web-security")
    
    # Configuração de estratégia normal para não interromper requisições complexas
    options.page_load_strategy = 'normal'
    
    if insecure_origins:
        origins_str = ",".join(insecure_origins)
        options.add_argument(f"--unsafely-treat-insecure-origin-as-secure={origins_str}")

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
    
    # --- CORREÇÃO CRÍTICA DO TIMEOUT ---
    # Eleva a paciência da renderização no browser para 50 minutos (3000s)
    driver.set_page_load_timeout(3000)
    driver.set_script_timeout(3000)
    
    # Hack profundo no urllib3 para forçar o Timeout da conexão de rede
    # Impede que a biblioteca corte a conexão se o Quirius demorar mais de 120s
    try:
        driver.command_executor._client_config._timeout = 3000
    except Exception as e:
        logger.warning(f"Não foi possível sobrescrever a config profunda do urllib3: {e}")
    # -----------------------------------

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


@contextmanager
def browser_session(headless: bool = False, download_dir: Path | None = None, insecure_origins: list[str] | None = None):
    driver = get_driver(headless=headless, download_dir=download_dir, insecure_origins=insecure_origins)
    try:
        yield driver
    finally:
        logger.info("Tearing down Chrome browser session.")
        driver.quit()