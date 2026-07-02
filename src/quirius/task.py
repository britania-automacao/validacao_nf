# src/quirius/task.py
import os
import logging
from pathlib import Path
from datetime import datetime
from dataclasses import dataclass
from dateutil.relativedelta import relativedelta
from selenium.webdriver.remote.webdriver import WebDriver

from config import settings
from .actions import QuiriusActions
from .file_manager import ReportFileManager

logger = logging.getLogger("quirius_fetcher.task")


@dataclass
class QuiriusExtractionTask:
    company_id: str
    download_dir: Path
    output_dir: Path
    username: str
    password: str
    date_range_days: int = 0  # 0 means only 1 day (yesterday)

    # O nome do filtro é fixo para extração de materiais do escopo do projeto
    FILTER_NAME = "validacao_nf_material"

    def run(self, driver: WebDriver, run_date: datetime) -> Path | None:
        if not self.username or not self.password:
            logger.error(f"Credentials missing for company ID '{self.company_id}'. Skipping task.")
            return None

        # Instantiate file manager and clear temp download directories
        output_dir = self.output_dir / self.company_id
        file_manager = ReportFileManager(self.download_dir, output_dir)
        file_manager.clear_temp_dir()

        # Calculate dates
        data_fim = run_date - relativedelta(days=1)
        data_inicio = data_fim - relativedelta(days=self.date_range_days)
        str_inicio = data_inicio.strftime("%d/%m/%Y")
        str_fim = data_fim.strftime("%d/%m/%Y")
        
        logger.info(f"=== Starting task for Company ID: {self.company_id} ===")
        logger.info(f"Query Period: {str_inicio} to {str_fim}")
        
        actions = QuiriusActions(driver)
        
        # Log in and navigate to the reports page
        actions.login(username=self.username, password=self.password)
        actions.selecionar_empresa(self.company_id)
        actions.selecionar_modulo("DF-e")
        actions.selecionar_tab("Relatórios")
        actions.selecionar_submenu("Resumo NF-e")
        actions.switch_to_main_iframe()
        
        # Apply filters and run query
        actions.selecionar_filtro(self.FILTER_NAME)
        actions.selecionar_data_emissao(str_inicio, str_fim)
        actions.pesquisar()
        
        if actions.pesquisa_tem_retorno(timeout=1200):
            downloaded_file = actions.exportar_nfe(self.download_dir, timeout=1200)
            logger.info(f"Download completed. File path: {downloaded_file}")
            
            # Delegate file extraction and clean up of temp items to ReportFileManager
            result_file = file_manager.extract_and_move_report(downloaded_file, self.company_id)
            
            # Clean up other files in temp dir as post-execution cleanup
            file_manager.clear_temp_dir()
            return result_file
        else:
            logger.info("Search query returned zero results.")
            return None
