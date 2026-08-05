# src/quirius/file_manager.py
import time
import zipfile
import logging
from pathlib import Path
from unistorage import Storage

logger = logging.getLogger("quirius_fetcher.file_manager")


class ReportFileManager:
    def __init__(self, temp_dir: Path | str, output_dir: Path | str, storage: Storage):
        self.temp_dir = str(temp_dir)
        self.output_dir = str(output_dir)
        self.storage = storage

    def clear_temp_dir(self) -> None:
        """Removes all files inside the temporary download directory."""
        self.storage.makedirs(self.temp_dir, exist_ok=True)
        try:
            for file_name in self.storage.listdir(self.temp_dir):
                file_path = f"{self.temp_dir}/{file_name}"
                if self.storage.isfile(file_path):
                    self.storage.remove(file_path)
        except Exception as e:
            logger.warning(f"Could not clean up temporary directory {self.temp_dir}: {e}")

    def extract_and_move_report(self, zip_path: Path | str, company_id: str, suffix: str = "") -> str | None:
        """
        Extracts spreadsheet reports from a ZIP file and relocates them to the final output directory.
        """
        zip_path_str = str(zip_path)
        if not self.storage.exists(zip_path_str):
            logger.error(f"ZIP archive not found at: {zip_path_str}")
            return None

        self.storage.makedirs(self.output_dir, exist_ok=True)

        if zip_path_str.lower().endswith('.zip'):
            logger.info(f"Extracting Excel report from {zip_path_str}...")
            with self.storage.open(zip_path_str, 'rb') as zf_file:
                with zipfile.ZipFile(zf_file, 'r') as zip_ref:
                    lista_interna = zip_ref.namelist()
                    arquivo_excel = next((f for f in lista_interna if f.endswith(('.xlsx', '.csv'))), None)
                    
                    if arquivo_excel:
                        excel_content = zip_ref.read(arquivo_excel)
                        # Aplica o sufixo para criar os arquivos parte1, parte2, etc.
                        nome_final = f"NFe_Materiais_Quirius{suffix}.xlsx"
                        caminho_final = f"{self.output_dir}/{nome_final}"
                        self.storage.write(caminho_final, excel_content)
                        logger.info(f"[SUCCESS] Materials report ready for ETL at: {caminho_final}")
                        return caminho_final
                    else:
                        logger.error("ZIP package did not contain any valid spreadsheet file (.xlsx/.csv).")
                        return None
        else:
            logger.warning(f"Downloaded file {zip_path_str} is not a ZIP. Leaving it in temp.")
            return zip_path_str