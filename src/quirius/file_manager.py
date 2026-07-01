# src/quirius/file_manager.py
import time
import shutil
import zipfile
import logging
from pathlib import Path

logger = logging.getLogger("quirius_fetcher.file_manager")


class ReportFileManager:
    def __init__(self, temp_dir: Path, output_dir: Path):
        self.temp_dir = Path(temp_dir).resolve()
        self.output_dir = Path(output_dir).resolve()

    def clear_temp_dir(self) -> None:
        """Removes all files inside the temporary download directory."""
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        for file_path in self.temp_dir.iterdir():
            try:
                if file_path.is_file():
                    file_path.unlink()
            except Exception as e:
                logger.warning(f"Could not clean up temporary file {file_path.name}: {e}")

    def extract_and_move_report(self, zip_path: Path, company_id: str) -> Path | None:
        """
        Extracts spreadsheet reports from a ZIP file and relocates them to the final output directory.
        """
        if not zip_path or not zip_path.exists():
            logger.error(f"ZIP archive not found at: {zip_path}")
            return None

        self.output_dir.mkdir(parents=True, exist_ok=True)

        if zip_path.suffix == '.zip':
            logger.info(f"Extracting Excel report from {zip_path.name}...")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                lista_interna = zip_ref.namelist()
                arquivo_excel = next((f for f in lista_interna if f.endswith(('.xlsx', '.csv'))), None)
                
                if arquivo_excel:
                    zip_ref.extract(arquivo_excel, self.temp_dir)
                    caminho_extraido = self.temp_dir / arquivo_excel
                    
                    nome_final = f"NFe_Materiais_Quirius_{company_id}_{int(time.time())}.xlsx"
                    caminho_final = self.output_dir / nome_final
                    shutil.move(str(caminho_extraido), str(caminho_final))
                    logger.info(f"[SUCCESS] Materials report ready for ETL at: {caminho_final.name}")
                    return caminho_final
                else:
                    logger.error("ZIP package did not contain any valid spreadsheet file (.xlsx/.csv).")
                    return None
        else:
            logger.warning(f"Downloaded file {zip_path.name} is not a ZIP. Leaving it in temp.")
            return zip_path
