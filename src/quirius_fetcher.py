# src/quirius_fetcher.py
import os
import logging
import argparse
from pathlib import Path
from datetime import datetime

from config import CaminhosProjeto, settings
from quirius import QuiriusExtractionTask, browser_session

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
        
    # Resolve download and output directories
    download_dir = caminhos.diretorio_base / "dados" / "temp_downloads"
    output_dir = caminhos.xml_materiais_quirius
    insecure_origin = settings.quirius_url
    
    # Construct task directly from CLI arguments
    task = QuiriusExtractionTask(
        company_id=args.company,
        username=settings.quirius_user,
        password=settings.quirius_pass,
        download_dir=download_dir,
        output_dir=output_dir,
        insecure_origin=insecure_origin,
        date_range_days=args.date_range_days
    )
    
    logger.info(f"Starting execution of task for company ID: {task.company_id}...")
    
    # Use browser_session context manager to control the browser lifecycle
    with browser_session(headless=args.headless, download_dir=download_dir, insecure_origin=insecure_origin) as driver:
        result = task.run(driver, run_date)
        
    if result:
        logger.info(f"Task for company ID '{task.company_id}' succeeded. Output file: {result}")
    else:
        logger.error(f"Task for company ID '{task.company_id}' failed.")


if __name__ == "__main__":
    main()
