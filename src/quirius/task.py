# src/quirius/task.py
import os
import logging
from pathlib import Path
from datetime import datetime
from dataclasses import dataclass
from dateutil.relativedelta import relativedelta
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By

from config import settings
from unistorage import Storage
from .actions import QuiriusActions
from .file_manager import ReportFileManager

logger = logging.getLogger("quirius_fetcher.task")


@dataclass
class QuiriusExtractionTask:
    company_id: str
    download_dir: Path | str
    output_dir: Path | str
    username: str
    password: str
    storage: Storage
    date_range_days: int = 0  # 0 means only 1 day (yesterday)

    # O nome do filtro é fixo para extração de materiais do escopo do projeto
    FILTER_NAME = "validacao_nf_material"

    def run(self, driver: WebDriver, run_date: datetime, timeout: int = 30) -> str | None:
        if not self.username or not self.password:
            logger.error(f"Credentials missing for company ID '{self.company_id}'. Skipping task.")
            return None

        # Instantiate file manager and clear temp download directories
        output_dir = f"{self.output_dir}/{self.company_id}"
        file_manager = ReportFileManager(self.download_dir, output_dir, self.storage)
        file_manager.clear_temp_dir()

        # Limpa arquivos antigos para não misturar com novas extrações
        try:
            if os.path.exists(output_dir):
                for f in os.listdir(output_dir):
                    if f.startswith("NFe_Materiais_Quirius") and f.endswith(".xlsx"):
                        os.remove(os.path.join(output_dir, f))
        except Exception as e:
            logger.warning(f"Não foi possível limpar os arquivos antigos: {e}")

        # Calculate dates
        data_fim_geral = run_date - relativedelta(days=1)
        data_inicio_geral = data_fim_geral - relativedelta(days=self.date_range_days)

        str_inicio = data_inicio_geral.strftime("%d/%m/%Y")
        str_fim = data_fim_geral.strftime("%d/%m/%Y")

        logger.info(f"=== Starting task for Company ID: {self.company_id} ===")
        logger.info(f"Total Query Period: {str_inicio} to {str_fim} (Única Extração)")

        actions = QuiriusActions(driver, timeout=timeout)
        arquivo_gerado = None
        sucesso = False
        
        # Executa um loop de 3 passagens (Para garantir a repescagem em instabilidades pesadas)
        for tentativa in [1, 2, 3]: 
            if sucesso:
                break
                
            if tentativa > 1:
                logger.info(f"\n{'='*50}\n INICIANDO RETENTATIVA {tentativa} DA EXTRAÇÃO ÚNICA\n{'='*50}")

            logger.info(f"\n--- Extraindo Bloco Único (Tentativa {tentativa}): {str_inicio} a {str_fim} ---")
            
            try:
                # Impede que o driver do Selenium fique eternamente bloqueado por requisições do Quirius
                TIMEOUT_SEGUNDOS = 3000 # 50 minutos
                driver.set_page_load_timeout(TIMEOUT_SEGUNDOS)
                driver.set_script_timeout(TIMEOUT_SEGUNDOS)

                try:
                    logger.info("Abrindo nova guia no navegador e limpando o histórico...")
                    driver.switch_to.new_window('tab')
                    while len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[0])
                        driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                    
                    driver.delete_all_cookies() 
                    
                except Exception as e_aba:
                    logger.warning(f"Erro ao recriar aba (o navegador pode estar instável): {e_aba}")

                driver.get(actions.URL_LOGIN)
                if "Login.aspx" in driver.current_url:
                    actions.login(username=self.username, password=self.password)
              
                actions.selecionar_empresa(self.company_id)
                actions.selecionar_modulo("DF-e")
                actions.selecionar_tab("Relatórios")
                actions.selecionar_submenu("Resumo NF-e")
                actions.switch_to_main_iframe()
                
                actions.selecionar_filtro(self.FILTER_NAME)
                actions.selecionar_data_emissao(str_inicio, str_fim)
                actions.pesquisar()
                
                logger.info("Esperando pesquisa terminar de carregar...")
                actions.esperar_carregar(timeout=300)
                
                # Escaneia a tela em busca do erro SQL "Incorrect Syntax" ou de expiração de sessão.
                # Se achar, crasha a execução na hora, evitando esperar 20 minutos à toa.
                try:
                    page_text = driver.find_element(By.TAG_NAME, "body").text.lower()
                    if "incorrect syntax near the keyword" in page_text or "sessão expirou" in page_text:
                        raise Exception("O Quirius retornou erro de Sintaxe SQL/Sessão na tela. A sessão atual do Quirius quebrou!")
                except Exception as inner_e:
                    if "O Quirius retornou erro" in str(inner_e):
                        raise inner_e # Repassa a exceção pra acionar o bloco de "except" lá em baixo
                    pass # Ignora se falhou em ler a tela e segue o jogo

                if actions.pesquisa_tem_retorno(timeout=300):
                    logger.info("Tabela carregada. Iniciando download (Timeout máximo absoluto de 20 minutos)...")
                    
                    downloaded_file = actions.exportar_nfe(self.download_dir, timeout=TIMEOUT_SEGUNDOS)
                    logger.info(f"Download da Extração Única concluído.")
                    
                    # Sem sufixo pois agora é um arquivo só
                    result_file = file_manager.extract_and_move_report(downloaded_file, self.company_id, suffix="")
                    
                    if result_file:
                        arquivo_gerado = result_file
                        sucesso = True
                else:
                    logger.info(f"Search query returned zero results para o período total.")
                    sucesso = True # Se buscou e veio vazio sem erros, é um "sucesso" técnico

            except Exception as e:
                # Se estourar os 20 min ou o Quirius retornar "Incorrect syntax", cai aqui
                logger.error(f"Falha Crítica ao processar a Extração: {e}")
                logger.error("=> Marcando para repescagem e FECHANDO a aba com defeito.")
                
                # MATA A ABA TRAVADA
                try:
                    driver.close() # Destrói fisicamente a aba onde a sessão expirou/travou
                    if len(driver.window_handles) == 0:
                        driver.switch_to.new_window('tab')
                    else:
                        driver.switch_to.window(driver.window_handles[-1])
                except Exception:
                    pass
                    
            finally:
                # Sempre limpa a pasta temp para não poluir
                file_manager.clear_temp_dir()

        if sucesso and arquivo_gerado:
            return f"1 part generated successfully."
        elif sucesso and not arquivo_gerado:
            return "No data found for the period."
        else:
            return None