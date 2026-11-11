@echo off
:: Garante que a acentuação apareça corretamente no terminal
chcp 65001 >nul

echo ===================================================
echo     INICIANDO PIPELINE DE VALIDACAO DE NOTAS
echo ===================================================

:: Garante que o terminal rode a partir da pasta raiz (validacao_nf)
cd /d "%~dp0"

echo.
echo [0/5] Preparando Ambiente Virtual (venv)...

:: Verifica se a pasta venv existe. Se não, cria uma nova.
if not exist venv\Scripts\activate.bat (
    echo - Criando ambiente virtual...
    python -m venv venv
    if %ERRORLEVEL% neq 0 goto erro
)

echo - Ativando ambiente virtual...
call venv\Scripts\activate.bat

echo - Verificando e instalando dependencias...
python -m pip install -r requirements.txt
if %ERRORLEVEL% neq 0 goto erro

echo.
echo [1/5] Executando Quirius Fetcher (Multiplas Empresas)...

echo - Extraindo Empresa 1...
python src\quirius_fetcher.py -c 1 -r 90
if %ERRORLEVEL% neq 0 goto erro

echo - Extraindo Empresa 15...
python src\quirius_fetcher.py -c 15 -r 90
if %ERRORLEVEL% neq 0 goto erro

echo - Extraindo Empresa 30...
python src\quirius_fetcher.py -c 30 -r 90
if %ERRORLEVEL% neq 0 goto erro

echo - Extraindo Empresa 35...
python src\quirius_fetcher.py -c 35 -r 90
if %ERRORLEVEL% neq 0 goto erro

echo.
echo [2/5] Executando NFSe Fetcher...
python src\nfse_fetcher.py
if %ERRORLEVEL% neq 0 goto erro

echo.
echo [3/5] Executando ETL de Materiais...
python src\materials_etl.py
if %ERRORLEVEL% neq 0 goto erro

echo.
echo [4/5] Executando ETL de Servicos...
python src\services_etl.py
if %ERRORLEVEL% neq 0 goto erro

echo.
echo [5/5] Executando Envio de E-mail...
python src\email_sender.py
if %ERRORLEVEL% neq 0 goto erro

echo.
echo ===================================================
echo        PIPELINE CONCLUIDO COM SUCESSO!
echo ===================================================
exit /b 0

:erro
echo.
echo ===================================================
echo   [ERRO] O pipeline parou devido a uma falha.
echo   Verifique os logs na pasta historico e as mensagens acima.
echo ===================================================
pause
exit /b 1