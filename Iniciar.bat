@echo off
title Processador do Diario Oficial
color 0A

echo ===================================================
echo     Iniciando o Processador do Diario Oficial...
echo ===================================================
echo.

:: Garante que o sistema rode a partir da pasta deste arquivo
cd /d "%~dp0"

:: Verifica se o ambiente virtual (Python isolado) ja existe
if not exist "venv\" (
    echo [1/3] Primeira execucao detectada. Configurando ambiente isolado...
    python -m venv venv
) else (
    echo [1/3] Ambiente isolado encontrado.
)
call venv\Scripts\activate

:: Roda sempre: na primeira vez instala tudo; depois so instala o que mudou no requirements.txt
echo [2/3] Verificando bibliotecas necessarias...
pip install -r requirements.txt --quiet --disable-pip-version-check

echo [3/3] Abrindo a interface no navegador...
echo.
:: O comando streamlit ja abre uma nova aba no navegador padrao
streamlit run app.py

pause
