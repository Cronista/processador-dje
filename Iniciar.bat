@echo off
title Processador do Diario Oficial
color 0A

echo ===================================================
echo     Iniciando o Processador do Diario Oficial...
echo ===================================================
echo.

:: Verifica se o ambiente virtual (Python isolado) ja existe
if not exist "venv\" (
    echo [1/3] Primeira execucao detectada. Configurando ambiente isolado...
    python -m venv venv
    
    echo [2/3] Instalando bibliotecas necessarias... (Isso ocorre apenas uma vez)
    call venv\Scripts\activate
    pip install -r requirements.txt --quiet
) else (
    :: Se ja existe, apenas ativa
    call venv\Scripts\activate
)

echo [3/3] Abrindo a interface no navegador...
echo.
:: O comando streamlit ja abre uma nova aba no navegador padrao
streamlit run app.py

pause