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
    :: O OCR (onnxruntime) ainda nao existe para o Python 3.15: usa o 3.14 ou o 3.13 se estiverem instalados
    py -3.14 -m venv venv 2>nul || py -3.13 -m venv venv 2>nul || python -m venv venv
) else (
    echo [1/3] Ambiente isolado encontrado.
)
call venv\Scripts\activate

:: Roda sempre: na primeira vez instala tudo; depois so instala o que mudou no sistema\requirements.txt
echo [2/3] Verificando bibliotecas necessarias...
pip install -r sistema\requirements.txt --quiet --disable-pip-version-check
:: OCR opcional: le trechos de PDF desenhados em vez de texto. Se falhar, o sistema funciona sem ele.
pip install -r sistema\requirements-ocr.txt --quiet --disable-pip-version-check || echo [aviso] OCR nao instalado. O sistema funciona normalmente sem ele.

echo [3/3] Abrindo a interface no navegador...
echo.
:: O comando streamlit ja abre uma nova aba no navegador padrao
streamlit run sistema\app.py

pause
