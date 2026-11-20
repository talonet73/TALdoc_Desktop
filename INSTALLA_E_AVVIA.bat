@echo off
title TALdoc Desktop - Installazione automatica
color 0A
echo.
echo ============================================
echo    TALdoc Desktop - Installazione guidata
echo ============================================
echo.

:: Controlla se Python è installato
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python non trovato.
    echo.
    echo 1. Scarica Python da: https://www.python.org/downloads/
    echo 2. Durante l'installazione SPUNTA "Add Python to PATH"
    echo 3. Riavvia questo file dopo aver installato Python
    echo.
    pause
    start https://www.python.org/downloads/
    exit /b
)

echo Python trovato. Procedo...
echo.

:: Installa le librerie
echo Installazione delle librerie necessarie...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo Errore durante l'installazione delle librerie.
    pause
    exit /b
)

echo.
echo Librerie installate correttamente.
echo.

:: Controlla Ollama
echo Controllo Ollama...
curl -s http://localhost:11434/api/tags >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo Ollama non risulta avviato.
    echo.
    echo Scaricalo da: https://ollama.com
    echo Installalo e avvialo, poi rilancia questo file.
    echo.
    pause
    start https://ollama.com
    exit /b
)

echo Ollama è attivo.
echo.
echo ============================================
echo    Installazione completata!
echo    Ora puoi usare TALdoc Desktop
echo ============================================
echo.
echo Avvio dell'app...
python main.py
pause
