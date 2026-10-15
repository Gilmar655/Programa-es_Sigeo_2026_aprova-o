@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Dashboard Programacoes 2026 - Enel - Atualizacao Automatica
set "EXCEL_PROGRAMACOES=C:\Users\BR0101687168\OneDrive - Enel Spa\Documents\Programações_2026\Programações_2026.xlsx"
echo.
echo =============================================================
echo          DASHBOARD DE PROGRAMACOES DE PROJETOS 2026
echo =============================================================
echo  Excel configurado:
echo  %EXCEL_PROGRAMACOES%
echo.
echo  O dashboard abrira no navegador neste computador.
echo  Mantenha esta janela aberta para atualizar automaticamente.
echo.
where py >nul 2>nul
if not errorlevel 1 (
    py -3 "%~dp0servidor.py"
    goto fim
)
where python >nul 2>nul
if not errorlevel 1 (
    python "%~dp0servidor.py"
    goto fim
)
echo ERRO: Python 3 nao foi encontrado neste Windows.
echo Instale Python 3 em https://www.python.org/downloads/
echo Durante a instalacao, habilite "Add Python to PATH".
echo Depois, execute este INICIAR_DASHBOARD.bat novamente.
:fim
echo.
echo O servidor foi encerrado ou nao conseguiu iniciar.
pause
endlocal
