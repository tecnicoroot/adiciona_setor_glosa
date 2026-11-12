@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set NOME=AdicionaSetorGlosa

echo ============================================
echo  Gerando executavel: %NOME%.exe
echo ============================================

where uv >nul 2>nul || (echo [ERRO] uv nao encontrado no PATH. & goto erro)

echo.
echo [1/3] Sincronizando dependencias...
uv sync || goto erro

echo.
echo [2/3] Compilando com PyInstaller...
uv run --with pyinstaller pyinstaller ^
    --noconfirm --clean ^
    --onefile --windowed ^
    --name %NOME% ^
    --paths src ^
    --add-data "sql\consulta_setor.sql;sql" ^
    --collect-submodules oracledb ^
    --collect-submodules robo_utils ^
    executar_gui.py || goto erro

echo.
echo [3/3] Copiando a query para ao lado do executavel...
if not exist "dist\sql" mkdir "dist\sql"
copy /y "sql\consulta_setor.sql" "dist\sql\" >nul || goto erro

echo.
echo ============================================
echo  Pronto! Executavel em: %CD%\dist\%NOME%.exe
echo  Distribua o .exe junto com a pasta dist\sql
echo ============================================
pause
exit /b 0

:erro
echo.
echo [ERRO] Falha ao gerar o executavel. Veja as mensagens acima.
pause
exit /b 1