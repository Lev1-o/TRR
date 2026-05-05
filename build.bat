@echo off
title EQUIPO - Gerador de Executavel
color 0A

echo.
echo  ==========================================
echo   EQUIPO ^| Gerando Executavel do Sistema
echo  ==========================================
echo.

:: Verifica se o PyInstaller esta instalado
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo  [!] PyInstaller nao encontrado. Instalando...
    pip install pyinstaller
    echo.
)

:: Remove builds anteriores
echo  [1/3] Limpando builds anteriores...
if exist "dist" rmdir /s /q dist
if exist "build" rmdir /s /q build
if exist "EQUIPO_Revisoes.spec" del /q EQUIPO_Revisoes.spec
echo       OK!
echo.

:: Gera o executavel
echo  [2/3] Gerando executavel... (aguarde, pode demorar alguns minutos)
echo.

pyinstaller --onefile --windowed ^
    --icon="Imagem/minha_imagem.png" ^
    --name="EQUIPO_Revisoes" ^
    --add-data "Imagem;Imagem" ^
    --add-data "Core;Core" ^
    --hidden-import=customtkinter ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=win32com ^
    --hidden-import=win32com.client ^
    --hidden-import=requests ^
    app.py

echo.

:: Verifica se gerou com sucesso
if exist "dist\EQUIPO_Revisoes.exe" (
    echo  [3/3] Executavel gerado com sucesso!
    echo.
    echo  ==========================================
    echo   Arquivo: dist\EQUIPO_Revisoes.exe
    echo  ==========================================
    echo.
    echo  Abrindo pasta dist...
    explorer dist
) else (
    echo  [ERRO] Falha ao gerar o executavel.
    echo  Verifique os erros acima e tente novamente.
)

echo.
pause
