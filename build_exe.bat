@echo off
REM Compila el Modulo QAQC como ejecutable de Windows.
REM Resultado en la carpeta dist: "Modulo QAQC.exe"
cd /d "%~dp0"

python -m pip install -r requirements.txt pyinstaller || goto :error

REM Regenerar la interfaz compilada a partir del .ui (por si se edito en Qt Designer)
python -m PyQt5.uic.pyuic qaqcbs.ui -o ui_qaqcbs.py || goto :error

python -m PyInstaller qaqc.spec --noconfirm --clean || goto :error

echo.
echo Listo: carpeta dist (ver UN_SOLO_ARCHIVO en qaqc.spec)
pause
exit /b 0

:error
echo.
echo *** Fallo la compilacion ***
pause
exit /b 1
