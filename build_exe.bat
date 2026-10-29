@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
echo ============================================================
echo   Сборка TrajCompare.exe
echo ============================================================

where py >nul 2>&1 && (set PY=py -3) || (set PY=python)
%PY% --version >nul 2>&1
if errorlevel 1 (
  echo.
  echo [!] Python не найден. Установите Python 3.10-3.12 с https://www.python.org/downloads/
  echo     При установке отметьте "Add python.exe to PATH". Затем запустите этот файл снова.
  pause
  exit /b 1
)

if not exist .venv (
  echo Создаю виртуальное окружение...
  %PY% -m venv .venv || goto :err
)
call .venv\Scripts\activate.bat

echo Устанавливаю зависимости (нужен интернет, 1-3 минуты)...
python -m pip install --upgrade pip -q
python -m pip install -q -r requirements.txt pyinstaller || goto :err

echo Собираю exe (2-5 минут)...
pyinstaller --noconfirm --clean --onefile --windowed ^
  --name TrajCompare ^
  --hidden-import mpl_toolkits.mplot3d ^
  --exclude-module scipy --exclude-module IPython --exclude-module PyQt5 --exclude-module PySide6 ^
  trajcompare_gui.py || goto :err

echo.
echo ============================================================
echo   ГОТОВО:  %~dp0dist\TrajCompare.exe
echo   Этот файл можно копировать на любые компьютеры с Windows 10/11,
echo   Python на них не нужен.
echo ============================================================
explorer dist
pause
exit /b 0

:err
echo.
echo [!] Сборка не удалась. Пришлите текст выше разработчику.
pause
exit /b 1
