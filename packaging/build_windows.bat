@echo off
rem Build markitdown.exe locally on Windows (Python 3.10+ must be installed).
rem Output: dist\markitdown.exe, dist\markitdown-skill.zip, dist\markitdown-skill-windows.zip
setlocal
cd /d "%~dp0.."

if not exist .venv-build\Scripts\python.exe (
    echo Creating virtual environment .venv-build ...
    py -3 -m venv .venv-build 2>nul || python -m venv .venv-build || goto :error
)
set PY=.venv-build\Scripts\python.exe

%PY% -m pip install --upgrade pip || goto :error
%PY% -m pip install -r packaging\requirements-build.txt || goto :error
%PY% -m PyInstaller --noconfirm --clean packaging\markitdown.spec || goto :error
%PY% tests\smoke_test.py --cmd dist\markitdown.exe || goto :error
%PY% packaging\package_skill.py --exe dist\markitdown.exe || goto :error

echo.
echo Build finished: %CD%\dist\markitdown.exe
pause
exit /b 0

:error
echo.
echo Build FAILED (see messages above).
pause
exit /b 1
