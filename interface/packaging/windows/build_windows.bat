@echo off
REM Construit PubMedSearch.exe (portable) et, si Inno Setup est installe,
REM l'installateur PubMedSearch-Setup-x.y.z.exe. A lancer par double-clic.
setlocal
cd /d "%~dp0\..\.."

where py >nul 2>nul && (set PY=py -3) || (set PY=python)
%PY% --version >nul 2>nul || (
  echo Python 3.10 ou plus recent est requis : https://www.python.org/downloads/
  echo Cochez "Add python.exe to PATH" pendant l'installation.
  pause & exit /b 1
)

echo [1/5] Environnement virtuel...
if not exist .venv-build %PY% -m venv .venv-build || goto :err
call .venv-build\Scripts\activate.bat
python -m pip install -q --upgrade pip
python -m pip install -q -r requirements-gui.txt || goto :err

echo [2/5] Version dossier (pour l'installateur)...
if exist build rmdir /s /q build
if exist dist\PubMedSearch rmdir /s /q dist\PubMedSearch
python -m PyInstaller --noconfirm pubmed_gui.spec || goto :err

echo [3/5] Version portable en un seul fichier...
set PUBMED_ONEFILE=1
python -m PyInstaller --noconfirm --distpath dist\portable pubmed_gui.spec || goto :err
set PUBMED_ONEFILE=

echo [4/5] Test : le Python integre execute le script...
dist\PubMedSearch\PubMedSearch.exe --run-script resources\pubmed_search.py --help > build\smoke.txt 2>&1
findstr /c:"--output" build\smoke.txt >nul || (echo ECHEC du test & type build\smoke.txt & goto :err)
echo     OK

echo [5/5] Installateur Inno Setup...
for /f %%v in ('python packaging\version.py') do set VER=%%v
set ISCC=
where iscc >nul 2>nul && set ISCC=iscc
if not defined ISCC if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set ISCC="%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set ISCC="%ProgramFiles%\Inno Setup 6\ISCC.exe"
if defined ISCC (
  %ISCC% /Q /DAppVersion=%VER% packaging\windows\installer.iss || goto :err
  echo     Installateur : dist\PubMedSearch-Setup-%VER%.exe
) else (
  echo     Inno Setup absent : installateur non cree ^(https://jrsoftware.org/isdl.php^)
)

echo.
echo Termine :
echo   Portable     : dist\portable\PubMedSearch.exe
echo   Dossier      : dist\PubMedSearch\PubMedSearch.exe
pause
exit /b 0

:err
echo.
echo La construction a echoue (voir les messages ci-dessus).
pause
exit /b 1
