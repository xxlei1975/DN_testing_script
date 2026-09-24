@echo off
rem ============================================================
rem  Africa_DN_Testing - one-click launcher
rem
rem  1. reuses the interpreter recorded in cache\python_path.txt, but ONLY
rem     when that record was written on this very computer and the
rem     interpreter still works. A cache folder copied from another PC is
rem     ignored instead of being trusted (it contains that PC's paths).
rem  2. reuses dig from PATH or from tools\bind\
rem  3. if anything is missing, calls scripts\setup_env.ps1 which installs
rem     it automatically (winget, or the official python.org installer)
rem  4. finally runs domain_dns_test.py with all given arguments
rem
rem  All text in this file is ASCII on purpose: cmd.exe re-reads a .bat
rem  file byte-by-byte and mangles multi-byte characters, which can make
rem  the script execute unexpected lines. Chinese messages come from the
rem  Python program instead.
rem
rem  Examples:
rem     run_dns_test.bat                 full run
rem     run_dns_test.bat --limit 5       first 5 domains only
rem     run_dns_test.bat --dry-run       preview, no file written
rem     run_dns_test.bat --inplace       write back to the original file
rem ============================================================
setlocal EnableExtensions

rem ---------- step 0: keep the standard Windows tools reachable ----------
rem Some machines hand this file a broken / minimal %PATH%, so nothing
rem outside Windows\System32 can be found. Symptoms seen in the wild:
rem   'powershell.exe' is not recognized as an internal or external command
rem   'where' / 'find' / ... is not recognized
rem Prepend the standard folders so powershell.exe and friends always work.
if not defined SystemRoot set "SystemRoot=C:\Windows"
set "SYS32=%SystemRoot%\System32"
set "PATH=%SYS32%;%SystemRoot%;%SYS32%\Wbem;%SYS32%\WindowsPowerShell\v1.0;%PATH%"

cd /d "%~dp0"

set "PYFILE=%~dp0cache\python_path.txt"
set "SETUP=%~dp0scripts\setup_env.ps1"
set "PY="
set "PSEXE="
set "RC_SETUP=1"

rem ---------- fast path: interpreter recorded by a previous run ----------
call :load_cached_python
if defined PY call :check_python
if not defined PY goto :provision

rem dig must be reachable as well, otherwise the environment needs fixing
call :check_dig
if not defined PY goto :provision
goto :launch

rem ---------- slow path: provision this machine first ----------
:provision
echo [INFO] Checking environment - Python / dependencies / dig may be installed now.
echo.
call :run_setup
if "%RC_SETUP%"=="0" goto :after_setup
echo [WARN] setup_env.ps1 returned %RC_SETUP%. Trying to continue anyway.
echo.
:after_setup
rem setup_env.ps1 has just rewritten the cache file for THIS computer, so it
rem is read again here - but it is never trusted blindly: whatever it claims
rem is re-verified (exists + can import the dependencies) before it is used.
set "PY="
call :load_cached_python
if defined PY call :check_python
if defined PY goto :launch
call :find_python
if not defined PY goto :no_python

rem ---------- run ----------
:launch
echo [INFO] Python : %PY%
echo [INFO] Program: %~dp0domain_dns_test.py
echo.
"%PY%" "%~dp0domain_dns_test.py" %*
set "RC=%errorlevel%"
echo.
if "%RC%"=="0" (
    echo [DONE] Finished successfully.
) else (
    echo [FAILED] exit code %RC% - see the messages above.
)
echo.
pause
exit /b %RC%

:no_python
echo [ERROR] Python is still unavailable. Run this manually to see details:
echo         "%PSEXE%" -NoProfile -ExecutionPolicy Bypass -File "%SETUP%"
echo.
pause
exit /b 1

rem ============================================================
rem  subroutines (kept out of IF-blocks on purpose: %var% inside a
rem  parenthesised block is expanded when the block is parsed, which is
rem  what used to print "setup_env.ps1 returned ." - i.e. an empty code)
rem ============================================================

rem ---- read cache\python_path.txt, machine checked ----
:load_cached_python
set "CACHE_MACHINE="
set "CACHE_PY="
if not exist "%PYFILE%" goto :load_cached_done
for /f "usebackq tokens=1,* delims==" %%A in ("%PYFILE%") do (
    if /i "%%A"=="machine" set "CACHE_MACHINE=%%B"
    if /i "%%A"=="python" set "CACHE_PY=%%B"
)
:load_cached_done
if not defined CACHE_PY goto :eof
if not defined COMPUTERNAME goto :load_cached_ok
if /i "%CACHE_MACHINE%"=="%COMPUTERNAME%" goto :load_cached_ok
echo [INFO] cache\python_path.txt was written on computer %CACHE_MACHINE% - ignoring it here.
goto :eof
:load_cached_ok
if exist "%CACHE_PY%" goto :load_cached_use
echo [INFO] the cached interpreter no longer exists - ignoring it.
goto :eof
:load_cached_use
set "PY=%CACHE_PY%"
goto :eof

rem ---- a usable interpreter must exist and import the dependencies ----
:check_python
if not defined PY goto :eof
if not exist "%PY%" goto :clear_py
"%PY%" -c "import openpyxl, requests" >nul 2>nul
if errorlevel 1 goto :clear_py
goto :eof
:clear_py
echo [INFO] the recorded interpreter is not usable on this computer - ignoring it.
set "PY="
goto :eof

rem ---- dig must be available: PATH or tools\bind ----
:check_dig
if not defined PY goto :eof
if exist "%~dp0tools\bind\dig.exe" goto :eof
where dig >nul 2>nul
if errorlevel 1 set "PY="
goto :eof

rem ---- last resort: python.exe / py.exe found on PATH ----
:find_python
for /f "delims=" %%I in ('where python.exe 2^>nul') do call :try_python "%%I"
if defined PY goto :eof
for /f "delims=" %%I in ('where py.exe 2^>nul') do if not defined PY call :try_python "%%I"
goto :eof

:try_python
if defined PY goto :eof
set "CANDIDATE=%~1"
if not exist "%CANDIDATE%" goto :eof
rem the Microsoft Store alias is a 0-byte stub which can import nothing,
rem so the dependency test below also filters it out
"%CANDIDATE%" -c "import openpyxl, requests" >nul 2>nul
if errorlevel 1 goto :eof
set "PY=%CANDIDATE%"
goto :eof

rem ---- locate PowerShell (never rely on %PATH% alone) ----
:find_powershell
if defined PSEXE goto :eof
set "PSDIR=%SYS32%\WindowsPowerShell\v1.0"
if exist "%PSDIR%\powershell.exe" set "PSEXE=%PSDIR%\powershell.exe"
if defined PSEXE goto :eof
if exist "%ProgramFiles%\PowerShell\7\pwsh.exe" set "PSEXE=%ProgramFiles%\PowerShell\7\pwsh.exe"
if defined PSEXE goto :eof
for /f "delims=" %%I in ('where powershell.exe 2^>nul') do if not defined PSEXE set "PSEXE=%%I"
if defined PSEXE goto :eof
for /f "delims=" %%I in ('where pwsh.exe 2^>nul') do if not defined PSEXE set "PSEXE=%%I"
goto :eof

rem ---- run setup_env.ps1 and remember its real exit code ----
:run_setup
set "RC_SETUP=1"
call :find_powershell
if not defined PSEXE goto :no_powershell
"%PSEXE%" -NoProfile -ExecutionPolicy Bypass -File "%SETUP%"
set "RC_SETUP=%errorlevel%"
goto :eof

:no_powershell
echo [WARN] Windows PowerShell could not be located on this computer.
echo        The environment cannot be provisioned automatically.
echo        Install Python 3 by hand (https://www.python.org/downloads/windows/,
echo        tick "Add python.exe to PATH"), then run this file again.
echo.
goto :eof
