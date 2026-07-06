@echo off
REM ===========================================================================
REM  AI Agent Workshop (1) - Windows 11 one-command setup (.bat)
REM
REM  Installs and configures the full AI Agent toolchain: VS Code, Git,
REM  Node.js LTS, uv (Python), GitHub CLI, OpenAI Codex CLI, and the common
REM  VS Code extensions.
REM
REM  Safe to re-run (idempotent): already-installed tools are skipped.
REM
REM  How to run:
REM    - double-click this file, or
REM    - PowerShell:       & "$HOME\Downloads\setup-windows.bat"
REM    - Command Prompt:   "%USERPROFILE%\Downloads\setup-windows.bat"
REM ===========================================================================

setlocal enabledelayedexpansion
chcp 65001 >nul

set "OK_COUNT=0"
set "FAIL_COUNT=0"

echo.
echo ==============================================
echo  AI Agent Workshop - Windows 11 setup
echo ==============================================
echo.

REM ---- 0. Check winget ----------------------------------------------------
where winget >nul 2>&1
if errorlevel 1 (
  echo [ERROR] winget not found. Install "App Installer" from the Microsoft Store, then re-run this script.
  echo.
  pause
  exit /b 1
)

REM ---- 1. winget packages -------------------------------------------------
REM  format:  call :install_pkg  <winget-id>  <verify-cmd>  <display-name>
call :install_pkg "Microsoft.VisualStudioCode" "code" "VS Code"
call :install_pkg "Git.Git"                    "git"  "Git"
call :install_pkg "OpenJS.NodeJS.LTS"          "node" "Node.js LTS"
call :install_pkg "astral-sh.uv"               "uv"   "uv (Python)"
call :install_pkg "GitHub.cli"                 "gh"   "GitHub CLI"

REM ---- 2. Make new commands resolvable in this window --------------------
REM  (batch cannot reload the updated PATH on its own)
REM  Prepend well-known install locations so node/npm/code/uv/gh resolve now.
set "PATH=%PATH%;%ProgramFiles%\nodejs;%ProgramFiles%\Git\cmd;%LOCALAPPDATA%\Programs\Microsoft VS Code\bin;%USERPROFILE%\.local\bin;%ProgramFiles%\GitHub CLI"

REM ---- 3. Codex CLI (via npm) ---------------------------------------------
echo ----------------------------------------------
where codex >nul 2>&1
if not errorlevel 1 (
  echo [Codex CLI] already installed, skipping.
  call :mark_ok "Codex CLI" "already installed"
) else (
  where npm >nul 2>&1
  if errorlevel 1 (
    echo [ERROR] npm not found, skipping Codex CLI. Confirm Node.js installed, then reopen the terminal.
    call :mark_fail "Codex CLI" "npm not available"
  ) else (
    echo [Codex CLI] installing via npm...
    call npm install -g "@openai/codex"
    if errorlevel 1 ( call :mark_fail "Codex CLI" "npm install failed" ) else ( call :mark_ok "Codex CLI" "installed" )
  )
)

REM ---- 4. VS Code extensions ----------------------------------------------
REM  Note: openai.chatgpt is the Codex IDE extension, OpenAI's coding agent.
echo ----------------------------------------------
where code >nul 2>&1
if errorlevel 1 (
  echo [WARN] 'code' not on PATH yet, skipping extensions. Reopen the terminal and run this script again.
  call :mark_fail "VS Code extensions" "'code' not on PATH yet"
) else (
  call :install_ext "ms-python.python"
  call :install_ext "ms-toolsai.jupyter"
  call :install_ext "eamodio.gitlens"
  call :install_ext "dbaeumer.vscode-eslint"
  call :install_ext "openai.chatgpt"
)

REM ---- 5. Summary ---------------------------------------------------------
echo.
echo ==============================================
echo  Summary
echo ==============================================
echo   OK:   !OK_COUNT!
echo   FAIL: !FAIL_COUNT!
echo.
if "!FAIL_COUNT!"=="0" (
  echo All done! Close and reopen the terminal, then run 'codex' to sign in.
) else (
  echo !FAIL_COUNT! item^(s^) incomplete. Usually fixed by reopening the terminal and running this script again.
)
echo.
pause
endlocal
exit /b 0

REM ===========================================================================
REM  Subroutines
REM ===========================================================================

:install_pkg
REM  %~1 = winget id, %~2 = verify command, %~3 = display name
echo ----------------------------------------------
where %~2 >nul 2>&1
if not errorlevel 1 (
  call echo [%%~3] already installed, skipping.
  call :mark_ok "%~3" "already installed"
  goto :eof
)
echo [%~3] installing...
winget install --id %~1 --source winget --exact --silent --accept-source-agreements --accept-package-agreements
where %~2 >nul 2>&1
if errorlevel 1 (
  call :mark_fail "%~3" "install reported, command not found yet ^(may need to reopen terminal^)"
) else (
  call :mark_ok "%~3" "installed"
)
goto :eof

:install_ext
REM  %~1 = extension id
echo [extension] %~1 ...
call code --install-extension %~1 --force >nul
if errorlevel 1 ( call :mark_fail "ext: %~1" "install failed" ) else ( call :mark_ok "ext: %~1" "installed" )
goto :eof

:mark_ok
set /a OK_COUNT+=1
echo   [OK]   %~1 - %~2
goto :eof

:mark_fail
set /a FAIL_COUNT+=1
echo   [FAIL] %~1 - %~2
goto :eof
