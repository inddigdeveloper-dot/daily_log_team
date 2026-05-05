@echo off
:: ============================================================
::  CODING AGENT LAUNCHER v2
::  - Cleans RAM
::  - Starts Ollama with qwen2.5-coder:3b
::  - Starts LiteLLM proxy (Anthropic format -> Ollama)
::  - Launches Claude Code pointed at local model
::  Run as Administrator for best RAM cleanup
:: ============================================================

title Coding Agent Launcher
color 0A

echo.
echo  ==========================================
echo   CODING AGENT LAUNCHER v2
echo   Local Model: qwen2.5-coder:3b
echo  ==========================================
echo.

:: ── STEP 1: RAM CLEANUP ──────────────────────────────────────
echo [1/5] Cleaning RAM...

taskkill /F /IM "SearchIndexer.exe"       >nul 2>&1
taskkill /F /IM "OneDrive.exe"            >nul 2>&1
taskkill /F /IM "Teams.exe"               >nul 2>&1
taskkill /F /IM "Spotify.exe"             >nul 2>&1
taskkill /F /IM "Discord.exe"             >nul 2>&1
taskkill /F /IM "slack.exe"               >nul 2>&1
taskkill /F /IM "MicrosoftEdgeUpdate.exe" >nul 2>&1

powershell -NoProfile -Command ^
  "[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers(); [System.GC]::Collect()" >nul 2>&1

echo  RAM cleanup done.
echo.

:: ── STEP 2: START OLLAMA ─────────────────────────────────────
echo [2/5] Starting Ollama...

where ollama >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Ollama not found!
    echo  Download from: https://ollama.com
    pause
    exit /b 1
)

:: Kill old instance and start fresh
taskkill /F /IM "ollama.exe" >nul 2>&1
timeout /t 2 >nul

start /min "" ollama serve
timeout /t 4 >nul

:: Pull model if not present
ollama list | findstr "qwen2.5-coder:3b" >nul 2>&1
if errorlevel 1 (
    echo  Pulling qwen2.5-coder:3b ^(first time only^)...
    ollama pull qwen2.5-coder:3b
    if errorlevel 1 (
        echo  ERROR: Failed to pull model. Check internet.
        pause
        exit /b 1
    )
) else (
    echo  Model ready.
)
echo.

:: ── STEP 3: START LITELLM PROXY ──────────────────────────────
echo [3/5] Starting LiteLLM proxy...
echo  ^(Translates Claude Code API calls to Ollama format^)

where python >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python not found! Download from https://python.org
    pause
    exit /b 1
)

:: Install LiteLLM if missing
python -c "import litellm" >nul 2>&1
if errorlevel 1 (
    echo  Installing LiteLLM ^(one-time^)...
    pip install "litellm[proxy]" -q
)

:: Kill anything on port 4000
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":4000 "') do (
    taskkill /F /PID %%a >nul 2>&1
)
timeout /t 1 >nul

:: Start proxy in minimized window
start /min "LiteLLM Proxy" cmd /c "python -m litellm --model ollama/qwen2.5-coder:3b --port 4000 --drop_params"

echo  Waiting for proxy to be ready...
timeout /t 6 >nul
echo  Proxy running at http://localhost:4000
echo.

:: ── STEP 4: SET ENVIRONMENT ──────────────────────────────────
echo [4/5] Setting up environment...

set ANTHROPIC_BASE_URL=http://localhost:4000
set ANTHROPIC_API_KEY=sk-local-ollama
set CLAUDE_DISABLE_TELEMETRY=1

:: Go to wherever this BAT file lives
cd /d "%~dp0"

echo  Directory : %CD%
echo  Endpoint  : http://localhost:4000 ^(Ollama local^)
echo  Model     : qwen2.5-coder:3b
echo.

:: ── STEP 5: LAUNCH CLAUDE CODE ───────────────────────────────
echo [5/5] Launching Claude Code...

where claude >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Claude Code not found!
    echo  Run: npm install -g @anthropic-ai/claude-code
    pause
    exit /b 1
)

echo.
echo  ==========================================
echo   AGENT READY - Type your task below
echo.
echo   Examples:
echo   ^> Build a Python file renamer script
echo   ^> Fix the bug in main.py
echo   ^> Create a REST API for a todo app
echo   ^> Explain this codebase and suggest improvements
echo  ==========================================
echo.

claude --dangerously-skip-permissions

:: ── CLEANUP ON EXIT ──────────────────────────────────────────
echo.
echo  Cleaning up processes...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":4000 "') do (
    taskkill /F /PID %%a >nul 2>&1
)
taskkill /F /IM "ollama.exe" >nul 2>&1
echo  Done. Press any key to close.
pause
