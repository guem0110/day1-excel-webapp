@echo off
chcp 65001 >nul
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  echo Node.js is required. Install it from https://nodejs.org
  pause
  exit /b 1
)
if not exist node_modules (
  echo Installing packages...
  call npm install
  if errorlevel 1 (
    pause
    exit /b 1
  )
)
echo Starting dashboard...
call npm run dev
pause
