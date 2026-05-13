@echo off
setlocal

cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo Node.js is required to run Excel Mapper.
  echo Install the LTS version from https://nodejs.org, then run this file again.
  pause
  exit /b 1
)

where npm >nul 2>nul
if errorlevel 1 (
  echo npm is required to run Excel Mapper. It is included with Node.js.
  pause
  exit /b 1
)

if not exist node_modules (
  echo Installing Excel Mapper dependencies. This can take a few minutes the first time.
  call npm install
  if errorlevel 1 (
    pause
    exit /b 1
  )
)

start "" "http://localhost:3010"
call npm run dev:local
pause
