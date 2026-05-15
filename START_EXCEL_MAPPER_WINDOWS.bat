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

for /f "usebackq delims=" %%A in (`node -p "process.arch"`) do set NODE_ARCH=%%A
if not "%NODE_ARCH%"=="x64" if not "%NODE_ARCH%"=="arm64" (
  echo Excel Mapper requires 64-bit Node.js on Windows.
  echo Your current Node.js architecture is: %NODE_ARCH%
  echo.
  echo The app cannot start with 32-bit Node.js because Next.js cannot load its compiler.
  echo Install the Windows x64 LTS version from https://nodejs.org, then run this file again.
  pause
  exit /b 1
)

set ARCH_MARKER=.excel-mapper-node-arch
set NEED_INSTALL=0
if not exist node_modules set NEED_INSTALL=1
if not exist %ARCH_MARKER% set NEED_INSTALL=1
if exist %ARCH_MARKER% (
  set /p INSTALLED_ARCH=<%ARCH_MARKER%
  if not "%INSTALLED_ARCH%"=="%NODE_ARCH%" (
    echo Node.js architecture changed from %INSTALLED_ARCH% to %NODE_ARCH%.
    echo Reinstalling dependencies for this Windows architecture.
    if exist node_modules rmdir /s /q node_modules
    set NEED_INSTALL=1
  )
)

if "%NEED_INSTALL%"=="1" (
  echo Installing Excel Mapper dependencies. This can take a few minutes the first time.
  call npm install
  if errorlevel 1 (
    pause
    exit /b 1
  )
  echo %NODE_ARCH%>%ARCH_MARKER%
)

start "" "http://localhost:3010"
call npm run dev:local
pause
