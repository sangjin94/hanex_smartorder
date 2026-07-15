@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 한익스프레스 스마트오더 라벨

rem --- 파이썬 찾기 (py 우선, 없으면 python) ---
set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY ( where python >nul 2>nul && set "PY=python" )
if not defined PY (
  echo.
  echo [설치 필요] 이 PC에 파이썬이 없습니다.
  echo   1) https://www.python.org/downloads/ 에서 Python 3.10 이상 설치
  echo   2) 설치 첫 화면에서 "Add python.exe to PATH" 를 꼭 체크하세요.
  echo   3) 설치 후 이 파일(실행.bat)을 다시 더블클릭하세요.
  echo.
  pause
  exit /b 1
)

rem --- 필요한 라이브러리 확인, 없으면 자동 설치(최초 1회, 인터넷 필요) ---
%PY% -c "import flask, openpyxl, xlrd" 1>nul 2>nul
if errorlevel 1 (
  echo.
  echo 처음 실행이라 필요한 라이브러리를 설치합니다. 1~2분 걸릴 수 있어요...
  %PY% -m pip install --disable-pip-version-check -r requirements.txt
  if errorlevel 1 (
    echo.
    echo [오류] 라이브러리 설치에 실패했습니다. 인터넷 연결을 확인하고 다시 시도하세요.
    pause
    exit /b 1
  )
)

echo.
echo ================================================
echo   한익스프레스 스마트오더 라벨 생성기
echo   잠시 후 브라우저가 자동으로 열립니다.
echo   주소: http://127.0.0.1:5057
echo   종료하려면 이 검은 창을 닫으세요.
echo ================================================
echo.
%PY% app.py
pause
