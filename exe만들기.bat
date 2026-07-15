@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 스마트오더 라벨 - exe 빌드

set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY ( where python >nul 2>nul && set "PY=python" )
if not defined PY ( echo 파이썬이 필요합니다. python.org 에서 설치하세요. & pause & exit /b 1 )

echo 필요한 도구 설치(최초 1회)...
%PY% -m pip install --disable-pip-version-check -r requirements.txt pyinstaller

echo.
echo exe 빌드 중... (1~3분)
%PY% -m PyInstaller --noconfirm --clean --name SmartOrderLabel --onedir --console ^
  --add-data "templates;templates" ^
  --add-data "static;static" ^
  --add-data "masters;masters" ^
  --hidden-import xlrd --hidden-import openpyxl ^
  app.py

echo.
echo 완료! dist\SmartOrderLabel\ 폴더 안의 SmartOrderLabel.exe 를 실행하면 됩니다.
echo (배포 시 dist\SmartOrderLabel 폴더 전체를 복사해서 주세요. _internal 폴더도 함께)
pause
