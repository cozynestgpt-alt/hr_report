@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title 월례회의 인원보고서 자동화

echo ============================================
echo   월례회의 인원보고서 자동화
echo   NAS 경로: \\nas\CN_AMD\월례회의_자료\hr_report
echo ============================================
echo.

pushd "%~dp0" 2>nul
if errorlevel 1 (
    echo [오류] 폴더에 접속할 수 없습니다: %~dp0
    echo NAS 연결 상태를 확인해주세요.
    echo.
    pause
    exit /b 1
)

where python >nul 2>nul
if errorlevel 1 (
    where py >nul 2>nul
    if errorlevel 1 (
        echo [오류] Python이 설치되어 있지 않습니다. IT 담당자에게 문의해주세요.
        echo.
        pause
        popd
        exit /b 1
    )
    set PYCMD=py
) else (
    set PYCMD=python
)

set /p YM=기준월을 입력하세요 (예: 202605, 엔터만 누르면 당월 자동 적용):

echo.
if "%YM%"=="" (
    %PYCMD% main.py
) else (
    %PYCMD% main.py %YM%
)

echo.
echo ============================================
echo   작업이 완료되었습니다.
echo   output 폴더에서 결과 파일을 확인하세요.
echo   (엔터를 누르면 창이 닫힙니다)
echo ============================================
pause >nul

popd
endlocal
