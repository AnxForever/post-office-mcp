@echo off
chcp 65001 >nul
echo ============================================================
echo   AionUi 邮局委办报表助手 - 一键配置
echo ============================================================
echo.
echo 请确认：
echo   1. AionUi 应用已启动（任务栏有图标）
echo   2. 已在 AionUi 中完成首次账号设置
echo.
pause

REM 优先用 python3，回退到 python
where python3 >nul 2>nul
if %errorlevel%==0 (
    python3 "%~dp0setup.py"
) else (
    python "%~dp0setup.py"
)

echo.
pause
