@echo off
chcp 65001 >nul 2>&1
echo ====================================
echo   茶颜看板自动刷新
echo ====================================
echo.
"C:\Users\62398\.workbuddy\binaries\python\envs\default\Scripts\python.exe" "C:\Users\62398\waterbar\auto_refresh.py" %*
echo.
echo 按任意键关闭...
pause >nul
