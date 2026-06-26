@echo off
chcp 65001 >nul
echo ============================================
echo 信贷外勤拍照 App - 打包向导
echo ============================================
echo.
echo 此脚本将在 WSL 中运行 Buildozer 打包 APK
echo 请确认已安装 WSL Ubuntu 子系统
echo.

where wsl >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未检测到 WSL!
    echo 请先安装 WSL:
    echo   1. 以管理员打开 PowerShell
    echo   2. 运行: wsl --install -d Ubuntu
    echo   3. 重启电脑
    echo   4. 设置 Ubuntu 用户名和密码
    echo   5. 再次运行此脚本
    pause
    exit /b 1
)

echo [1/4] 启动 WSL 环境...
wsl -e bash -c "echo WSL 已就绪"

echo.
echo [2/4] 复制项目文件到 WSL...
wsl -e bash -c "mkdir -p ~/loan_photo_app"
xcopy /E /I /Y "%~dp0" "C:\Users\%USERNAME%\loan_photo_app\" >nul 2>&1
wsl -e bash -c "cp -r /mnt/c/Users/%USERNAME%/loan_photo_app/* ~/loan_photo_app/ 2>/dev/null; cp -r /mnt/d/hermes/loan_photo_app/* ~/loan_photo_app/ 2>/dev/null; echo 文件复制完成"

echo.
echo [3/4] 安装依赖并打包...
echo 这可能需要较长时间 (5-30分钟)...
wsl -e bash -c "cd ~/loan_photo_app && chmod +x setup_and_build.sh && bash setup_and_build.sh"

echo.
echo [4/4] 打包完成!
echo APK 文件在 bin 目录下
pause
