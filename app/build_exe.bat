@echo off
cd /d "%~dp0"

echo ================================================
echo   梁溪区工单分析工具 - 一键打包 exe
echo ================================================
echo.

echo [1/5] 检查 Python 环境...
where python >nul 2>&1
if errorlevel 1 goto NO_PYTHON
python -c "import struct;assert struct.calcsize('P')*8==64"
if errorlevel 1 goto PY32
echo   Python 检查通过 (64位)

echo [2/5] 安装依赖（首次约 2-4 分钟）...
python -m pip install -r requirements.txt
if errorlevel 1 goto PIP_FAIL

echo [3/5] 打包前代码自检...
python -c "import sys;sys.path.insert(0,'.');from api import Api;print('   [通过] 核心模块导入正常')"
if errorlevel 1 goto CHECK_FAIL

echo [4/5] PyInstaller 打包（约 2-5 分钟）...
python -m PyInstaller --clean -y 工单分析工具.spec
if errorlevel 1 goto BUILD_FAIL

echo [5/5] 自动验证 exe（启动 6 秒后自动关闭）...
if not exist "dist\梁溪区工单分析工具.exe" goto NO_EXE
start "" "dist\梁溪区工单分析工具.exe"
timeout /t 6 /nobreak >nul
tasklist /fi "imagename eq 梁溪区工单分析工具.exe" | find /i "梁溪区工单分析工具.exe" >nul
if errorlevel 1 goto EXE_DIED
taskkill /f /im "梁溪区工单分析工具.exe" >nul 2>&1
echo   [通过] exe 正常启动并保持运行，已自动关闭。
echo.
echo ================================================
echo   打包成功！产物: %~dp0dist\梁溪区工单分析工具.exe
echo   交付: 把这一个 exe 文件拷到客户电脑即可。
echo ================================================
pause
exit /b 0

:NO_PYTHON
echo   [错误] 未检测到 Python。请安装 Python 3.11 64位，
echo   安装时勾选 Add Python to PATH。
pause
exit /b 1

:PY32
echo   [错误] 当前 Python 是 32 位，请改用 64 位 Python。
pause
exit /b 1

:PIP_FAIL
echo   [错误] 依赖安装失败，请检查网络后重试。
pause
exit /b 1

:CHECK_FAIL
echo   [错误] 代码自检未通过，请截图本窗口发给开发。
pause
exit /b 1

:BUILD_FAIL
echo   [错误] PyInstaller 打包失败，请截图本窗口发给开发。
pause
exit /b 1

:NO_EXE
echo   [错误] 未找到产物 exe，请截图本窗口发给开发。
pause
exit /b 1

:EXE_DIED
taskkill /f /im "梁溪区工单分析工具.exe" >nul 2>&1
echo   [警告] exe 启动后 6 秒内退出了。
if exist error.log type error.log
echo   若无崩溃日志，多半是杀毒软件拦截，请添加信任后重试。
pause
exit /b 1
