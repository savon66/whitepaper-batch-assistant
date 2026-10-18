@echo off
setlocal EnableExtensions
chcp 65001 >nul

title 白皮批量制作助手 v1.0

rem ==================================================
rem 自动进入本 BAT 所在的插件根目录
rem ==================================================
cd /d "%~dp0"

echo.
echo ==================================================
echo            白皮批量制作助手 v1.0
echo ==================================================
echo.
echo 当前插件目录：
echo %CD%
echo.

rem ==================================================
rem 1. 检查 Node.js
rem ==================================================
set "NODE_EXE="

for /f "delims=" %%I in ('where node.exe 2^>nul') do (
    if not defined NODE_EXE set "NODE_EXE=%%I"
)

if not defined NODE_EXE (
    echo [错误] 当前电脑没有检测到 Node.js。
    echo.
    echo 请先安装 Node.js LTS Windows 安装版。
    echo 安装时请保持“添加到 PATH”选项启用。
    echo 安装完成后关闭本窗口，再重新双击本 BAT。
    echo.
    pause
    exit /b 1
)

echo [正常] 已找到 Node.js：
echo %NODE_EXE%
echo.

"%NODE_EXE%" --version

if errorlevel 1 (
    echo.
    echo [错误] Node.js 已安装，但无法正常运行。
    echo 请重新安装与当前 Windows 系统架构匹配的 Node.js LTS。
    echo.
    pause
    exit /b 1
)

echo.

rem ==================================================
rem 2. 检查 npm
rem ==================================================
set "NPM_CMD="

for /f "delims=" %%I in ('where npm.cmd 2^>nul') do (
    if not defined NPM_CMD set "NPM_CMD=%%I"
)

if not defined NPM_CMD (
    echo [错误] 没有检测到 npm.cmd。
    echo.
    echo npm 通常随 Node.js 一起安装。
    echo 请重新安装 Node.js LTS，并确保加入系统 PATH。
    echo.
    pause
    exit /b 1
)

echo [正常] 已找到 npm：
echo %NPM_CMD%
echo.

call "%NPM_CMD%" --version

if errorlevel 1 (
    echo.
    echo [错误] npm 无法正常运行。
    echo 请重新安装 Node.js LTS。
    echo.
    pause
    exit /b 1
)

echo.

rem ==================================================
rem 3. 查找 wpsjs.cmd
rem Windows 必须调用 wpsjs.cmd，避免误调用无扩展名文件
rem ==================================================
set "WPSJS_CMD="

if exist "%APPDATA%\npm\wpsjs.cmd" (
    set "WPSJS_CMD=%APPDATA%\npm\wpsjs.cmd"
)

if not defined WPSJS_CMD (
    for /f "delims=" %%I in ('where wpsjs.cmd 2^>nul') do (
        if not defined WPSJS_CMD set "WPSJS_CMD=%%I"
    )
)

rem ==================================================
rem 4. 未安装 wpsjs 时自动安装
rem ==================================================
if not defined WPSJS_CMD (
    echo [提示] 尚未安装 WPS 开发工具 wpsjs。
    echo [提示] 正在自动安装，请保持网络连接……
    echo.

    call "%NPM_CMD%" install -g wpsjs

    if errorlevel 1 (
        echo.
        echo [错误] wpsjs 自动安装失败。
        echo.
        echo 可能原因：
        echo 1. 当前电脑无法连接 npm；
        echo 2. 网络代理设置异常；
        echo 3. 安全软件拦截；
        echo 4. npm 全局目录没有写入权限。
        echo.
        echo 可在 CMD 中手动执行：
        echo npm install -g wpsjs
        echo.
        pause
        exit /b 1
    )

    rem 安装后重新查找
    if exist "%APPDATA%\npm\wpsjs.cmd" (
        set "WPSJS_CMD=%APPDATA%\npm\wpsjs.cmd"
    )

    if not defined WPSJS_CMD (
        for /f "delims=" %%I in ('where wpsjs.cmd 2^>nul') do (
            if not defined WPSJS_CMD set "WPSJS_CMD=%%I"
        )
    )
)

if not defined WPSJS_CMD (
    echo.
    echo [错误] wpsjs 已安装，但没有找到 wpsjs.cmd。
    echo 请关闭本窗口，重新打开后再试。
    echo.
    pause
    exit /b 1
)

echo [正常] 已找到 wpsjs：
echo %WPSJS_CMD%
echo.

rem ==================================================
rem 5. 正确验证 wpsjs
rem 注意：必须使用 --version 或大写 -V，不能使用小写 -v
rem ==================================================
call "%WPSJS_CMD%" --version

if errorlevel 1 (
    echo.
    echo [警告] wpsjs 已存在，但版本检查失败。
    echo [提示] 正在自动重新安装一次……
    echo.

    call "%NPM_CMD%" uninstall -g wpsjs

    call "%NPM_CMD%" install -g wpsjs

    if errorlevel 1 (
        echo.
        echo [错误] wpsjs 重新安装失败。
        echo.
        pause
        exit /b 1
    )

    rem 重新定位 wpsjs.cmd
    set "WPSJS_CMD="

    if exist "%APPDATA%\npm\wpsjs.cmd" (
        set "WPSJS_CMD=%APPDATA%\npm\wpsjs.cmd"
    )

    if not defined WPSJS_CMD (
        for /f "delims=" %%I in ('where wpsjs.cmd 2^>nul') do (
            if not defined WPSJS_CMD set "WPSJS_CMD=%%I"
        )
    )

    if not defined WPSJS_CMD (
        echo.
        echo [错误] 重新安装后仍未找到 wpsjs.cmd。
        echo.
        pause
        exit /b 1
    )

    call "%WPSJS_CMD%" --version

    if errorlevel 1 (
        echo.
        echo [错误] wpsjs 重新安装后仍无法正常运行。
        echo.
        pause
        exit /b 1
    )
)

echo.

rem ==================================================
rem 6. 检查插件项目文件
rem ==================================================
if not exist "%CD%\package.json" (
    echo [错误] 当前目录中没有找到 package.json。
    echo.
    echo 请把本 BAT 放到插件项目根目录。
    echo 当前目录：
    echo %CD%
    echo.
    pause
    exit /b 1
)

if not exist "%CD%\main.js" (
    echo [错误] 当前目录中没有找到 main.js。
    echo 请确认复制的是完整插件文件夹。
    echo.
    pause
    exit /b 1
)

if not exist "%CD%\ribbon.xml" (
    echo [错误] 当前目录中没有找到 ribbon.xml。
    echo 请确认复制的是完整插件文件夹。
    echo.
    pause
    exit /b 1
)

if not exist "%CD%\taskpane.html" (
    echo [错误] 当前目录中没有找到 taskpane.html。
    echo 请确认复制的是完整插件文件夹。
    echo.
    pause
    exit /b 1
)

if not exist "%CD%\taskpane.js" (
    echo [错误] 当前目录中没有找到 taskpane.js。
    echo 请确认复制的是完整插件文件夹。
    echo.
    pause
    exit /b 1
)

echo [正常] 插件项目文件检查通过。
echo.

rem ==================================================
rem 7. Install project dependencies when missing
rem Skip npm install when node_modules already exists
rem ==================================================
if not exist "%CD%\node_modules\" (
    echo [提示] 当前项目尚未安装 Node 依赖。
    echo [提示] 正在执行 npm install，首次运行可能需要一些时间……
    echo.

    call "%NPM_CMD%" install

    if errorlevel 1 (
        echo.
        echo [错误] 项目依赖安装失败。
        echo 请检查网络、package.json 或 npm 配置。
        echo.
        pause
        exit /b 1
    )

    echo.
    echo [正常] 项目依赖安装完成。
    echo.
) else (
    echo [正常] 已检测到项目依赖 node_modules。
    echo.
)

rem ==================================================
rem 8. 启动插件
rem ==================================================
echo ==================================================
echo 环境检查完成，正在启动白皮批量制作助手 v1.0
echo ==================================================
echo.
echo 使用说明
echo [ 运行时请不要关闭此黑色窗口]
echo [ 关闭本窗口后，插件调试服务会停止]
echo [ WPS 打开后应显示“白皮制作”选项卡]
echo [ 首次启动 WPS 可能需要等待几秒]
echo.

call "%WPSJS_CMD%" debug

set "EXIT_CODE=%ERRORLEVEL%"

echo.
echo ==================================================
echo 插件服务已经结束
echo 返回代码：%EXIT_CODE%
echo ==================================================
echo.

if not "%EXIT_CODE%"=="0" (
    echo [提示] 如果 WPS 没有正常启动，请查看上方报错信息。
    echo.
)

pause
exit /b %EXIT_CODE%
