@echo off
chcp 65001 > nul
echo ========================================
echo Error Code Comparer 快速打包工具
echo ========================================

REM 設定變數
set APP_NAME=ErrorCodeComparer
set DIST_DIR=dist_exe

echo.
echo 1. 清理舊的建置檔案...
if exist %DIST_DIR% rmdir /s /q %DIST_DIR%
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del *.spec

echo.
echo 2. 創建發佈目錄...
mkdir %DIST_DIR%

echo.
echo 3. 使用 PyInstaller 打包...
pyinstaller --onefile --windowed --icon=pal.ico --name=%APP_NAME% main.py

echo.
echo 4. 複製執行檔和必要檔案...
copy "dist\%APP_NAME%.exe" "%DIST_DIR%\"
copy "pal.ico" "%DIST_DIR%\"
copy "README.md" "%DIST_DIR%\"
copy "setup.txt" "%DIST_DIR%\"
xcopy "guide_popup" "%DIST_DIR%\guide_popup\" /E /I
xcopy "EXCEL" "%DIST_DIR%\EXCEL\" /E /I

echo.
echo 5. 清理臨時檔案...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del *.spec

echo.
echo ========================================
echo 快速打包完成！
echo ========================================
echo 執行檔位置: %DIST_DIR%\%APP_NAME%.exe
echo.
echo 按任意鍵退出...
pause > nul
