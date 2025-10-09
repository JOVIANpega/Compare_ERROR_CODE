@echo off
chcp 65001 > nul
echo ========================================
echo Error Code Comparer 打包工具
echo ========================================

REM 設定變數
set APP_NAME=ErrorCodeComparer
set DIST_DIR=dist_exe

echo.
echo 1. 清理舊的建置檔案...
if exist %DIST_DIR% rmdir /s /q %DIST_DIR%
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo 2. 創建發佈目錄...
mkdir %DIST_DIR%

echo.
echo 3. 安裝 PyInstaller（如果尚未安裝）...
pip install pyinstaller

echo.
echo 4. 使用 PyInstaller 打包...
pyinstaller --onefile ^
    --windowed ^
    --icon=pal.ico ^
    --name=%APP_NAME% ^
    --add-data "guide_popup;guide_popup" ^
    --add-data "EXCEL;EXCEL" ^
    --add-data "VERSION.py;." ^
    --add-data "version_tool.py;." ^
    --add-data "version_config.py;." ^
    --add-data "version_manager.py;." ^
    --add-data "update_version.py;." ^
    --add-data "README.md;." ^
    --add-data "VERSION_USAGE.md;." ^
    --add-data "AI_RECOMMENDATION_USAGE.md;." ^
    main.py

echo.
echo 5. 複製執行檔到發佈目錄...
copy "dist\%APP_NAME%.exe" "%DIST_DIR%\"

echo.
echo 6. 複製額外檔案到發佈目錄...
copy "pal.ico" "%DIST_DIR%\"
copy "README.md" "%DIST_DIR%\"
copy "setup.txt" "%DIST_DIR%\"
copy "VERSION_USAGE.md" "%DIST_DIR%\"
copy "AI_RECOMMENDATION_USAGE.md" "%DIST_DIR%\"
xcopy "guide_popup" "%DIST_DIR%\guide_popup\" /E /I
xcopy "EXCEL" "%DIST_DIR%\EXCEL\" /E /I

echo.
echo 7. 創建啟動腳本...
echo @echo off > "%DIST_DIR%\start.bat"
echo echo 啟動 Error Code Comparer... >> "%DIST_DIR%\start.bat"
echo %APP_NAME%.exe >> "%DIST_DIR%\start.bat"
echo pause >> "%DIST_DIR%\start.bat"

echo.
echo 8. 創建版本資訊檔案...
echo Error Code Comparer v1.5.0 > "%DIST_DIR%\version.txt"
echo 建置日期: %date% %time% >> "%DIST_DIR%\version.txt"
echo 建置環境: Windows >> "%DIST_DIR%\version.txt"

echo.
echo 9. 清理 PyInstaller 臨時檔案...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del *.spec

echo.
echo ========================================
echo 打包完成！
echo ========================================
echo 執行檔位置: %DIST_DIR%\%APP_NAME%.exe
echo 發佈目錄: %DIST_DIR%\
echo.
echo 檔案清單:
dir /b "%DIST_DIR%"
echo.
echo 按任意鍵退出...
pause > nul
