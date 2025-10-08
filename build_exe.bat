@echo off
echo ========================================
echo Error Code Comparer 打包工具
echo ========================================

REM 設定變數
set APP_NAME=ErrorCodeComparer
set BUILD_DIR=build_exe
set DIST_DIR=dist_exe

echo.
echo 1. 清理舊的建置檔案...
if exist %BUILD_DIR% rmdir /s /q %BUILD_DIR%
if exist %DIST_DIR% rmdir /s /q %DIST_DIR%

echo.
echo 2. 創建建置目錄...
mkdir %BUILD_DIR%
mkdir %DIST_DIR%

echo.
echo 3. 複製必要檔案到建置目錄...
copy "main.py" "%BUILD_DIR%\"
copy "ui_manager.py" "%BUILD_DIR%\"
copy "excel_handler.py" "%BUILD_DIR%\"
copy "ai_recommendation_engine.py" "%BUILD_DIR%\"
copy "ai_prompt_templates.py" "%BUILD_DIR%\"
copy "file_finder.py" "%BUILD_DIR%\"
copy "config_manager.py" "%BUILD_DIR%\"
copy "error_code_compare.py" "%BUILD_DIR%\"
copy "excel_errorcode_search_ui.py" "%BUILD_DIR%\"
copy "VERSION.py" "%BUILD_DIR%\"
copy "version_tool.py" "%BUILD_DIR%\"
copy "version_config.py" "%BUILD_DIR%\"
copy "version_manager.py" "%BUILD_DIR%\"
copy "update_version.py" "%BUILD_DIR%\"
copy "requirements.txt" "%BUILD_DIR%\"
copy "pal.ico" "%BUILD_DIR%\"

echo.
echo 4. 複製 guide_popup 目錄...
xcopy "guide_popup" "%BUILD_DIR%\guide_popup\" /E /I

echo.
echo 5. 複製 EXCEL 目錄（範例檔案）...
xcopy "EXCEL" "%BUILD_DIR%\EXCEL\" /E /I

echo.
echo 6. 複製文檔檔案...
copy "README.md" "%BUILD_DIR%\"
copy "VERSION_USAGE.md" "%BUILD_DIR%\"
copy "AI_RECOMMENDATION_USAGE.md" "%BUILD_DIR%\"

echo.
echo 7. 安裝 PyInstaller（如果尚未安裝）...
pip install pyinstaller

echo.
echo 8. 使用 PyInstaller 打包...
cd %BUILD_DIR%
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
echo 9. 複製執行檔到發佈目錄...
copy "dist\%APP_NAME%.exe" "..\%DIST_DIR%\"
cd ..

echo.
echo 10. 複製額外檔案到發佈目錄...
copy "pal.ico" "%DIST_DIR%\"
copy "README.md" "%DIST_DIR%\"
copy "VERSION_USAGE.md" "%DIST_DIR%\"
copy "AI_RECOMMENDATION_USAGE.md" "%DIST_DIR%\"
xcopy "guide_popup" "%DIST_DIR%\guide_popup\" /E /I
xcopy "EXCEL" "%DIST_DIR%\EXCEL\" /E /I

echo.
echo 11. 創建啟動腳本...
echo @echo off > "%DIST_DIR%\start.bat"
echo echo 啟動 Error Code Comparer... >> "%DIST_DIR%\start.bat"
echo %APP_NAME%.exe >> "%DIST_DIR%\start.bat"
echo pause >> "%DIST_DIR%\start.bat"

echo.
echo 12. 創建版本資訊檔案...
echo Error Code Comparer v1.5.0 > "%DIST_DIR%\version.txt"
echo 建置日期: %date% %time% >> "%DIST_DIR%\version.txt"
echo 建置環境: Windows >> "%DIST_DIR%\version.txt"

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
