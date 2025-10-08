@echo off
echo 打包 Error Code Comparer...

REM 清理舊檔案
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

REM 安裝 PyInstaller
pip install pyinstaller

REM 打包
pyinstaller --onefile --windowed --icon=pal.ico --name=ErrorCodeComparer main.py

REM 複製必要檔案
mkdir dist\guide_popup
mkdir dist\EXCEL
xcopy guide_popup dist\guide_popup\ /E /I
xcopy EXCEL dist\EXCEL\ /E /I
copy VERSION.py dist\
copy README.md dist\

echo 打包完成！執行檔在 dist 目錄中
pause
