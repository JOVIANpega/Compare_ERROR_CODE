import PyInstaller.__main__
import os

# 獲取當前目錄
current_dir = os.path.dirname(os.path.abspath(__file__))

# 設置圖標路徑（如果有圖標的話）
# icon_path = os.path.join(current_dir, 'icon.ico')

# 設置打包參數
PyInstaller.__main__.run([
    'error_code_compare.py',  # 主程式文件
    '--name=ErrorCodeComparer',  # 生成的執行檔名稱
    '--onefile',  # 打包成單一執行檔
    '--windowed',  # 使用 GUI 模式
    '--clean',  # 清理臨時文件
    '--noconfirm',  # 不詢問確認
    '--hidden-import=pandas',  # 添加隱藏依賴
    '--hidden-import=openpyxl',
    '--hidden-import=tkinter',
    '--hidden-import=tkinter.ttk',
    '--hidden-import=tkinter.filedialog',
    '--hidden-import=tkinter.messagebox',
    '--collect-all=pandas',  # 收集所有 pandas 相關文件
    '--collect-all=openpyxl',
    '--hidden-import=numpy',
    '--hidden-import=numpy.random',
    '--hidden-import=numpy.random._pickle',
    '--hidden-import=numpy.random.bit_generator',
    # f'--icon={icon_path}',  # 如果有圖標的話，取消這行的註釋
    '--add-data=README.md;.',  # 添加額外的文件
]) 