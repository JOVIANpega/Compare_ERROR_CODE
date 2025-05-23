import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import os
import sys

def get_resource_path(relative_path, for_write=False):
    # 永遠用 EXE 目錄（py模式用__file__目錄）
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, relative_path)

def show_guide(root, setup_file, guide_title="新手導覽"):
    print(f"[DEBUG] show_guide: start with setup_file={setup_file}")
    # 讀取 ShowGuide 設定與內容
    show_guide_flag = True
    picture_number = 1
    guide_contents = {}
    
    # 修正 setup.txt 路徑
    setup_path = get_resource_path(setup_file)
    print(f"[DEBUG] setup.txt path: {setup_path}")
    
    if os.path.exists(setup_path):
        print("[DEBUG] setup.txt exists")
        with open(setup_path, 'r', encoding='utf-8') as f:
            for line in f:
                if line.strip().startswith('ShowGuide='):
                    value = line.strip().split('=', 1)[1]
                    print(f"[DEBUG] ShowGuide value: {value}")
                    if value == '0':
                        show_guide_flag = False
                if line.strip().startswith('ShowGuidePictureNumber='):
                    try:
                        picture_number = int(line.strip().split('=', 1)[1])
                        print(f"[DEBUG] ShowGuidePictureNumber: {picture_number}")
                    except:
                        picture_number = 1
                if line.strip().startswith('ShowGuideContent_'):
                    k, v = line.strip().split('=', 1)
                    try:
                        idx = int(k.split('_')[1])
                        guide_contents[idx] = v
                    except:
                        pass
    else:
        print(f"[DEBUG] setup.txt not found at {setup_path}")
    
    print(f"[DEBUG] show_guide_flag: {show_guide_flag}")
    print(f"[DEBUG] picture_number: {picture_number}")
    
    if not show_guide_flag or picture_number < 1:
        print("[DEBUG] Not showing guide due to flag or picture number")
        return

    print("[DEBUG] Creating guide window")
    guide_win = tk.Toplevel(root)
    guide_win.title(guide_title)
    
    # 設定較小的視窗大小
    guide_win.geometry("800x600")  # 設定視窗大小
    guide_win.resizable(True, True)
    
    # 置中
    ws = guide_win.winfo_screenwidth()
    hs = guide_win.winfo_screenheight()
    x = (ws // 2) - (800 // 2)
    y = (hs // 2) - (600 // 2)
    guide_win.geometry(f"800x600+{x}+{y}")
    guide_win.grab_set()
    guide_win.transient(root)

    frame = ttk.Frame(guide_win, padding=20)
    frame.pack(fill=tk.BOTH, expand=True)

    # 圖片顯示區
    img_label = ttk.Label(frame, anchor="center")
    img_label.pack(pady=10, fill=tk.BOTH, expand=True)

    text_label = ttk.Label(frame, text="", font=("Microsoft JhengHei", 12), wraplength=700, justify="left")
    text_label.pack(pady=(0, 20))

    # 頁數顯示
    page_label = ttk.Label(frame, text="", font=("Microsoft JhengHei", 10))
    page_label.pack()

    # 下方控制區
    bottom = ttk.Frame(guide_win, padding=(20, 0, 20, 20))
    bottom.pack(side="bottom", fill="x")

    var = tk.IntVar()
    chk = ttk.Checkbutton(bottom, text="下次不再顯示此導覽", variable=var)
    chk.pack(side="left")

    # 建立大字體 style
    style = ttk.Style()
    style.configure("Big.TButton", font=("Microsoft JhengHei", 16), padding=10)

    btn_prev = ttk.Button(bottom, text="上一步", width=16, style="Big.TButton")
    btn_next = ttk.Button(bottom, text="下一步", width=16, style="Big.TButton")
    btn_finish = ttk.Button(bottom, text="我知道了", width=16, style="Big.TButton")

    btn_prev.pack(side="right", padx=5)
    btn_next.pack(side="right", padx=5)

    # 載入所有圖片，並根據螢幕大小縮放
    images = []
    for i in range(1, picture_number + 1):
        img_path = get_resource_path(os.path.join("guide_popup", f"guide{i}.png"))
        if not os.path.exists(img_path):
            img_path = get_resource_path(os.path.join("guide_popup", f"guide{i}.jpg"))
        if os.path.exists(img_path):
            img = Image.open(img_path)
            # 依螢幕大小縮放，保留比例
            img.thumbnail((700, 400), Image.LANCZOS)
            images.append(ImageTk.PhotoImage(img))
        else:
            from PIL import ImageDraw
            blank = Image.new("RGB", (700, 400), (240, 240, 240))
            d = ImageDraw.Draw(blank)
            d.text((int(700/2-80), int(400/2)), f"No guide{i}.png", fill=(128, 128, 128))
            images.append(ImageTk.PhotoImage(blank))

    current_page = [0]

    def update_page():
        idx = current_page[0]
        img_label.config(image=images[idx], anchor="center")
        text = guide_contents.get(idx+1, f"第 {idx+1} 頁說明未設定")
        text_label.config(text=text.replace("\\n", "\n"))
        page_label.config(text=f"第 {idx+1} / {picture_number} 頁")
        btn_prev["state"] = tk.NORMAL if idx > 0 else tk.DISABLED
        btn_next.pack_forget()
        btn_finish.pack_forget()
        if idx < picture_number - 1:
            btn_next.pack(side="right", padx=5)
        else:
            btn_finish.pack(side="right", padx=5)

    def go_prev():
        if current_page[0] > 0:
            current_page[0] -= 1
            update_page()

    def go_next():
        if current_page[0] < picture_number - 1:
            current_page[0] += 1
            update_page()

    def close_guide():
        if var.get():
            # 設定 ShowGuide=0
            lines = []
            # 強制用 setup_path（主程式實際用的 setup.txt 路徑）
            if os.path.exists(setup_path):
                with open(setup_path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
            found = False
            for i, line in enumerate(lines):
                if line.startswith('ShowGuide='):
                    lines[i] = 'ShowGuide=0\n'
                    found = True
            if not found:
                lines.append('ShowGuide=0\n')
            # 寫入主程式實際用的 setup.txt
            with open(setup_path, 'w', encoding='utf-8') as f:
                f.writelines(lines)
        guide_win.destroy()

    btn_prev.config(command=go_prev)
    btn_next.config(command=go_next)
    btn_finish.config(command=close_guide)
    guide_win.protocol("WM_DELETE_WINDOW", close_guide)

    update_page()
    root.wait_window(guide_win) 