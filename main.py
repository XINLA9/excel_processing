import tkinter as tk
from tkinter import ttk
import sender_app
import excel_app   # 建议类名首字母大写，保持风格一致

def open_ocr():
    win = tk.Toplevel(root)   # 子窗口
    try:
        style = ttk.Style(win)
        style.theme_use('clam')
    except Exception:
        pass
    sender_app.OcrApp(win)

def open_excel():
    excel_app()

if __name__ == '__main__':
    root = tk.Tk()
    root.title("主入口")
    root.geometry("400x200")

    tk.Button(root, text="打开 OCR 工具", command=open_ocr).pack(pady=20)
    tk.Button(root, text="打开 Excel 工具", command=open_excel).pack(pady=20)

    root.mainloop()
