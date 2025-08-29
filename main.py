import pandas as pd
import pyautogui, pyperclip, time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# --------------------------
# 全局变量
# --------------------------
root = tk.Tk()
root.title("欠费通知自动发送工具")
root.geometry("600x500")
file_path_var = tk.StringVar()
target_sheets = ["30天通报", "60天通报", "90天通报"]
region_contact = None
region_message = None

# --------------------------
# 核心函数
# --------------------------
def select_file():
    """选择 Excel 文件"""
    path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if path:
        file_path_var.set(path)
        file_label.config(text=path)

def log(msg):
    """向日志框输出信息"""
    log_box.insert(tk.END, msg + "\n")
    log_box.see(tk.END)  # 滚动到末尾
    print(msg)

def send_message_day(row, sheet_name):
    print(f"{sheet_name}: 发送信息给 补充客户经理 {row["补充客户经理"]}")
    manager = row["客户经理电话"]
    message = row["短信模板"]
    send_message(manager, message)
    if sheet_name != target_sheets[0]:
        print(f"{sheet_name}: 发送信息给 总监 {row["总监"]}")
        director = row["总监电话"]
        send_message(director, message)
    if sheet_name == target_sheets[2]:
        print(f"{sheet_name}: 发送信息给 分管领导 {row["分管领导"]}")
        leader = row["分管领导电话"]
        send_message(leader, message)

def send_message(phone_number, message):
    """模拟发送消息"""
    try:
        # 假设当前窗口就是通讯录窗口
        pyautogui.hotkey('ctrl', 'f')
        time.sleep(1)
        pyautogui.write(str(phone_number))
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        pyperclip.copy(message)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(1)
    except Exception as e:
        log(f"操作失败，错误：{e}")

def start_processing():
    """开始处理 Excel 文件"""
    path = file_path_var.get()
    if not path:
        messagebox.showwarning("提示", "请先选择 Excel 文件")
        return

    log("开始处理 Excel 文件...")
    try:
        sheets = pd.read_excel(path, sheet_name=None)
    except Exception as e:
        messagebox.showerror("错误", f"无法读取文件: {e}")
        return

    # 打开移动办公
    pyautogui.hotkey('win')
    time.sleep(1)
    pyperclip.copy("移动办公")
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(2)

    for sheet_name in sheets:
        if sheet_name not in target_sheets:
            continue

        log(f"处理 Sheet: {sheet_name}")
        try:
            df = sheets[sheet_name]
        except Exception as e:
            log(f"读取 Sheet {sheet_name} 失败: {e}")
            break
        print(df)

        for index, row in df.iterrows():
            try:
                send_message_day(row, sheet_name)
            except KeyError as e:
                log(f"缺少列: {e}")
            except Exception as e:
                log(f"处理第 {index+1} 行失败: {e}")

    log("所有 预警 发送完成！")

# --------------------------
# GUI
# --------------------------
root.geometry("600x500")

file_btn = tk.Button(root, text="选择 Excel 文件", command=select_file)
file_btn.pack(pady=5)

file_label = tk.Label(root, text="未选择文件", anchor="w")
file_label.pack(fill=tk.X, padx=10)

start_btn = tk.Button(root, text="开始处理", command=start_processing)
start_btn.pack(pady=10)

log_box = tk.Text(root, height=20, width=70)
log_box.pack(padx=10, pady=5)

root.mainloop()
