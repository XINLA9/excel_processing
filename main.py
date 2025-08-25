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

def send_message_via_ui(manager_name, company_name, debt_amount, phone_number):
    """模拟发送消息"""
    log(f"准备发送给客户经理 {manager_name} ...")

    try:
        # 假设当前窗口就是通讯录窗口
        pyautogui.hotkey('ctrl', 'f')
        time.sleep(1)
        pyautogui.write(str(phone_number))
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)

        message = f"你好，这是测试，{company_name} 客户已产生欠费，欠费金额为：¥{debt_amount}。请及时联系客户处理。谢谢！"
        pyperclip.copy(message)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)
        pyautogui.press('enter')

        log(f"成功向 {manager_name} 发送消息。")
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
        xls = pd.ExcelFile(path)
    except Exception as e:
        messagebox.showerror("错误", f"无法读取文件: {e}")
        return

    for sheet_name in xls.sheet_names:
        log(f"处理 Sheet: {sheet_name}")
        try:
            df = pd.read_excel(path, sheet_name=sheet_name)
        except Exception as e:
            log(f"读取 Sheet {sheet_name} 失败: {e}")
            continue

        for index, row in df.iterrows():
            try:
                send_message_via_ui(
                    row['客户经理'],
                    row['企业名称'],
                    row['欠费金额'],
                    row['电话号码']
                )
            except KeyError as e:
                log(f"缺少列: {e}")
            except Exception as e:
                log(f"处理第 {index+1} 行失败: {e}")

    log("所有 Sheet 处理完成！")

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
