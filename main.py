import pandas as pd
import pyautogui
import time
import tkinter as tk

import pyperclip

# --- 1. 配置信息 ---
BILL_FILE_PATH = '账单表格.xlsx'
# 搜索框的坐标，你需要手动获取
# 1. 运行 `print(pyautogui.position())`
# 2. 将鼠标移动到搜索框，然后复制打印出的坐标
SEARCH_BOX_COORDS = (100, 150)  # 示例坐标，请替换为你自己的

# --- 2. 核心函数 ---

def send_message_via_ui(manager_name, company_name, debt_amount, phone_number):
    """
    通过模拟按键发送消息的函数
    """
    print(f"正在为客户经理 {manager_name} 准备发送通知...")

    try:
        # 激活通讯录窗口（这一步可能需要手动完成或使用pygetwindow）
        # 假设当前窗口就是通讯录窗口

        # 1. 点击搜索框或按下快捷键
        # 模拟 Ctrl+F 快捷键
        pyautogui.hotkey('ctrl', 'f')
        time.sleep(1)  # 等待搜索框出现

        # 2. 输入客户经理名字
        pyautogui.write(str(phone_number))
        time.sleep(1)  # 等待搜索结果出现

        # 3. 模拟按 Enter 键进入聊天
        # 假设搜索结果的第一个就是正确的客户经理，直接按Enter
        pyautogui.press('enter')
        time.sleep(2)  # 等待聊天窗口加载

        # 4. 构造并输入消息内容
        message = f"你好，这是测试，{company_name} 客户已产生欠费，欠费金额为：¥{debt_amount}。请及时联系客户处理。谢谢！"
        pyperclip.copy(message)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)

        # 5. 模拟按 Enter 键发送消息
        pyautogui.press('enter')
        print(f"成功向 {manager_name} 发送消息。")
        time.sleep(2)  # 发送后等待一段时间

    except Exception as e:
        print(f"操作失败，错误：{e}")

# --- 3. 主逻辑 ---
def main():
    print("Start working!")
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(1)
    pyautogui.write("移动办公")
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(2)

    try:
        df = pd.read_excel(BILL_FILE_PATH)
    except FileNotFoundError:
        print(f"错误：未找到文件 {BILL_FILE_PATH}")
        return
    print("successfully load!")
    print(df)
    # 遍历每一行数据
    for index, row in df.iterrows():
        company_name = row['企业名称']
        debt_amount = row['欠费金额']
        manager_name = row['客户经理']
        phone_number = row['电话号码']

        send_message_via_ui(manager_name, company_name, debt_amount, phone_number)

if __name__ == "__main__":
    # 在运行前，请确保通讯录窗口处于激活状态
    # 建议在运行前打开通讯录并将其置于屏幕中心
    main()