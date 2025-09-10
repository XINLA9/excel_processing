import os
import json
import time
import datetime
import difflib
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
import pyautogui
import pyperclip

from PIL import Image, ImageGrab, ImageOps, ImageFilter

# =============== 可选依赖：pytesseract ==================
# 需要先安装 Tesseract OCR 程序，并配置路径
# Windows 常见路径：C:\\Program Files\\Tesseract-OCR\\tesseract.exe
try:
    import pytesseract

    TESS_AVAILABLE = True
except Exception:
    pytesseract = None
    TESS_AVAILABLE = False

# ================== 全局默认配置 ========================
DEFAULT_CONFIG = {
    "excel_path": "",
    "target_sheets": ["30天通报", "60天通报", "90天通报"],
    "region_contact": None,  # (x1, y1, x2, y2)
    "region_message": None,  # (x1, y1, x2, y2)
    "tesseract_path": "C:\\Program Files\\Tesseract-OCR\\tesseract.exe",  # Tesseract
    "ocr_lang": "chi_sim",  # 简体中文
    "ocr_threshold": 0.70,  # 相似度阈值
    "max_retries": 3,  # 发送失败重试次数
    "post_send_wait_sec": 2.0,  # 按回车后等待消息渲染时间
    "search_wait_sec": 2.0,  # 搜索/切换联系人等待时间
    "use_ocr": True,
}


# ================== 工具 & OCR ===========================
# 计算相似度
def ratio(a: str, b: str) -> float:
    """计算两个字符串的相似度"""
    a = (a or '').strip()
    b = (b or '').strip()
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


# 对图像进行预处理
def preprocess_for_ocr(img: Image.Image) -> Image.Image:
    """基础预处理：灰度 -> 自适应对比度 -> 轻度锐化，提升 OCR 稳定性。"""
    g = ImageOps.grayscale(img)
    # 轻度增强对比
    g = ImageOps.autocontrast(g)
    # 轻度锐化
    g = g.filter(ImageFilter.SHARPEN)
    return g


# 抓取区域
def grab_region(region):
    """根据坐标 (x1, y1, x2, y2) 截取屏幕区域"""
    if not region:
        return None
    box = (int(region[0]), int(region[1]), int(region[2]), int(region[3]))
    return ImageGrab.grab(bbox=box)


# ocr提取文字
def ocr_text_from_region(region, lang='chi_sim') -> str:
    """对指定区域进行 OCR 文本识别"""
    print("进行ocr")
    if not TESS_AVAILABLE or pytesseract is None:
        return ""
    img = grab_region(region)
    if img is None:
        return ""
    img = preprocess_for_ocr(img)
    try:
        text = pytesseract.image_to_string(img, lang=lang)
        return (text or '').strip()
    except Exception as e:
        print(f"OCR 失败: {e}")
        return ""


# ================== 可视化截图选区 =======================
# 截图GUI类
class ScreenCapture:
    """用于截图选区的 GUI 类"""

    def __init__(self):
        self.start_x = None
        self.start_y = None
        self.rect_id = None
        self.region = None

        self.root = tk.Tk()
        self.root.attributes('-fullscreen', True)
        # Windows: 置顶 + 透明蒙版
        self.root.attributes('-alpha', 0.30)
        self.root.configure(bg='gray')

        self.canvas = tk.Canvas(self.root, cursor="cross", bg="gray", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.canvas.bind("<ButtonPress-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)

        self.root.mainloop()

    def on_mouse_down(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.rect_id = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y,
                                                    outline='red', width=2)

    def on_mouse_drag(self, event):
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        if self.rect_id is not None:
            self.canvas.coords(self.rect_id, self.start_x, self.start_y, cur_x, cur_y)

    def on_mouse_up(self, event):
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)
        x1, y1 = int(min(self.start_x, end_x)), int(min(self.start_y, end_y))
        x2, y2 = int(max(self.start_x, end_x)), int(max(self.start_y, end_y))
        # 避免用户误点
        if abs(x2 - x1) < 5 or abs(y2 - y1) < 5:
            self.region = None
        else:
            self.region = (x1, y1, x2, y2)
        self.root.quit()
        self.root.destroy()


def select_region_blocking() -> tuple:
    """启动截图选区程序并返回选区坐标"""
    sc = ScreenCapture()
    return sc.region


# ================== 发送 & 验证 ==========================
# 发送器类
class Sender:
    """负责执行消息发送和OCR验证的类"""

    def __init__(self, cfg, log_func):
        self.cfg = cfg
        self.log = log_func
        # PyAutoGUI 设置
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.05
        # Tesseract 路径
        if TESS_AVAILABLE and cfg.get('tesseract_path'):
            try:
                pytesseract.pytesseract.tesseract_cmd = cfg['tesseract_path']
            except Exception as e:
                self.log(f"设置 Tesseract 路径失败: {e}")

    # ---------- OCR 双重验证 ----------
    def verify_contact(self, expected_name: str) -> bool:
        """验证当前联系人是否正确"""
        if not self.cfg.get('use_ocr'):
            return True
        if not expected_name:
            # 没有名字可比对就跳过
            return True
        region = self.cfg.get('region_contact')
        if not region:
            self.log("未设置联系人 OCR 区域，跳过联系人校验")
            return True
        text = ocr_text_from_region(region, self.cfg.get('ocr_lang', 'chi_sim'))
        self.log(f"[OCR-联系人] 期望: {expected_name} | 识别: {text}")
        if not text:
            return False
        th = float(self.cfg.get('ocr_threshold', 0.7))
        # 直接包含 或 相似度
        return (expected_name in text) or (ratio(expected_name, text) >= th)

    def verify_message(self, message: str) -> bool:
        """验证消息是否发送成功并显示在界面上"""
        if not self.cfg.get('use_ocr'):
            return True
        region = self.cfg.get('region_message')
        if not region:
            self.log("未设置消息 OCR 区域，跳过消息校验")
            return True
        text = ocr_text_from_region(region, self.cfg.get('ocr_lang', 'chi_sim'))
        # 为了避免长消息被 OCR 丢字，使用片段匹配（取前后片段）
        frag_head = (message or '')[:14]
        frag_tail = (message or '')[-14:]
        self.log(f"[OCR-消息] 识别到: {text}")
        if not text:
            return False
        th = float(self.cfg.get('ocr_threshold', 0.7))
        rules = [
            (message and message in text),
            (frag_head and frag_head in text),
            (frag_tail and frag_tail in text),
            ratio(message, text) >= th,
        ]
        return any(rules)

    # ---------- 发送单条消息 ----------
    def send_one(self, phone_number: str, message: str, contact_name: str = None) -> bool:
        """模拟按键操作，发送单条消息"""
        try:
            # 搜索联系人（假设 Ctrl+F 能聚焦搜索框）
            pyautogui.hotkey('ctrl', 'f')
            time.sleep(0.8)
            pyautogui.typewrite(str(phone_number), interval=0.02)
            time.sleep(0.8)
            pyautogui.press('enter')
            time.sleep(self.cfg.get('search_wait_sec', 2.0))

            # 联系人 OCR 校验
            if not self.verify_contact(contact_name or str(phone_number)):
                self.log(f"联系人校验失败 -> 期望: {contact_name or phone_number}")
                return False

            # 粘贴并发送
            pyperclip.copy(message)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.2)
            pyautogui.press('enter')
            time.sleep(self.cfg.get('post_send_wait_sec', 2.0))

            # 消息 OCR 校验
            if not self.verify_message(message):
                return False

            return True
        except Exception as e:
            self.log(f"发送异常: {e}")
            return False

    # ---------- 发送（带重试） ----------
    def send_with_retry(self, phone_number: str, message: str, contact_name: str = None) -> bool:
        """发送单条消息，失败时重试"""
        retries = int(self.cfg.get('max_retries', 3))
        for i in range(1, retries + 1):
            ok = self.send_one(phone_number, message, contact_name)
            if ok:
                self.log(f"✅ 发送成功 -> {contact_name or phone_number}")
                return True
            else:
                self.log(f"❌ 验证失败/异常，第 {i} 次尝试")
                time.sleep(0.8)
        self.log(f"🛑 发送失败 -> {contact_name or phone_number}")
        return False


# ================== GUI 主程序 ===========================

class App:
    """主 GUI 应用程序"""

    def __init__(self, root):
        self.root = root
        self.root.title("欠费通知自动发送工具")
        self.root.geometry("820x700")
        self.cfg = DEFAULT_CONFIG.copy()
        self.last_failed_file_path = None

        # Tesseract 检查
        if TESS_AVAILABLE and self.cfg.get('tesseract_path'):
            try:
                pytesseract.pytesseract.tesseract_cmd = self.cfg['tesseract_path']
            except Exception as e:
                print(f"Tesseract 路径设置失败: {e}")

        self.build_ui()
        self.sender = Sender(self.cfg, self.log)
        self.update_button_states(False)

    # ---------- UI ----------
    def build_ui(self):
        """构建主界面"""
        pad = 8

        # Excel 选择
        frm_file = ttk.LabelFrame(self.root, text="Excel 文件")
        frm_file.pack(fill=tk.X, padx=pad, pady=pad)
        self.var_excel = tk.StringVar(value="")
        ttk.Entry(frm_file, textvariable=self.var_excel).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=pad, pady=pad)
        ttk.Button(frm_file, text="选择...", command=self.select_excel).pack(side=tk.LEFT, padx=pad)

        # OCR 与 Tesseract
        frm_ocr = ttk.LabelFrame(self.root, text="OCR 设置")
        frm_ocr.pack(fill=tk.X, padx=pad, pady=pad)
        self.var_use_ocr = tk.BooleanVar(value=bool(self.cfg.get('use_ocr', True)))
        ttk.Checkbutton(frm_ocr, text="启用 OCR 验证（联系人 + 消息）", variable=self.var_use_ocr).grid(row=0, column=0,
                                                                                                      sticky='w',
                                                                                                      padx=pad,
                                                                                                      pady=pad)

        ttk.Label(frm_ocr, text="相似度阈值").grid(row=0, column=2, sticky='e', padx=pad)
        self.var_threshold = tk.DoubleVar(value=float(self.cfg.get('ocr_threshold', 0.70)))
        ttk.Entry(frm_ocr, textvariable=self.var_threshold, width=8).grid(row=0, column=3, sticky='w', padx=pad)

        # OCR 区域选择
        frm_region = ttk.LabelFrame(self.root, text="选择 OCR 区域")
        frm_region.pack(fill=tk.X, padx=pad, pady=pad)
        ttk.Button(frm_region, text="选择联系人区域", command=self.choose_contact_region).grid(row=0, column=0,
                                                                                               padx=pad, pady=pad)
        ttk.Button(frm_region, text="选择消息区域", command=self.choose_message_region).grid(row=0, column=1, padx=pad,
                                                                                             pady=pad)
        ttk.Button(frm_region, text="预览联系人OCR", command=self.preview_contact_ocr).grid(row=0, column=2, padx=pad,
                                                                                            pady=pad)
        ttk.Button(frm_region, text="预览消息OCR", command=self.preview_message_ocr).grid(row=0, column=3, padx=pad,
                                                                                          pady=pad)

        # 重试 & 等待时间
        frm_retry = ttk.LabelFrame(self.root, text="发送策略")
        frm_retry.pack(fill=tk.X, padx=pad, pady=pad)
        ttk.Label(frm_retry, text="最大重试次数").grid(row=0, column=0, sticky='w', padx=pad)
        self.var_retries = tk.IntVar(value=int(self.cfg.get('max_retries', 3)))
        ttk.Entry(frm_retry, textvariable=self.var_retries, width=6).grid(row=0, column=1, sticky='w', padx=pad)

        ttk.Label(frm_retry, text="搜索等待(s)").grid(row=0, column=2, sticky='e', padx=pad)
        self.var_search_wait = tk.DoubleVar(value=float(self.cfg.get('search_wait_sec', 2.0)))
        ttk.Entry(frm_retry, textvariable=self.var_search_wait, width=6).grid(row=0, column=3, sticky='w', padx=pad)

        ttk.Label(frm_retry, text="发送后等待(s)").grid(row=0, column=4, sticky='e', padx=pad)
        self.var_post_wait = tk.DoubleVar(value=float(self.cfg.get('post_send_wait_sec', 2.0)))
        ttk.Entry(frm_retry, textvariable=self.var_post_wait, width=6).grid(row=0, column=5, sticky='w', padx=pad)

        # 按钮区
        frm_btn = ttk.Frame(self.root)
        frm_btn.pack(fill=tk.X, padx=pad, pady=pad)
        ttk.Button(frm_btn, text="开始处理", command=self.start_processing).pack(side=tk.LEFT, padx=pad)
        self.btn_resend = ttk.Button(frm_btn, text="二次发送", command=self.start_reprocessing, state='disabled')
        self.btn_resend.pack(side=tk.LEFT, padx=pad)
        self.btn_open_failed = ttk.Button(frm_btn, text="未发送名单", command=self.open_failed_file, state='disabled')
        self.btn_open_failed.pack(side=tk.LEFT)
        ttk.Button(frm_btn, text="使用说明", command=self.show_instructions).pack(side=tk.LEFT, padx=pad)

        # 日志
        frm_log = ttk.LabelFrame(self.root, text="日志")
        frm_log.pack(fill=tk.BOTH, expand=True, padx=pad, pady=pad)
        self.txt_log = tk.Text(frm_log, height=18)
        self.txt_log.pack(fill=tk.BOTH, expand=True, padx=pad, pady=pad)

    def update_button_states(self, has_failed_file: bool):
        """根据是否有失败文件来更新按钮状态"""
        if has_failed_file:
            self.btn_resend.config(state='normal')
            self.btn_open_failed.config(state='normal')
        else:
            self.btn_resend.config(state='disabled')
            self.btn_open_failed.config(state='disabled')

    def open_failed_file(self):
        """打开最近生成的失败文件"""
        if self.last_failed_file_path and os.path.exists(self.last_failed_file_path):
            try:
                os.startfile(self.last_failed_file_path)
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件: {e}")
                self.log(f"无法打开文件: {e}")
        else:
            messagebox.showwarning("提示", "没有可打开的失败文件。")

    # 新增的使用说明方法
    def show_instructions(self):
        """弹出使用说明窗口"""
        instructions = (
            "使用说明\n"
            "本工具通过模拟鼠标键盘操作，实现从 Excel 读取数据并自动发送预警通知。\n\n"
            "步骤：\n"
            "1. 配置：如果您需要启用 OCR 验证，请先安装 Tesseract OCR 程序，并在“Tesseract 路径”中配置其可执行文件路径。\n"
            "2. 截图选区：点击“选择联系人区域”和“选择消息区域”按钮，分别框选您通讯软件中【联系人姓名】和【已发送消息】所在的屏幕位置。此步骤至关重要，决定了 OCR 验证的准确性。\n"
            "3. 导入数据：点击“选择...”按钮，导入您的 Excel 数据文件。请确保 Excel 表格的表头名称符合程序预期：如“客户经理电话”、“短信模板”、“总监电话”、“分管领导电话”等。\n"
            "4. 开始：点击“开始处理”，程序会自动打开“移动办公”应用，并按行读取 Excel 数据，依次发送通知。\n\n"
            "重要提示：\n"
            "* 在程序运行时，请勿移动鼠标或操作键盘。\n"
            "* 在开始处理前，请确保您的通讯软件已打开，并处于可以搜索联系人的状态。\n"
            "* OCR 验证失败时，程序会自动重试。\n"
        )
        messagebox.showinfo("使用说明", instructions)

    def log(self, msg: str):
        """在日志文本框中打印消息"""
        try:
            self.txt_log.insert(tk.END, str(msg) + "\n")
            self.txt_log.see(tk.END)
            self.root.update_idletasks()  # 确保日志能立即显示
        except Exception:
            print(msg)

    # ---------- 事件 ----------
    def select_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx;*.xls")])
        if path:
            self.var_excel.set(path)
            # 在选择新文件时，清空之前保存的失败文件路径
            self.last_failed_file_path = None
            self.update_button_states(False)

    def choose_contact_region(self):
        messagebox.showinfo("提示", "请选取【联系人名称】所在区域")
        region = select_region_blocking()
        if region:
            self.cfg['region_contact'] = region
            self.log(f"联系人区域: {region}")
        else:
            self.log("联系人区域选择取消或无效")

    def choose_message_region(self):
        messagebox.showinfo("提示", "请选取【消息内容】所在区域")
        region = select_region_blocking()
        if region:
            self.cfg['region_message'] = region
            self.log(f"消息区域: {region}")
        else:
            self.log("消息区域选择取消或无效")

    def preview_contact_ocr(self):
        if not self.cfg.get('region_contact'):
            messagebox.showwarning("提示", "请先选择联系人区域")
            return
        if not TESS_AVAILABLE:
            messagebox.showwarning("提示", "未安装 pytesseract")
            return
        text = ocr_text_from_region(self.cfg['region_contact'], self.cfg['ocr_lang'])
        self.log(f"[预览-联系人] -> {text}")

    def preview_message_ocr(self):
        if not self.cfg.get('region_message'):
            messagebox.showwarning("提示", "请先选择消息区域")
            return
        if not TESS_AVAILABLE:
            messagebox.showwarning("提示", "未安装 pytesseract")
            return
        text = ocr_text_from_region(self.cfg['region_message'], self.cfg['ocr_lang'])
        self.log(f"[预览-消息] -> {text}")

    # ---------- 业务主流程 ----------
    def start_processing(self):
        """开始处理主流程，发送并记录失败项"""
        # 获取最新的配置（虽然已无文件加载，但仍保留逻辑以处理用户UI修改）
        self.cfg['ocr_threshold'] = float(self.var_threshold.get())
        self.cfg['max_retries'] = int(self.var_retries.get())
        self.cfg['search_wait_sec'] = float(self.var_search_wait.get())
        self.cfg['post_send_wait_sec'] = float(self.var_post_wait.get())
        self.cfg['use_ocr'] = bool(self.var_use_ocr.get())

        excel = self.var_excel.get()
        if not excel or not os.path.exists(excel):
            messagebox.showerror("错误", "请先选择有效的 Excel 文件")
            return

        try:
            sheets = pd.read_excel(excel, sheet_name=None)
        except Exception as e:
            messagebox.showerror("错误", f"无法读取 Excel: {e}")
            return

        # 打开“移动办公”应用（按需调整：这里示例 Win 键呼出搜索）
        try:
            pyautogui.hotkey('win')
            time.sleep(0.8)
            pyperclip.copy("移动办公")
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.6)
            pyautogui.press('enter')
            time.sleep(1.0)
            pyautogui.press('enter')  # 有的系统首次需要确认
            time.sleep(1.5)
        except Exception as e:
            self.log(f"尝试启动应用失败（可忽略，若已打开）: {e}")

        tgt = set(self.cfg.get('target_sheets', []))
        total, okcnt, failcnt = 0, 0, 0
        failed_sends = []

        for sheet_name, df in sheets.items():
            if sheet_name not in tgt:
                continue
            self.log(f"==== 处理 Sheet: {sheet_name} ====")
            if not isinstance(df, pd.DataFrame) or df.empty:
                self.log(f"Sheet {sheet_name} 为空，跳过")
                continue

            # 逐行发送
            for idx, row in df.iterrows():
                try:
                    is_failed = False

                    # 1) 发送客户经理
                    phone_manager = str(row.get('客户经理电话', '')).strip()
                    msg = str(row.get('短信模板', '')).strip()
                    contact_name = str(row.get('补充客户经理', '')).strip() or str(row.get('客户经理', '')).strip()
                    if phone_manager and msg:
                        total += 1
                        ok = self.sender.send_with_retry(phone_manager, msg, contact_name=contact_name or None)
                        if ok:
                            okcnt += 1
                        else:
                            failcnt += 1
                            is_failed = True

                    # 2) 发送总监（除 30 天）
                    if sheet_name != self.cfg.get('target_sheets', [])[0]:
                        phone_dir = str(row.get('总监电话', '')).strip()
                        name_dir = str(row.get('总监', '')).strip()
                        if phone_dir and msg:
                            total += 1
                            ok = self.sender.send_with_retry(phone_dir, msg, contact_name=name_dir or None)
                            if ok:
                                okcnt += 1
                            else:
                                failcnt += 1
                                is_failed = True

                    # 3) 发送分管领导（仅 90 天）
                    if sheet_name == self.cfg.get('target_sheets', [None, None, ''])[2]:
                        phone_lead = str(row.get('分管领导电话', '')).strip()
                        name_lead = str(row.get('分管领导', '')).strip()
                        if phone_lead and msg:
                            total += 1
                            ok = self.sender.send_with_retry(phone_lead, msg, contact_name=name_lead or None)
                            if ok:
                                okcnt += 1
                            else:
                                failcnt += 1
                                is_failed = True

                    if is_failed:
                        failed_sends.append(row)

                except Exception as e:
                    self.log(f"行 {idx + 1} 处理异常: {e}")
                    failed_sends.append(row)  # 异常也计入失败

        self.log(f"完成。总计: {total} | 成功: {okcnt} | 失败: {failcnt}")
        messagebox.showinfo("发送结果", f"发送完成！\n成功: {okcnt} 条\n失败: {failcnt} 条")

        # 记录失败项到新的 Excel 文件
        if failed_sends:
            df_failed = pd.DataFrame(failed_sends)
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            failed_file_path = f"failed_sends_{timestamp}.xlsx"
            df_failed.to_excel(failed_file_path, index=False)
            self.last_failed_file_path = failed_file_path
            self.update_button_states(True)
            self.log(f"⚠️ {failcnt} 条发送失败，已保存至文件: {os.path.abspath(failed_file_path)}")
        else:
            self.last_failed_file_path = None
            self.update_button_states(False)
            self.log("🎉 所有信息发送成功，没有失败记录")

    def process_from_file(self, file_path):
        """通用处理函数，用于处理任何给定的 Excel 文件"""
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("错误", "请选择有效的 Excel 文件进行二次发送")
            return

        try:
            df = pd.read_excel(file_path)
            if df.empty:
                messagebox.showinfo("提示", "选择的 Excel 文件为空")
                self.log("文件为空，无需二次发送")
                return
        except Exception as e:
            messagebox.showerror("错误", f"无法读取 Excel: {e}")
            return

        self.log(f"==== 正在进行二次发送 (文件: {os.path.basename(file_path)}) ====")
        total, okcnt, failcnt = 0, 0, 0

        for idx, row in df.iterrows():
            # 简化二次发送逻辑，只按电话和模板发送
            phone = str(row.get('客户经理电话', '')).strip()
            msg = str(row.get('短信模板', '')).strip()
            contact_name = str(row.get('补充客户经理', '')).strip() or str(row.get('客户经理', '')).strip()

            if phone and msg:
                total += 1
                ok = self.sender.send_with_retry(phone, msg, contact_name=contact_name or None)
                okcnt += 1 if ok else 0
                failcnt += 0 if ok else 1

        self.log(f"二次发送完成。总计: {total} | 成功: {okcnt} | 失败: {failcnt}")
        messagebox.showinfo("二次发送结果", f"二次发送完成！\n成功: {okcnt} 条\n失败: {failcnt} 条")

    def start_reprocessing(self):
        """启动二次发送流程，先选择文件"""
        if self.last_failed_file_path and os.path.exists(self.last_failed_file_path):
            self.process_from_file(self.last_failed_file_path)
        else:
            messagebox.showwarning("提示", "没有可供二次发送的失败文件，请先运行主程序。")


if __name__ == '__main__':
    root = tk.Tk()
    style = ttk.Style()
    try:
        style.theme_use('clam')
    except Exception:
        pass
    app = App(root)
    root.mainloop()