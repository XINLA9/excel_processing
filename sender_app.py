import os
import sys
import time
import difflib
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
# ssh

import pandas as pd
import pyautogui
import pyperclip
import openpyxl

from PIL import Image, ImageGrab, ImageOps, ImageFilter

# =============== 可选依赖：pytesseract ==================
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
    "region_contact": None,
    "region_message": None,
    "click_point": None,
    "tesseract_path": "C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
    "ocr_lang": "chi_sim",
    "ocr_threshold": 0.70,
    "max_retries": 1,
    "post_send_wait_sec": 2.0,
    "search_wait_sec": 2.0,
    "use_ocr": True,
    "use_click": True,
}


# ================== OCR & 截图管理类 ===========================
class OCRManager:
    def __init__(self, tesseract_path: str = None):
        self.tesseract_available = TESS_AVAILABLE
        if self.tesseract_available and tesseract_path:
            try:
                pytesseract.pytesseract.tesseract_cmd = tesseract_path
            except Exception as e:
                messagebox.showerror("Tesseract 错误", f"设置 Tesseract 路径失败: {e}")
                self.tesseract_available = False
                sys.exit(1)

        # 检查 Tesseract 是否真正可用
        if self.tesseract_available:
            try:
                _ = pytesseract.get_tesseract_version()
            except Exception:
                messagebox.showerror("Tesseract 未安装", "未检测到 Tesseract，请使用软件包的文件进行安装。")
                self.tesseract_available = False
                sys.exit(1)

        # 检查语言包是否存在
        lang = "chi_sim"
        available_langs = pytesseract.get_languages(config='')
        if lang not in available_langs:
            messagebox.showerror("缺少语言包", f"缺少语言包 '{lang}'，请将“chi_sim.traineddata” "
                                               r"文件放置到 “C:\Program Files\Tesseract-OCR\tessdata” 文件夹下。")
            sys.exit(1)


    def _preprocess_for_ocr(self, img: Image.Image) -> Image.Image:
        g = ImageOps.grayscale(img)
        g = ImageOps.autocontrast(g)
        g = g.filter(ImageFilter.SHARPEN)
        return g

    def _grab_region(self, region):
        if not region:
            return None
        box = (int(region[0]), int(region[1]), int(region[2]), int(region[3]))
        return ImageGrab.grab(bbox=box)

    def recognize_text(self, region, lang='chi_sim') -> str:
        if not self.tesseract_available:
            return ""
        img = self._grab_region(region)
        if img is None:
            return ""
        img = self._preprocess_for_ocr(img)
        try:
            text = pytesseract.image_to_string(img, lang=lang)
            cleaned_text = (text or '').replace(' ', '').replace('\n', '').strip()
            return cleaned_text
        except Exception as e:
            print(f"OCR 识别失败: {e}")
            return ""

    @staticmethod
    def ratio(a: str, b: str) -> float:
        a = (a or '').strip()
        b = (b or '').strip()
        if not a or not b:
            return 0.0
        return difflib.SequenceMatcher(None, a, b).ratio()

    def select_region_gui(self) -> tuple:
        capture_app = self._ScreenCaptureGUI()
        return capture_app.region

    def select_point_gui(self) -> tuple:
        capture_app = self._ClickPointGUI()
        return capture_app.point

    class _ScreenCaptureGUI:
        def __init__(self):
            self.start_x = None
            self.start_y = None
            self.rect_id = None
            self.region = None

            self.root = tk.Tk()
            self.root.attributes('-fullscreen', True)
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
            if abs(x2 - x1) < 5 or abs(y2 - y1) < 5:
                self.region = None
            else:
                self.region = (x1, y1, x2, y2)
            self.root.quit()
            self.root.destroy()

    class _ClickPointGUI:
        def __init__(self):
            self.point = None
            self.root = tk.Tk()
            self.root.attributes('-fullscreen', True)
            self.root.attributes('-alpha', 0.10)  # 透明度可以调低
            self.root.configure(bg='gray')
            self.canvas = tk.Canvas(self.root, cursor="cross", bg="gray", highlightthickness=0)
            self.canvas.pack(fill=tk.BOTH, expand=True)
            self.canvas.bind("<ButtonPress-1>", self.on_mouse_down)
            self.root.mainloop()

        def on_mouse_down(self, event):
            x = self.canvas.canvasx(event.x)
            y = self.canvas.canvasy(event.y)
            self.point = (x, y)
            self.root.quit()
            self.root.destroy()

# ================== 发送 & 验证 ==========================
class Sender:
    def __init__(self, cfg, log_func, ocr_manager: OCRManager):
        self.cfg = cfg
        self.log = log_func
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.05
        self.ocr_manager = ocr_manager

    def verify_contact(self, expected_name: str) -> bool:
        if not self.cfg.get('use_ocr'):
            return True
        if not expected_name:
            return True
        region = self.cfg.get('region_contact')
        if not region:
            self.log("未设置联系人 OCR 区域，跳过联系人校验")
            return True
        text = self.ocr_manager.recognize_text(region, self.cfg.get('ocr_lang', 'chi_sim'))
        self.log(f"[OCR-联系人] 期望: {expected_name} | 识别: {text}")
        if not text:
            return False
        th = float(self.cfg.get('ocr_threshold', 0.7))
        return (expected_name in text) or (self.ocr_manager.ratio(expected_name, text) >= th)

    def verify_message(self, message: str) -> bool:
        if not self.cfg.get('use_ocr'):
            return True
        region = self.cfg.get('region_message')
        if not region:
            self.log("未设置消息 OCR 区域，跳过消息校验")
            return True
        text = self.ocr_manager.recognize_text(region, self.cfg.get('ocr_lang', 'chi_sim'))
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
            self.ocr_manager.ratio(message, text) >= th,
        ]
        return any(rules)

    def send_one(self, phone_number: str, message: str, contact_name: str = None) -> bool:
        try:
            pyautogui.hotkey('ctrl', 'f')
            time.sleep(0.8)
            pyautogui.typewrite(str(phone_number), interval=0.02)
            time.sleep(1.5)
            if self.cfg.get('use_click'):
                x, y = self.cfg['click_point']
                pyautogui.click(x, y)
            else:
                pyautogui.press('enter')
            time.sleep(self.cfg.get('search_wait_sec', 2.0))
            if not self.verify_contact(contact_name or str(phone_number)):
                self.log(f"联系人校验失败 -> 期望: {contact_name or phone_number}")
                return False
            pyperclip.copy(message)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.2)
            pyautogui.press('enter')
            time.sleep(self.cfg.get('post_send_wait_sec', 2.0))
            if not self.verify_message(message):
                return False
            return True
        except Exception as e:
            self.log(f"发送异常: {e}")
            return False

    def send_with_retry(self, phone_number: str, message: str, contact_name: str = None) -> bool:
        retries = int(self.cfg.get('max_retries', 1))
        for i in range(1, retries + 1):
            ok = self.send_one(phone_number, message, contact_name)
            if ok:
                self.log(f"✅ 发送成功 -> {contact_name or phone_number}")
                return True
            else:
                self.log(f"验证失败/异常，第 {i} 次尝试")
                time.sleep(0.8)
        self.log(f"发送失败 -> {contact_name or phone_number}")
        return False

# ================== GUI 主程序 ===========================

class SenderApp:
    def __init__(self, root):

        self.root = root
        self.root.title("欠费通知自动发送工具")
        self.root.geometry("820x700")
        self.cfg = DEFAULT_CONFIG.copy()

        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.failed_file_path = os.path.join(self.base_dir, "未发送消息.xlsx")

        self.ocr_manager = OCRManager(tesseract_path=self.cfg.get('tesseract_path'))
        self.sender = Sender(self.cfg, self.log, self.ocr_manager)

        self.build_ui()
        self.update_button_states(os.path.exists(self.failed_file_path))

    # ---------- UI ----------
    def build_ui(self):
        pad = 8
        frm_file = ttk.LabelFrame(self.root, text="Excel 文件")
        frm_file.pack(fill=tk.X, padx=pad, pady=pad)
        self.var_excel = tk.StringVar(value="")
        ttk.Entry(frm_file, textvariable=self.var_excel).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=pad, pady=pad)
        ttk.Button(frm_file, text="选择...", command=self.select_excel).pack(side=tk.LEFT, padx=pad)

        # 鼠标点击设置
        frm_click = ttk.LabelFrame(self.root, text="模拟鼠标点击设置")
        frm_click.pack(fill=tk.X, padx=pad, pady=pad)
        self.use_click = tk.BooleanVar(value=bool(self.cfg.get('use_click', True)))
        ttk.Button(frm_click, text="选择确认区域", command=self.choose_click_point).grid(row=0, column=1,
                                                                                          padx=pad, pady=pad)
        ttk.Checkbutton(frm_click, text="启用 鼠标点击", variable=self.use_click).grid(row=0, column=0,
                                                                                       sticky='w',
                                                                                       padx=pad,
                                                                                       pady=pad)

        frm_ocr = ttk.LabelFrame(self.root, text="OCR 设置")
        frm_ocr.pack(fill=tk.X, padx=pad, pady=pad)
        self.use_ocr = tk.BooleanVar(value=bool(self.cfg.get('use_ocr', True)))
        ttk.Checkbutton(frm_ocr, text="启用 OCR 验证（联系人 + 消息）", variable=self.use_ocr).grid(row=0,
                                                                                                  column=0,
                                                                                                  sticky='w',
                                                                                                  padx=pad,
                                                                                                  pady=pad)


        ttk.Label(frm_ocr, text="相似度阈值").grid(row=0, column=2, sticky='e', padx=pad)
        self.var_threshold = tk.DoubleVar(value=float(self.cfg.get('ocr_threshold', 0.70)))
        ttk.Entry(frm_ocr, textvariable=self.var_threshold, width=8).grid(row=0, column=3, sticky='w', padx=pad)

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

        frm_retry = ttk.LabelFrame(self.root, text="发送策略")
        frm_retry.pack(fill=tk.X, padx=pad, pady=pad)
        ttk.Label(frm_retry, text="最大重试次数").grid(row=0, column=0, sticky='w', padx=pad)
        self.var_retries = tk.IntVar(value=int(self.cfg.get('max_retries', 1)))
        ttk.Entry(frm_retry, textvariable=self.var_retries, width=6).grid(row=0, column=1, sticky='w', padx=pad)
        ttk.Label(frm_retry, text="搜索等待(s)").grid(row=0, column=2, sticky='e', padx=pad)
        self.var_search_wait = tk.DoubleVar(value=float(self.cfg.get('search_wait_sec', 2.0)))
        ttk.Entry(frm_retry, textvariable=self.var_search_wait, width=6).grid(row=0, column=3, sticky='w', padx=pad)
        ttk.Label(frm_retry, text="发送后等待(s)").grid(row=0, column=4, sticky='e', padx=pad)
        self.var_post_wait = tk.DoubleVar(value=float(self.cfg.get('post_send_wait_sec', 2.0)))
        ttk.Entry(frm_retry, textvariable=self.var_post_wait, width=6).grid(row=0, column=5, sticky='w', padx=pad)

        frm_btn = ttk.Frame(self.root)
        frm_btn.pack(fill=tk.X, padx=pad, pady=pad)
        ttk.Button(frm_btn, text="开始处理", command=self.on_start_processing_click).pack(side=tk.LEFT, padx=pad)
        self.btn_open_failed = ttk.Button(frm_btn, text="未发送名单", command=self.open_failed_file, state='disabled')
        self.btn_open_failed.pack(side=tk.LEFT)
        ttk.Button(frm_btn, text="使用说明", command=self.show_instructions).pack(side=tk.LEFT, padx=pad)

        frm_log = ttk.LabelFrame(self.root, text="日志")
        frm_log.pack(fill=tk.BOTH, expand=True, padx=pad, pady=pad)
        self.txt_log = tk.Text(frm_log, height=18)
        self.txt_log.pack(fill=tk.BOTH, expand=True, padx=pad, pady=pad)

    def update_button_states(self, has_failed_file: bool):
        if has_failed_file:
            self.btn_open_failed.config(state='normal')
        else:
            self.btn_open_failed.config(state='disabled')

    def open_failed_file(self):
        if os.path.exists(self.failed_file_path):
            try:
                os.startfile(self.failed_file_path)
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件: {e}")
                self.log(f"无法打开文件: {e}")
        else:
            messagebox.showwarning("提示", "没有可打开的失败文件。")

    def show_instructions(self):
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
        try:
            self.txt_log.insert(tk.END, str(msg) + "\n")
            self.txt_log.see(tk.END)
            self.root.update_idletasks()
        except Exception:
            print(msg)

    def select_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx;*.xls")])
        if path:
            self.var_excel.set(path)
            # 用户选择新文件时，默认禁用“未发送名单”按钮
            self.update_button_states(False)

    def choose_click_point(self):
        messagebox.showinfo("提示", "请选取【联系人名称】所在区域")
        click_point = self.ocr_manager.select_point_gui()
        if click_point:
            self.cfg['click_point'] = click_point
            self.log(f"点击区域: {click_point}")
        else:
            self.log("点击区域选择取消或无效")

    def choose_contact_region(self):
        messagebox.showinfo("提示", "请选取【联系人名称】所在区域")
        region = self.ocr_manager.select_region_gui()
        if region:
            self.cfg['region_contact'] = region
            self.log(f"联系人区域: {region}")
        else:
            self.log("联系人区域选择取消或无效")

    def choose_message_region(self):
        messagebox.showinfo("提示", "请选取【消息内容】所在区域")
        region = self.ocr_manager.select_region_gui()
        if region:
            self.cfg['region_message'] = region
            self.log(f"消息区域: {region}")
        else:
            self.log("消息区域选择取消或无效")

    def preview_contact_ocr(self):
        if not self.cfg.get('region_contact'):
            messagebox.showwarning("提示", "请先选择联系人区域")
            return
        if not self.ocr_manager.tesseract_available:
            messagebox.showwarning("提示", "未安装 pytesseract")
            return
        text = self.ocr_manager.recognize_text(self.cfg['region_contact'], self.cfg['ocr_lang'])
        self.log(f"[预览-联系人] -> {text}")

    def preview_message_ocr(self):
        if not self.cfg.get('region_message'):
            messagebox.showwarning("提示", "请先选择消息区域")
            return
        if not self.ocr_manager.tesseract_available:
            messagebox.showwarning("提示", "未安装 pytesseract")
            return
        text = self.ocr_manager.recognize_text(self.cfg['region_message'], self.cfg['ocr_lang'])
        self.log(f"[预览-消息] -> {text}")

    def on_start_processing_click(self):
        excel_path = self.var_excel.get()
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("错误", "请先选择有效的 Excel 文件")
            return
        self.start_processing(excel_path)

    # 业务主流程，现在可以传入任何 Excel 路径
    def start_processing(self, excel_path: str):
        self.cfg['ocr_threshold'] = float(self.var_threshold.get())
        self.cfg['max_retries'] = int(self.var_retries.get())
        self.cfg['search_wait_sec'] = float(self.var_search_wait.get())
        self.cfg['post_send_wait_sec'] = float(self.var_post_wait.get())
        self.cfg['use_ocr'] = bool(self.use_ocr.get())

        if self.cfg['use_click'] and not self.cfg.get('click_point'):
            # 弹出警告框，让用户决定是否继续
            messagebox.showerror(
                "警告",
                "已勾选“启用 鼠标点击”，但未设置点击坐标。\n请设置点击坐标后再使用。"
            )
            return

        try:
            sheets = pd.read_excel(excel_path, sheet_name=None)
        except Exception as e:
            messagebox.showerror("错误", f"无法读取 Excel: {e}")
            return

        try:
            pyautogui.hotkey('win')
            time.sleep(0.8)
            pyperclip.copy("移动办公")
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.6)
            pyautogui.press('enter')
            time.sleep(1.0)
            pyautogui.press('enter')
            time.sleep(1.5)
        except Exception as e:
            self.log(f"尝试启动应用失败（可忽略，若已打开）: {e}")

        tgt = set(self.cfg.get('target_sheets', []))
        total, okcnt, failcnt = 0, 0, 0
        failed_sends_by_sheet = {}

        for sheet_name, df in sheets.items():
            if sheet_name not in tgt:
                continue
            self.log(f"==== 处理 Sheet: {sheet_name} ====")
            if not isinstance(df, pd.DataFrame) or df.empty:
                self.log(f"Sheet {sheet_name} 为空，跳过")
                continue
            for idx, row in df.iterrows():
                try:
                    is_failed = False
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
                        if sheet_name not in failed_sends_by_sheet:
                            failed_sends_by_sheet[sheet_name] = []
                        failed_sends_by_sheet[sheet_name].append(row)
                except Exception as e:
                    self.log(f"行 {idx + 1} 处理异常: {e}")
                    if sheet_name not in failed_sends_by_sheet:
                        failed_sends_by_sheet[sheet_name] = []
                    failed_sends_by_sheet[sheet_name].append(row)

        self.log(f"完成。总计: {total} | 成功: {okcnt} | 失败: {failcnt}")
        messagebox.showinfo("发送结果", f"发送完成！\n成功: {okcnt} 条\n失败: {failcnt} 条")

        if failed_sends_by_sheet:
            try:
                with pd.ExcelWriter(self.failed_file_path, engine='openpyxl') as writer:
                    for sheet, failed_rows in failed_sends_by_sheet.items():
                        df_failed = pd.DataFrame(failed_rows)
                        df_failed.to_excel(writer, sheet_name=sheet, index=False)
                self.update_button_states(True)
                self.log(f"⚠️ {failcnt} 条发送失败，已自动保存至 {self.failed_file_path}。您可以通过点击“未发送名单”按钮来查看详情，"
                         f"并可将此文件作为新的数据源进行二次发送。")
            except Exception as e:
                self.log(f"保存失败文件时出错: {e}")
                messagebox.showerror("错误", f"保存失败文件时出错: {e}")
        else:
            if os.path.exists(self.failed_file_path):
                os.remove(self.failed_file_path)
            self.update_button_states(False)
            self.log("🎉 所有信息发送成功，没有失败记录")

if __name__ == '__main__':
    root = tk.Tk()
    SenderApp(root)
    root.mainloop()