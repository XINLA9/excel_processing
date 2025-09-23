import os
import re
import sys
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox, ttk

import pandas as pd


class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, string):
        self.widget.configure(state="normal")
        self.widget.insert("end", string, (self.tag,))
        self.widget.see("end")
        self.widget.configure(state="disabled")

    def flush(self):
        pass


def main():
    """主函数：创建文件选择界面"""
    # 创建主窗口
    root = tk.Tk()
    root.title("数据催款处理工具")
    root.geometry("800x700")  # 增加窗口高度以容纳新控件
    root.configure(bg="#f0f0f0")

    # 存储文件路径的变量
    file_paths = {
        "原始数据文件": "",
        "维度表文件": "",
        "剔除工单号文件": "",
        "保存文件夹": ""
    }

    # 存储阈值
    threshold_var = tk.IntVar(value=20)

    # 创建样式
    style = ttk.Style()
    style.configure("Green.TLabel", foreground="green")
    style.configure("Red.TLabel", foreground="red")

    # 创建标签字典
    labels = {}

    def select_file(file_type):
        """选择文件"""
        if file_type == "保存文件夹":
            path = filedialog.askdirectory(title=f"选择{file_type}")
        else:
            filetypes = [("Excel文件", "*.xlsx;*.xls"), ("所有文件", "*.*")]
            path = filedialog.askopenfilename(title=f"选择{file_type}", filetypes=filetypes)

        if path:
            file_paths[file_type] = path
            update_labels()
        return path

    def update_labels():
        """更新界面上的文件路径显示"""
        for file_type, path in file_paths.items():
            label = labels[file_type]
            if path:
                # 显示简短路径（只显示最后两级目录）
                parts = path.split(os.sep)
                if len(parts) > 2:
                    short_path = os.sep.join(parts[-2:])
                else:
                    short_path = path
                label.config(text=f"{file_type}: {short_path}")
                label.configure(style="Green.TLabel")
            else:
                label.config(text=f"{file_type}: 未选择")
                label.configure(style="Red.TLabel")

    def start_processing():
        """开始处理数据"""
        # 检查所有文件是否已选择
        missing_files = [file_type for file_type, path in file_paths.items()
                         if not path and file_type != "剔除工单号文件"]  # 剔除工单号文件是可选的

        if missing_files:
            messagebox.showerror("错误", f"请先选择以下文件：\n{', '.join(missing_files)}")
            return

        # 获取阈值
        threshold = threshold_var.get()

        # 禁用开始按钮，避免重复点击
        process_button.config(state="disabled")
        status_label.config(text="处理中，请稍候...")

        # 清空输出框
        output_text.configure(state="normal")
        output_text.delete(1.0, tk.END)
        output_text.configure(state="disabled")

        # 在新线程中处理数据，避免界面卡死
        import threading
        thread = threading.Thread(
            target=lambda: process_data(file_paths, threshold, root, process_button, status_label, output_text))
        thread.daemon = True
        thread.start()

    # 创建界面
    main_frame = ttk.Frame(root, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)

    # 标题
    title_label = ttk.Label(main_frame, text="数据催款处理工具", font=("Arial", 16, "bold"))
    title_label.pack(pady=10)

    # 说明文字
    info_label = ttk.Label(main_frame, text="请选择以下文件，然后点击\"开始处理\"按钮", font=("Arial", 10))
    info_label.pack(pady=5)

    # 创建文件选择区域
    file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
    file_frame.pack(fill=tk.X, pady=10)

    for i, file_type in enumerate(["原始数据文件", "维度表文件", "剔除工单号文件", "保存文件夹"]):
        row_frame = ttk.Frame(file_frame)
        row_frame.pack(fill=tk.X, pady=5)

        # 创建标签并存储到字典中
        label = ttk.Label(row_frame, text=f"{file_type}: 未选择", width=60, anchor="w")
        label.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        labels[file_type] = label

        button = ttk.Button(row_frame, text="选择",
                            command=lambda ft=file_type: select_file(ft))
        button.pack(side=tk.RIGHT)

    # 阈值设置区域
    threshold_frame = ttk.LabelFrame(main_frame, text="客户经理未回款发票阈值设置", padding="10")
    threshold_frame.pack(fill=tk.X, pady=10)

    threshold_label = ttk.Label(threshold_frame, text="阈值:")
    threshold_label.pack(side=tk.LEFT, padx=(0, 10))

    threshold_spinbox = ttk.Spinbox(threshold_frame, from_=1, to=1000, textvariable=threshold_var, width=10)
    threshold_spinbox.pack(side=tk.LEFT, padx=(0, 20))

    threshold_info = ttk.Label(threshold_frame,
                               text="设置客户经理未回款发票数量的阈值，超过此值的客户经理明细数据将被单独保存")
    threshold_info.pack(side=tk.LEFT, fill=tk.X, expand=True)

    # 处理按钮
    process_button = ttk.Button(main_frame, text="开始处理", command=start_processing, state="normal")
    process_button.pack(pady=10)

    # 状态标签
    status_label = ttk.Label(main_frame, text="准备就绪", font=("Arial", 9))
    status_label.pack(pady=5)

    # 输出框标签
    output_label = ttk.Label(main_frame, text="处理日志:", font=("Arial", 10))
    output_label.pack(anchor="w", pady=(10, 5))

    # 输出文本框
    output_frame = ttk.Frame(main_frame)
    output_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

    output_text = tk.Text(output_frame, height=10, state="disabled", wrap=tk.WORD)
    output_scrollbar = ttk.Scrollbar(output_frame, orient="vertical", command=output_text.yview)
    output_text.configure(yscrollcommand=output_scrollbar.set)

    output_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    output_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # 初始化标签显示
    update_labels()

    # 启动界面
    root.mainloop()


# 修改 process_data 函数以接受 threshold 参数
def process_data(file_paths, threshold, root, process_button, status_label, output_text):
    """处理数据的主函数"""
    try:
        # 重定向标准输出到文本框
        old_stdout = sys.stdout
        sys.stdout = TextRedirector(output_text)

        # 从参数获取文件路径
        file_path = file_paths["原始数据文件"]
        file_path2 = file_paths["维度表文件"]
        exclude_file_path = file_paths["剔除工单号文件"]
        save_dir = file_paths["保存文件夹"]

        print(f"开始处理数据...")
        print(f"原始数据文件: {file_path}")
        print(f"维度表文件: {file_path2}")
        print(f"剔除工单号文件: {exclude_file_path}")
        print(f"保存文件夹: {save_dir}")
        print(f"客户经理未回款发票阈值: {threshold}")

        # 读取原始数据
        df = pd.read_excel(file_path)
        df_copy = df.copy()

        print(df_copy.head(5))
        print(df_copy.shape)

        # 处理表头
        df_copy = df_copy.drop(0)
        df_copy.columns = df_copy.iloc[0]
        df_copy = df_copy.iloc[1:].reset_index(drop=True)

        # S列字段数据【发票状态】=已开具
        df_copy = df_copy[df_copy['发票状态'] == '已开具']

        # 【发票总金额】格式类型，要从文本变为数字 并筛选【发票总金额】>0
        df_copy['发票总金额'] = df_copy['发票总金额'].astype(str).apply(lambda x: re.sub(r'[^\d.-]', '', x))
        df_copy['发票总金额'] = pd.to_numeric(df_copy['发票总金额'], errors='coerce')
        # 检查是否有转换失败的NaN值
        na_count = df_copy['发票总金额'].isna().sum()
        if na_count > 0:
            print(f"注意：有{na_count}条数据转换失败，已设为NaN")
        # 检查负数保留情况
        negative_count = (df_copy['发票总金额'] < 0).sum()
        print(f"转换后保留的负数数量: {negative_count}")
        df_copy = df_copy[df_copy['发票总金额'] > 0]

        # L列字段【是否完全销账】选择"否"
        df_copy = df_copy[df_copy['是否已完全销账'] == '否']

        # 按P列【发票号码】去重，相同的发票号仅保留一个
        df_copy = df_copy.drop_duplicates(subset=['发票号码'], keep='first')

        # 读取需要剔除的工单号文件
        if exclude_file_path:
            try:
                df_exclude = pd.read_excel(exclude_file_path)
                print(f"读取到需要剔除的发票号码数量: {len(df_exclude)}")

                # 获取需要剔除的发票号码列表
                exclude_invoice_numbers = df_exclude['发票号码'].tolist()
                print(f"需要剔除的发票号码示例: {exclude_invoice_numbers[:5]}")

                # 剔除这些发票号码
                original_count = len(df_copy)
                df_copy = df_copy[~df_copy['发票号码'].isin(exclude_invoice_numbers)]
                removed_count = original_count - len(df_copy)
                print(f"剔除了 {removed_count} 条数据")

            except Exception as e:
                print(f"读取剔除工单号文件时出错: {e}")

        # 读取四个sheet页
        try:
            # 读取提单人维表
            df_bill_person = pd.read_excel(file_path2, sheet_name=0)  # 第一个sheet
            # 读取客户经理维表
            df_account_manager = pd.read_excel(file_path2, sheet_name=1)  # 第二个sheet
            # 读取集团名称维表
            df_customer_group = pd.read_excel(file_path2, sheet_name=2)  # 第三个sheet
            # 读取客户经理通讯录
            df_contact_list = pd.read_excel(file_path2, sheet_name=3)  # 第四个sheet

            print("提单人维表:")
            print(df_bill_person.head())
            print("客户经理维表:")
            print(df_account_manager.head())
            print("集团名称维表:")
            print(df_customer_group.head())
            print("客户经理通讯录:")
            print(df_contact_list.head())

        except Exception as e:
            print(f"读取维度表时出错: {e}")
            # 在界面线程中显示错误信息
            root.after(0, lambda: messagebox.showerror("错误", f"读取维度表时出错: {e}"))
            return

        # 创建映射字典
        # 提单人名称映射
        bill_person_name_mapping = df_bill_person.drop_duplicates(subset=['提单人名称']).set_index('提单人名称')[
            '分公司']
        # 提单人工号映射
        bill_person_id_mapping = df_bill_person.drop_duplicates(subset=['提单人工号']).set_index('提单人工号')['分公司']
        # 客户经理映射
        account_manager_mapping = df_account_manager.drop_duplicates(subset=['客户经理']).set_index('客户经理')[
            '分公司']
        # 客户名称映射
        customer_group_mapping = df_customer_group.drop_duplicates(subset=['客户名称']).set_index('客户名称')['分公司']

        # 创建客户经理通讯录映射字典
        director_mapping = df_contact_list.drop_duplicates(subset=['姓名']).set_index('姓名')['总监']
        director_phone_mapping = df_contact_list.drop_duplicates(subset=['姓名']).set_index('姓名')['总监电话']
        leader_mapping = df_contact_list.drop_duplicates(subset=['姓名']).set_index('姓名')['分管领导']
        leader_phone_mapping = df_contact_list.drop_duplicates(subset=['姓名']).set_index('姓名')['分管领导电话']
        # 新增：客户经理电话映射
        manager_phone_mapping = df_contact_list.drop_duplicates(subset=['姓名']).set_index('姓名')['联系电话']

        # 初始化所属分公司列
        df_copy['所属分公司'] = '未知'

        # 第一步：按提单人名称匹配
        print("开始按提单人名称匹配分公司...")
        name_mask = (~df_copy['提单人名称'].isna()) & (df_copy['所属分公司'] == '未知')
        df_copy.loc[name_mask, '所属分公司'] = df_copy.loc[name_mask, '提单人名称'].map(
            bill_person_name_mapping).fillna('未知')

        # 第二步：按提单人工号匹配（针对名称匹配失败的数据）
        print("开始按提单人工号匹配分公司...")
        id_mask = (df_copy['所属分公司'] == '未知') & (~df_copy['提单人工号'].isna())
        df_copy.loc[id_mask, '所属分公司'] = df_copy.loc[id_mask, '提单人工号'].map(bill_person_id_mapping).fillna(
            '未知')

        # 第三步：按客户经理名称匹配
        print("开始按客户经理名称匹配分公司...")
        manager_mask = (df_copy['所属分公司'] == '未知') & (~df_copy['客户经理名称'].isna())
        df_copy.loc[manager_mask, '所属分公司'] = df_copy.loc[manager_mask, '客户经理名称'].map(
            account_manager_mapping).fillna('未知')

        # 第四步：按客户名称匹配
        print("开始按客户名称匹配分公司...")
        customer_mask = (df_copy['所属分公司'] == '未知') & (~df_copy['客户名称'].isna())
        df_copy.loc[customer_mask, '所属分公司'] = df_copy.loc[customer_mask, '客户名称'].map(
            customer_group_mapping).fillna('未知')

        # 统计分公司匹配结果
        match_stats = df_copy['所属分公司'].value_counts()
        print("分公司匹配结果统计:")
        print(match_stats)

        # 获取未知分公司的数据
        unknown_branch = df_copy[df_copy['所属分公司'] == '未知']
        print(f"未知分公司的数据条数: {len(unknown_branch)}")

        # ==================== 新增：补充客户经理匹配 ====================
        print("\n开始补充客户经理匹配...")

        # 创建客户经理维表的映射字典
        # 根据工号映射客户经理
        manager_id_mapping = df_account_manager.drop_duplicates(subset=['对应工号']).set_index('对应工号')['客户经理']
        # 根据集团名称映射客户经理（取第一个匹配的）
        group_name_mapping = df_account_manager.drop_duplicates(subset=['集团名称']).set_index('集团名称')['客户经理']

        # 初始化补充客户经理列
        df_copy['补充客户经理'] = '未知'

        # 第一步：复制原有的客户经理名称（如果不为空）
        print("第一步：复制原有客户经理名称...")
        existing_manager_mask = ~df_copy['客户经理名称'].isna()
        df_copy.loc[existing_manager_mask, '补充客户经理'] = df_copy.loc[existing_manager_mask, '客户经理名称']

        # 第三步：通过客户名称匹配集团名称（针对前两步仍未匹配的数据）
        print("第三步：通过客户名称匹配客户经理...")
        group_match_mask = (df_copy['补充客户经理'] == '未知') & (~df_copy['客户名称'].isna())
        df_copy.loc[group_match_mask, '补充客户经理'] = df_copy.loc[group_match_mask, '客户名称'].map(
            group_name_mapping).fillna('未知')

        # 统计补充客户经理匹配结果
        manager_match_stats = df_copy['补充客户经理'].value_counts()
        print("补充客户经理匹配结果统计:")
        print(manager_match_stats)

        # 获取未知客户经理的数据
        unknown_manager = df_copy[df_copy['补充客户经理'] == '未知']
        print(f"未知客户经理的数据条数: {len(unknown_manager)}")

        # ==================== 日期处理和回款计算 ====================
        # 设置日期，计算回款金额
        # 首先将开票日期转换为datetime格式
        df_copy['开票日期'] = pd.to_datetime(df_copy['开票日期'], format='%Y%m%d', errors='coerce')

        # 过滤掉日期转换失败的数据
        invalid_date_count = df_copy['开票日期'].isna().sum()
        if invalid_date_count > 0:
            print(f"注意：有{invalid_date_count}条数据的开票日期格式无效，已过滤")
            df_copy = df_copy.dropna(subset=['开票日期'])

        # 定义基准日期（当前日期）
        base_date = datetime.today()

        # 计算回款天数
        df_copy['回款天数'] = (base_date - df_copy['开票日期']).dt.days

        # 新增【催款类型】列：按回款天数范围分类
        def get_collection_type(days):
            if days <= 30:
                return "小于30天"
            elif 30 < days <= 60:
                return "大于30天并小于等于60天"
            elif 60 < days <= 90:
                return "大于60天并小于等于90天"
            else:
                return "大于90天"

        # 应用分类逻辑
        df_copy['催款类型'] = df_copy['回款天数'].apply(get_collection_type)

        # ==================== 新增：客户经理统计和明细提取 ====================
        print("\n开始统计客户经理未回款发票数量...")

        # 筛选出超过30天的数据
        df_over_30_days = df_copy[df_copy['回款天数'] > 30]
        print(f"超过30天的未回款发票数量: {len(df_over_30_days)}")

        # 统计每个客户经理的超过30天的未回款发票数量
        manager_stats = df_over_30_days['补充客户经理'].value_counts().reset_index()
        manager_stats.columns = ['客户经理', '未回款发票数量']

        # 筛选出超过阈值的客户经理
        exceeded_managers = manager_stats[manager_stats['未回款发票数量'] > threshold]
        print(f"超过阈值({threshold})的客户经理数量: {len(exceeded_managers)}")

        # 创建未达标客户经理文件夹
        unqualified_dir = os.path.join(save_dir, "未达标客户经理")
        os.makedirs(unqualified_dir, exist_ok=True)

        # 为每个超过阈值的客户经理创建明细文件
        if not exceeded_managers.empty:
            print("开始提取超过阈值的客户经理明细数据...")

            for _, row in exceeded_managers.iterrows():
                manager_name = row['客户经理']
                invoice_count = row['未回款发票数量']

                # 筛选该客户经理的超过30天的明细数据
                manager_data = df_over_30_days[df_over_30_days['补充客户经理'] == manager_name]

                # 选择需要的字段
                manager_detail = manager_data[[
                    '所属分公司', '补充客户经理', '客户名称', '发票类型', '预开类型',
                    '发票号码', '集团产品号码/CMIOT账户编号', '发票总金额', '发票状态',
                    '账期', '是否已完全销账', '开票日期', '催款类型', '回款天数'
                ]]

                # 保存到Excel文件
                manager_filename = f"{manager_name}_未回款明细_{datetime.now().strftime('%Y%m%d')}.xlsx"
                manager_filepath = os.path.join(unqualified_dir, manager_filename)
                manager_detail.to_excel(manager_filepath, index=False)

                print(f"已保存 {manager_name} 的明细数据，共 {invoice_count} 条记录")

        # ==================== 创建不同催款类型的数据表 ====================
        # 一、30天通报
        df_30_days = df_copy[df_copy['催款类型'] == '大于30天并小于等于60天'].copy()
        if not df_30_days.empty:
            df_30_days['收票日期'] = df_30_days['开票日期'] + timedelta(days=30)
            df_30_days['总监'] = ''
            df_30_days['分管领导'] = ''
            df_30_days['总监电话'] = ''
            df_30_days['分管领导电话'] = ''
            # 新增客户经理电话列
            df_30_days['客户经理电话'] = df_30_days['补充客户经理'].map(manager_phone_mapping).fillna('')
            df_30_days['短信模板'] = df_30_days.apply(
                lambda
                    row: f"客户经理{row['补充客户经理']}名下{row['客户名称']}于{row['开票日期'].strftime('%Y-%m-%d')}开具发票，票号{row['发票号码']}，金额{row['发票总金额']}，逾期未回款30天以上，请尽快回款。如客户违约拒不回款的，应及时与客户确认后冲红发票。",
                axis=1
            )
            df_30_days = df_30_days[
                ['补充客户经理', '客户经理电话', '催款类型', '客户名称', '开票日期', '发票号码', '发票总金额',
                 '收票日期', '短信模板', '总监', '分管领导', '总监电话', '分管领导电话']]

        # 二、60天通报
        df_60_days = df_copy[df_copy['催款类型'] == '大于60天并小于等于90天'].copy()
        if not df_60_days.empty:
            df_60_days['收票日期'] = df_60_days['开票日期'] + timedelta(days=60)
            df_60_days['总监'] = df_60_days['补充客户经理'].map(director_mapping).fillna('')
            df_60_days['总监电话'] = df_60_days['补充客户经理'].map(director_phone_mapping).fillna('')
            df_60_days['分管领导'] = ''
            df_60_days['分管领导电话'] = ''
            # 新增客户经理电话列
            df_60_days['客户经理电话'] = df_60_days['补充客户经理'].map(manager_phone_mapping).fillna('')
            df_60_days['短信模板'] = df_60_days.apply(
                lambda
                    row: f"客户经理{row['补充客户经理']}名下{row['客户名称']}于{row['开票日期'].strftime('%Y-%m-%d')}开具发票，票号{row['发票号码']}，金额{row['发票总金额']}，逾期未回款60天以上，请尽快回款。如客户违约拒不回款的，应及时与客户确认后冲红发票。",
                axis=1
            )
            df_60_days = df_60_days[
                ['补充客户经理', '客户经理电话', '催款类型', '客户名称', '开票日期', '发票号码', '发票总金额',
                 '收票日期', '短信模板', '总监', '分管领导', '总监电话', '分管领导电话']]

        # 三、90天通报
        df_90_days = df_copy[df_copy['催款类型'] == '大于90天'].copy()
        if not df_90_days.empty:
            # 计算一周后的日期
            one_week_later = datetime.today() + timedelta(days=7)
            df_90_days['收票日期'] = one_week_later
            df_90_days['总监'] = df_90_days['补充客户经理'].map(director_mapping).fillna('')
            df_90_days['总监电话'] = df_90_days['补充客户经理'].map(director_phone_mapping).fillna('')
            df_90_days['分管领导'] = df_90_days['补充客户经理'].map(leader_mapping).fillna('')
            df_90_days['分管领导电话'] = df_90_days['补充客户经理'].map(leader_phone_mapping).fillna('')
            # 新增客户经理电话列
            df_90_days['客户经理电话'] = df_90_days['补充客户经理'].map(manager_phone_mapping).fillna('')
            df_90_days['短信模板'] = df_90_days.apply(
                lambda
                    row: f"客户经理{row['补充客户经理']}名下{row['客户名称']}于{row['开票日期'].strftime('%Y-%m-%d')}开具发票，票号{row['发票号码']}，金额{row['发票总金额']}，逾期未回款90天以上，请于{row['收票日期'].strftime('%Y-%m-%d')}内回款。如客户违约拒不回款的，应及时与客户确认后冲红发票。",
                axis=1
            )
            df_90_days = df_90_days[
                ['补充客户经理', '客户经理电话', '催款类型', '客户名称', '开票日期', '发票号码', '发票总金额',
                 '收票日期', '短信模板', '总监', '分管领导', '总监电话', '分管领导电话']]

        # 创建数据透视表
        payment_types = ["小于30天", "大于30天并小于等于60天", "大于60天并小于等于90天", "大于90天"]
        pivot_table = pd.pivot_table(
            df_copy,
            index='所属分公司',
            columns='催款类型',
            values='发票总金额',
            aggfunc='sum'
        ).reindex(columns=payment_types).fillna(0)

        print(pivot_table)

        # 获取当前时间作为文件名的一部分
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        result_filename = f"催款处理结果_{current_time}.xlsx"
        result_filepath = os.path.join(save_dir, result_filename)

        # 将所有结果保存到一个Excel文件的不同sheet页中
        with pd.ExcelWriter(result_filepath, engine='openpyxl') as writer:
            # 保存完整处理后的数据
            df_copy.to_excel(writer, sheet_name='处理后的数据', index=False)

            # 保存数据透视表
            pivot_table.to_excel(writer, sheet_name='数据汇总', index=True)

            # 保存客户经理统计
            manager_stats.to_excel(writer, sheet_name='客户经理统计', index=False)

            # 保存未匹配的数据
            if not unknown_branch.empty:
                unknown_branch.to_excel(writer, sheet_name='未匹配分公司数据', index=False)

            if not unknown_manager.empty:
                unknown_manager.to_excel(writer, sheet_name='未匹配客户经理数据', index=False)

            # 保存不同催款类型的数据
            if not df_30_days.empty:
                df_30_days.to_excel(writer, sheet_name='30天通报', index=False)

            if not df_60_days.empty:
                df_60_days.to_excel(writer, sheet_name='60天通报', index=False)

            if not df_90_days.empty:
                df_90_days.to_excel(writer, sheet_name='90天通报', index=False)

        print("处理完成！")

        # 在界面线程中显示完成消息
        success_message = f"文件已成功保存至:\n{result_filepath}"
        if not exceeded_managers.empty:
            success_message += f"\n\n超过阈值的客户经理明细已保存至:\n{unqualified_dir}"

        root.after(0, lambda: messagebox.showinfo("完成", success_message))
        root.after(0, lambda: status_label.config(text="处理完成"))

    except Exception as e:
        print(f"处理过程中出错: {e}")
        # 在界面线程中显示错误信息
        root.after(0, lambda: messagebox.showerror("错误", f"处理过程中出错: {e}"))
        root.after(0, lambda: status_label.config(text=f"处理失败: {str(e)}"))

    finally:
        # 恢复标准输出
        sys.stdout = old_stdout
        # 重新启用开始按钮
        root.after(0, lambda: process_button.config(state="normal"))


# 启动程序
if __name__ == "__main__":
    main()