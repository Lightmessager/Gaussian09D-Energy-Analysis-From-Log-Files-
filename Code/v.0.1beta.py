import pandas as pd
import re
import os
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sys
import platform


class GaussianEnergyAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("Gaussian 能量分析工具")
        self.root.geometry("1200x800")
        self.root.resizable(True, True)

        # 设置应用图标（如果有）
        self.set_icon()

        # 初始化数据
        self.scf_files = []
        self.gibbs_files = []
        self.data = pd.DataFrame(columns=["文件名", "SCF能量(a.u.)", "Gibbs校正(a.u.)", "总能量(a.u.)"])

        # 创建主框架
        self.main_frame = ttk.Frame(root, padding=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # 文件选择部分 - 左右两列布局
        files_frame = ttk.Frame(self.main_frame)
        files_frame.pack(fill=tk.X, pady=10)

        # SCF文件选择
        scf_frame = ttk.LabelFrame(files_frame, text="SCF Done 文件", padding=10)
        scf_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        self.scf_listbox = tk.Listbox(scf_frame, height=10, selectmode=tk.EXTENDED)
        self.scf_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        scf_btn_frame = ttk.Frame(scf_frame)
        scf_btn_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(scf_btn_frame, text="添加SCF文件", command=self.add_scf_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(scf_btn_frame, text="上移选中", command=lambda: self.move_item(self.scf_listbox, -1)).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(scf_btn_frame, text="下移选中", command=lambda: self.move_item(self.scf_listbox, 1)).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(scf_btn_frame, text="移除选中",
                   command=lambda: self.remove_selected(self.scf_listbox, self.scf_files)).pack(side=tk.LEFT, padx=2)

        # Gibbs文件选择
        gibbs_frame = ttk.LabelFrame(files_frame, text="Gibbs校正文件", padding=10)
        gibbs_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        self.gibbs_listbox = tk.Listbox(gibbs_frame, height=10, selectmode=tk.EXTENDED)
        self.gibbs_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        gibbs_btn_frame = ttk.Frame(gibbs_frame)
        gibbs_btn_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(gibbs_btn_frame, text="添加Gibbs文件", command=self.add_gibbs_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(gibbs_btn_frame, text="上移选中", command=lambda: self.move_item(self.gibbs_listbox, -1)).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(gibbs_btn_frame, text="下移选中", command=lambda: self.move_item(self.gibbs_listbox, 1)).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(gibbs_btn_frame, text="移除选中",
                   command=lambda: self.remove_selected(self.gibbs_listbox, self.gibbs_files)).pack(side=tk.LEFT,
                                                                                                    padx=2)

        # 处理按钮
        process_frame = ttk.Frame(self.main_frame)
        process_frame.pack(fill=tk.X, pady=10)

        ttk.Button(process_frame, text="自动匹配", command=self.auto_match).pack(side=tk.LEFT, padx=10)
        ttk.Button(process_frame, text="手动匹配", command=self.manual_match).pack(side=tk.LEFT, padx=10)
        ttk.Button(process_frame, text="提取能量数据", command=self.extract_energies).pack(side=tk.LEFT, padx=10)

        # 表格显示
        table_frame = ttk.LabelFrame(self.main_frame, text="能量数据", padding=10)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # 创建表格
        self.tree = ttk.Treeview(table_frame, columns=("SCF能量", "Gibbs校正", "总能量"), show="headings")
        self.tree.heading("#0", text="文件名")
        self.tree.heading("SCF能量", text="SCF能量(a.u.)")
        self.tree.heading("Gibbs校正", text="Gibbs校正(a.u.)")
        self.tree.heading("总能量", text="总能量(a.u.)")

        # 设置列宽
        self.tree.column("#0", width=250)
        self.tree.column("SCF能量", width=150)
        self.tree.column("Gibbs校正", width=150)
        self.tree.column("总能量", width=150)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 导出按钮
        export_frame = ttk.Frame(self.main_frame)
        export_frame.pack(fill=tk.X, pady=10)

        ttk.Button(export_frame, text="导出为Excel", command=self.export_to_excel).pack(side=tk.RIGHT, padx=10)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(self.main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).pack(
            fill=tk.X, pady=5)

    def set_icon(self):
        """尝试设置应用图标（如果可用）"""
        try:
            if sys.platform == "win32":
                self.root.iconbitmap("icon.ico")  # 如果有图标文件
        except:
            pass

    def add_scf_files(self):
        files = filedialog.askopenfilenames(
            title="选择SCF Done文件",
            filetypes=[("Gaussian输出文件", "*.log"), ("所有文件", "*.*")]
        )
        if files:
            self.scf_files.extend(files)
            self.update_listbox(self.scf_listbox, self.scf_files)
            self.status_var.set(f"已添加 {len(files)} 个SCF文件")

    def add_gibbs_files(self):
        files = filedialog.askopenfilenames(
            title="选择Gibbs校正文件",
            filetypes=[("Gaussian输出文件", "*.log"), ("所有文件", "*.*")]
        )
        if files:
            self.gibbs_files.extend(files)
            self.update_listbox(self.gibbs_listbox, self.gibbs_files)
            self.status_var.set(f"已添加 {len(files)} 个Gibbs校正文件")

    def update_listbox(self, listbox, files):
        listbox.delete(0, tk.END)
        for file in files:
            listbox.insert(tk.END, os.path.basename(file))

    def move_item(self, listbox, direction):
        selected = listbox.curselection()
        if not selected:
            return

        for idx in selected:
            if direction < 0 and idx > 0:  # 上移
                listbox.insert(idx - 1, listbox.get(idx))
                listbox.delete(idx + 1)
                listbox.selection_set(idx - 1)
            elif direction > 0 and idx < listbox.size() - 1:  # 下移
                listbox.insert(idx + 2, listbox.get(idx))
                listbox.delete(idx)
                listbox.selection_set(idx + 1)

    def remove_selected(self, listbox, files_list):
        selected_indices = listbox.curselection()
        if selected_indices:
            # 从后往前删除避免索引变化
            for idx in sorted(selected_indices, reverse=True):
                del files_list[idx]
            self.update_listbox(listbox, files_list)

    def auto_match(self):
        """根据文件名自动匹配SCF和Gibbs文件"""
        if not self.scf_files:
            messagebox.showwarning("无文件", "请先添加SCF文件")
            return

        # 创建匹配字典
        scf_basenames = {os.path.basename(f)[3:] for f in self.scf_files if os.path.basename(f).startswith("sp_")}
        gibbs_basenames = [os.path.basename(f) for f in self.gibbs_files]

        # 自动重新排序Gibbs文件列表以匹配SCF顺序
        new_gibbs_files = []
        unmatched = []

        for scf_file in self.scf_files:
            scf_base = os.path.basename(scf_file)
            if scf_base.startswith("sp_"):
                base_name = scf_base[3:]
                found = False

                # 查找匹配的Gibbs文件
                for gibbs_file in self.gibbs_files:
                    if os.path.basename(gibbs_file) == base_name:
                        new_gibbs_files.append(gibbs_file)
                        found = True
                        break

                if not found:
                    new_gibbs_files.append(None)
                    unmatched.append(scf_base)
            else:
                new_gibbs_files.append(None)
                unmatched.append(scf_base)

        # 添加未匹配的Gibbs文件
        for gibbs_file in self.gibbs_files:
            if os.path.basename(gibbs_file) not in [os.path.basename(f) for f in new_gibbs_files if f]:
                new_gibbs_files.append(gibbs_file)
                self.scf_files.append(None)  # 添加对应的SCF空值

        self.gibbs_files = new_gibbs_files
        self.update_listbox(self.scf_listbox, [f if f else "(无匹配文件)" for f in self.scf_files])
        self.update_listbox(self.gibbs_listbox, [f if f else "(无匹配文件)" for f in self.gibbs_files])

        if unmatched:
            self.status_var.set(f"自动匹配完成，{len(unmatched)}个SCF文件未找到对应Gibbs文件")
        else:
            self.status_var.set("所有SCF文件已成功匹配对应Gibbs文件")

    def manual_match(self):
        """手动调整SCF和Gibbs文件顺序"""
        if not self.scf_files and not self.gibbs_files:
            messagebox.showwarning("无文件", "请先添加文件")
            return

        # 创建手动匹配对话框
        match_dialog = tk.Toplevel(self.root)
        match_dialog.title("手动调整文件匹配")
        match_dialog.geometry("800x500")
        match_dialog.transient(self.root)
        match_dialog.grab_set()

        # 创建两列框架
        main_frame = ttk.Frame(match_dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        # SCF文件列表
        ttk.Label(left_frame, text="SCF文件顺序", font=("Arial", 10, "bold")).pack(pady=5)
        scf_list = tk.Listbox(left_frame, selectmode=tk.SINGLE)
        scf_list.pack(fill=tk.BOTH, expand=True, pady=5)
        for file in self.scf_files:
            scf_list.insert(tk.END, os.path.basename(file) if file else "(无匹配文件)")

        # Gibbs文件列表
        ttk.Label(right_frame, text="Gibbs文件顺序", font=("Arial", 10, "bold")).pack(pady=5)
        gibbs_list = tk.Listbox(right_frame, selectmode=tk.SINGLE)
        gibbs_list.pack(fill=tk.BOTH, expand=True, pady=5)
        for file in self.gibbs_files:
            gibbs_list.insert(tk.END, os.path.basename(file) if file else "(无匹配文件)")

        # 控制按钮
        control_frame = ttk.Frame(match_dialog)
        control_frame.pack(fill=tk.X, pady=10)

        def move(direction, listbox):
            selected = listbox.curselection()
            if not selected:
                return
            idx = selected[0]

            if direction == "up" and idx > 0:
                item = listbox.get(idx)
                listbox.delete(idx)
                listbox.insert(idx - 1, item)
                listbox.selection_set(idx - 1)
            elif direction == "down" and idx < listbox.size() - 1:
                item = listbox.get(idx)
                listbox.delete(idx)
                listbox.insert(idx + 1, item)
                listbox.selection_set(idx + 1)

        ttk.Button(control_frame, text="SCF上移", command=lambda: move("up", scf_list)).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="SCF下移", command=lambda: move("down", scf_list)).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Gibbs上移", command=lambda: move("up", gibbs_list)).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Gibbs下移", command=lambda: move("down", gibbs_list)).pack(side=tk.LEFT, padx=5)

        def apply_changes():
            # 更新SCF顺序
            new_scf = []
            for i in range(scf_list.size()):
                for f in self.scf_files:
                    if f and os.path.basename(f) == scf_list.get(i):
                        new_scf.append(f)
                        break
                else:
                    new_scf.append(None)

            # 更新Gibbs顺序
            new_gibbs = []
            for i in range(gibbs_list.size()):
                for f in self.gibbs_files:
                    if f and os.path.basename(f) == gibbs_list.get(i):
                        new_gibbs.append(f)
                        break
                else:
                    new_gibbs.append(None)

            self.scf_files = new_scf
            self.gibbs_files = new_gibbs
            self.update_listbox(self.scf_listbox, [f if f else "(无匹配文件)" for f in self.scf_files])
            self.update_listbox(self.gibbs_listbox, [f if f else "(无匹配文件)" for f in self.gibbs_files])
            self.status_var.set("手动匹配已应用")
            match_dialog.destroy()

        ttk.Button(control_frame, text="应用", command=apply_changes).pack(side=tk.RIGHT, padx=10)
        ttk.Button(control_frame, text="取消", command=match_dialog.destroy).pack(side=tk.RIGHT, padx=10)

    def extract_value(self, file_path, pattern):
        """从文件中提取指定值"""
        if not file_path:
            return None

        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
                match = re.search(pattern, content)
                if match:
                    return float(match.group(1))
                else:
                    return None
        except Exception as e:
            self.status_var.set(f"文件读取错误: {os.path.basename(file_path)}")
            return None

    def extract_energies(self):
        """从所有文件中提取能量数据"""
        if not self.scf_files:
            messagebox.showwarning("无文件", "请先添加SCF文件")
            return

        # 准备数据列表
        data = []
        scf_pattern = r"SCF Done:\s+E\(.*\)\s+=\s+([-\d.]+)\s+A.U."
        gibbs_pattern = r"Thermal correction to Gibbs Free Energy=\s+([-\d.]+)"

        # 为每个SCF-Gibbs对提取数据
        for i, scf_file in enumerate(self.scf_files):
            row = {}

            # 获取文件名 - 以SCF文件名为准
            if scf_file:
                scf_base = os.path.basename(scf_file)
                base_name = scf_base[3:] if scf_base.startswith("sp_") else scf_base
                row["文件名"] = os.path.splitext(base_name)[0]

                # 提取SCF值
                scf_value = self.extract_value(scf_file, scf_pattern)
                row["SCF能量(a.u.)"] = scf_value
            else:
                row["文件名"] = ""
                row["SCF能量(a.u.)"] = None

            # 获取对应的Gibbs值
            gibbs_file = self.gibbs_files[i] if i < len(self.gibbs_files) else None
            if gibbs_file:
                gibbs_value = self.extract_value(gibbs_file, gibbs_pattern)
                row["Gibbs校正(a.u.)"] = gibbs_value
            else:
                row["Gibbs校正(a.u.)"] = None

            # 计算总能量（如果两项都存在）
            if row["SCF能量(a.u.)"] is not None and row["Gibbs校正(a.u.)"] is not None:
                row["总能量(a.u.)"] = row["SCF能量(a.u.)"] + row["Gibbs校正(a.u.)"]
            else:
                row["总能量(a.u.)"] = None

            data.append(row)

        # 更新表格显示
        self.data = pd.DataFrame(data)
        self.update_table()
        self.status_var.set("能量数据提取完成")

    def update_table(self):
        """更新表格视图以显示当前数据"""
        # 清空现有内容
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 添加新数据
        for _, row in self.data.iterrows():
            scf_value = row["SCF能量(a.u.)"]
            gibbs_value = row["Gibbs校正(a.u.)"]
            total_value = row["总能量(a.u.)"]

            # 格式化显示值
            scf_display = f"{scf_value:.8f}" if scf_value is not None else ""
            gibbs_display = f"{gibbs_value:.8f}" if gibbs_value is not None else ""
            total_display = f"{total_value:.8f}" if total_value is not None else ""

            self.tree.insert("", tk.END, text=row["文件名"],
                             values=(scf_display, gibbs_display, total_display))

    def export_to_excel(self):
        """将数据导出到Excel文件，优化格式和排列"""
        if self.data.empty:
            messagebox.showwarning("无数据", "请先提取能量数据")
            return

        output_file = filedialog.asksaveasfilename(
            title="保存Excel文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )

        if not output_file:
            return

        try:
            # 创建数据副本用于格式化
            excel_data = self.data.copy()

            # 重命名列标题为更专业的名称
            excel_data.columns = ["分子名称", "SCF能量(a.u.)", "Gibbs校正(a.u.)", "总能量(a.u.)"]

            # 添加当前日期时间戳
            timestamp = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")

            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                # 创建主数据表
                excel_data.to_excel(writer, sheet_name='能量分析', index=False, startrow=3)

                # 获取工作簿和工作表对象
                workbook = writer.book
                worksheet = writer.sheets['能量分析']

                # ================= 专业格式设置 =================
                # 1. 标题格式
                title_format = workbook.add_format({
                    'bold': True,
                    'font_size': 16,
                    'align': 'center',
                    'valign': 'vcenter'
                })

                # 2. 副标题格式
                subtitle_format = workbook.add_format({
                    'italic': True,
                    'font_size': 10,
                    'align': 'center',
                    'valign': 'vcenter'
                })

                # 3. 表头格式
                header_format = workbook.add_format({
                    'bold': True,
                    'border': 1,
                    'bg_color': '#4F81BD',  # 深蓝色
                    'color': 'white',
                    'align': 'center',
                    'valign': 'vcenter'
                })

                # 4. 数据单元格格式
                data_format = workbook.add_format({
                    'border': 1,
                    'align': 'center',
                    'valign': 'vcenter',
                    'num_format': '0.00000000'  # 8位小数精度
                })

                # 5. 空单元格格式
                empty_format = workbook.add_format({
                    'border': 1,
                    'align': 'center',
                    'valign': 'vcenter',
                    'bg_color': '#F2F2F2'
                })

                # ================= 添加标题和元数据 =================
                # 主标题
                worksheet.merge_range('A1:D1', 'Gaussian 能量分析报告', title_format)

                # 副标题
                worksheet.merge_range('A2:D2', f'生成时间: {timestamp} | Gaussian版本: 09 Rev D.01', subtitle_format)

                # ================= 设置列宽 =================
                # 根据内容自动调整列宽
                worksheet.set_column('A:A', max(excel_data['分子名称'].astype(str).map(len).max(), 20))
                worksheet.set_column('B:B', 18)
                worksheet.set_column('C:C', 18)
                worksheet.set_column('D:D', 18)

                # ================= 应用表头格式 =================
                for col_num, value in enumerate(excel_data.columns.values):
                    worksheet.write(3, col_num, value, header_format)

                # ================= 应用数据格式 =================
                for row_num in range(len(excel_data)):
                    for col_num in range(len(excel_data.columns)):
                        cell_value = excel_data.iloc[row_num, col_num]

                        if pd.isna(cell_value) or cell_value is None:
                            # 空值应用特殊格式
                            worksheet.write(row_num + 4, col_num, "", empty_format)
                        elif isinstance(cell_value, (int, float)):
                            # 数值应用数据格式
                            worksheet.write(row_num + 4, col_num, cell_value, data_format)
                        else:
                            # 文本应用基本格式
                            worksheet.write(row_num + 4, col_num, cell_value, data_format)

                # ================= 添加自动筛选 =================
                worksheet.autofilter(3, 0, len(excel_data) + 3, len(excel_data.columns) - 1)

                # ================= 添加条件格式 =================
                # 对能量列添加数据条以直观显示数值大小
                for col in range(1, len(excel_data.columns)):
                    worksheet.conditional_format(
                        4, col, len(excel_data) + 3, col,
                        {
                            'type': 'data_bar',
                            'bar_color': '#5B9BD5',  # 蓝色数据条
                            'bar_border_color': '#5B9BD5'
                        }
                    )

                # ================= 添加注释工作表 =================
                notes_sheet = workbook.add_worksheet('分析说明')

                # 注释内容
                notes = [
                    ("Gaussian 能量分析报告", "标题格式"),
                    ("", ""),
                    ("本报告包含以下列：", "小标题"),
                    ("  分子名称 - 去除'sp_'前缀的文件名", "正常文本"),
                    ("  SCF能量(a.u.) - 从SCF Done行提取的电子能量", "正常文本"),
                    ("  Gibbs校正(a.u.) - 热校正吉布斯自由能", "正常文本"),
                    ("  总能量(a.u.) - SCF能量与Gibbs校正之和", "正常文本"),
                    ("", ""),
                    ("数据处理说明：", "小标题"),
                    ("  - 缺失值显示为灰色单元格", "正常文本"),
                    ("  - 所有能量值保留8位小数精度", "正常文本"),
                    ("  - 数据条可视化能量值相对大小", "正常文本"),
                    ("", ""),
                    ("使用建议：", "小标题"),
                    ("  1. 使用自动筛选功能可快速过滤数据", "正常文本"),
                    ("  2. 数据条可直观比较能量值大小", "正常文本"),
                    ("  3. 排序功能可按任意列排序", "正常文本")
                ]

                # 注释格式
                title_note = workbook.add_format({'bold': True, 'font_size': 14})
                subtitle_note = workbook.add_format({'bold': True, 'font_size': 12})
                normal_note = workbook.add_format({'text_wrap': True})

                # 写入注释
                row = 0
                for note, note_type in notes:
                    if note_type == "标题格式":
                        notes_sheet.write(row, 0, note, title_note)
                        row += 2
                    elif note_type == "小标题":
                        notes_sheet.write(row, 0, note, subtitle_note)
                        row += 1
                    else:
                        notes_sheet.write(row, 0, note, normal_note)
                        row += 1

                # 设置注释列宽
                notes_sheet.set_column('A:A', 60)

                # ================= 添加图表分析 =================
                if len(excel_data) > 1 and not excel_data['总能量(a.u.)'].isnull().all():
                    # 创建图表工作表
                    chart_sheet = workbook.add_worksheet('能量图表')

                    # 创建柱状图
                    chart = workbook.add_chart({'type': 'column'})

                    # 配置图表数据
                    chart.add_series({
                        'name': '能量分析!$D$4',
                        'categories': f'=能量分析!$A$5:$A${len(excel_data) + 4}',
                        'values': f'=能量分析!$D$5:$D${len(excel_data) + 4}',
                        'data_labels': {'value': True, 'num_format': '0.0000'},
                        'fill': {'color': '#4472C4'}
                    })

                    # 设置图表标题和样式
                    chart.set_title({'name': '分子总能量比较'})
                    chart.set_x_axis({'name': '分子名称', 'text_rotation': -45})
                    chart.set_y_axis({'name': '总能量 (a.u.)'})
                    chart.set_legend({'none': True})
                    chart.set_style(11)  # 使用预定义样式

                    # 在工作表中插入图表
                    chart_sheet.insert_chart('B2', chart, {'x_scale': 2, 'y_scale': 1.5})

                self.status_var.set(f"成功导出到: {output_file}")
                messagebox.showinfo("导出成功",
                                    f"能量数据已成功导出到:\n{output_file}\n\n"
                                    "报告包含:\n"
                                    "1. 专业格式的能量数据表\n"
                                    "2. 详细的数据说明文档\n"
                                    "3. 能量比较图表(当数据量>1时)")

                # 自动打开Excel文件
                self.open_file(output_file)

        except Exception as e:
            error_msg = f"导出Excel文件时出错:\n{str(e)}"
            messagebox.showerror("导出错误", error_msg)
            self.status_var.set("导出失败")

    def open_file(self, filepath):
        """尝试打开文件（跨平台）"""
        try:
            if platform.system() == 'Darwin':  # macOS
                os.system(f'open "{filepath}"')
            elif platform.system() == 'Windows':  # Windows
                os.startfile(filepath)
            else:  # linux variants
                os.system(f'xdg-open "{filepath}"')
        except:
            self.status_var.set("文件已保存但无法自动打开")


if __name__ == "__main__":
    root = tk.Tk()
    app = GaussianEnergyAnalyzer(root)
    root.mainloop()
