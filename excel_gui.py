"""
Excel 数据处理工具
功能：合并多个 Excel/CSV 文件，数据清洗，分组汇总
支持拖拽文件、进度条、记住上次输出目录等优化
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import os
import sys

# 尝试导入拖拽支持库，如果失败则提示但程序仍可运行（仅按钮选择）
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DRAG_SUPPORT = True
except ImportError:
    DRAG_SUPPORT = False
    # 降级使用普通 Tk
    print("提示：未安装 tkinterdnd2，拖拽文件功能不可用。可使用 pip install tkinterdnd2 安装。")


class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 数据处理工具")
        self.root.geometry("700x600")
        self.files = []  # 存储选中的文件路径
        self.last_output_dir = os.path.expanduser("~")  # 记住上次输出目录

        # ========== 选择文件区域 ==========
        frame_select = tk.Frame(root, padx=10, pady=10)
        frame_select.pack(fill=tk.X)

        btn_select = tk.Button(frame_select, text="📂 选择 Excel/CSV 文件（可多选）", command=self.select_files)
        btn_select.pack(side=tk.LEFT)

        self.file_label = tk.Label(frame_select, text="未选择任何文件", fg="gray")
        self.file_label.pack(side=tk.LEFT, padx=10)

        # 清空按钮
        btn_clear = tk.Button(frame_select, text="🗑️ 清空列表", command=self.clear_files)
        btn_clear.pack(side=tk.LEFT, padx=5)

        # 拖拽区域（如果支持）
        if DRAG_SUPPORT:
            self.drop_label = tk.Label(root, text="✨ 也可将文件/文件夹拖拽至此 ✨", 
                                       relief="sunken", bg="lightyellow", height=2)
            self.drop_label.pack(fill=tk.X, padx=10, pady=5)
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind('<<Drop>>', self.on_drop)

        # ========== 清洗选项 ==========
        frame_clean = tk.LabelFrame(root, text="数据清洗选项", padx=10, pady=5)
        frame_clean.pack(fill=tk.X, padx=10, pady=5)

        self.dropna_var = tk.BooleanVar(value=True)
        tk.Checkbutton(frame_clean, text="删除完全空白的行", variable=self.dropna_var).pack(anchor=tk.W)

        self.drop_dup_var = tk.BooleanVar(value=True)
        tk.Checkbutton(frame_clean, text="删除重复行（基于所有列）", variable=self.drop_dup_var).pack(anchor=tk.W)

        # 自定义删除列
        tk.Label(frame_clean, text="要删除的列名（逗号分隔，留空则不删除）：").pack(anchor=tk.W)
        self.drop_cols_entry = tk.Entry(frame_clean, width=60)
        self.drop_cols_entry.pack(anchor=tk.W, fill=tk.X, pady=2)

        # ========== 汇总选项 ==========
        frame_summary = tk.LabelFrame(root, text="分组汇总（可选）", padx=10, pady=5)
        frame_summary.pack(fill=tk.X, padx=10, pady=5)

        self.summary_var = tk.BooleanVar(value=False)
        self.summary_check = tk.Checkbutton(frame_summary, text="启用分组汇总", 
                                            variable=self.summary_var, command=self.toggle_summary)
        self.summary_check.pack(anchor=tk.W)

        self.summary_frame = tk.Frame(frame_summary)
        self.summary_frame.pack(fill=tk.X, pady=5)

        tk.Label(self.summary_frame, text="分组列名：").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.group_col_entry = tk.Entry(self.summary_frame, width=20)
        self.group_col_entry.grid(row=0, column=1, padx=5)
        self.group_col_entry.config(state='disabled')

        tk.Label(self.summary_frame, text="数值列名（可选，用于求和/平均）：").grid(row=1, column=0, sticky=tk.W, padx=5)
        self.value_col_entry = tk.Entry(self.summary_frame, width=20)
        self.value_col_entry.grid(row=1, column=1, padx=5)
        self.value_col_entry.config(state='disabled')

        # ========== 日志区域 ==========
        log_frame = tk.LabelFrame(root, text="处理日志", padx=10, pady=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=12)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # ========== 进度条 ==========
        self.progress = ttk.Progressbar(root, mode='indeterminate')
        # 先不 pack，需要时再显示

        # ========== 底部按钮 ==========
        btn_frame = tk.Frame(root)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        self.process_btn = tk.Button(btn_frame, text="🚀 开始处理", command=self.process, 
                                     bg="lightgreen", font=("", 12))
        self.process_btn.pack(side=tk.LEFT, padx=5)

        # ========== 初始状态 ==========
        self.toggle_summary()  # 确保汇总控件初始禁用

    def toggle_summary(self):
        """启用/禁用汇总相关输入框"""
        if self.summary_var.get():
            self.group_col_entry.config(state='normal')
            self.value_col_entry.config(state='normal')
        else:
            self.group_col_entry.config(state='disabled')
            self.value_col_entry.config(state='disabled')

    def select_files(self):
        """通过对话框选择多个文件"""
        files = filedialog.askopenfilenames(
            title="选择一个或多个 Excel/CSV 文件",
            filetypes=[("支持的文件", "*.xlsx *.xls *.csv"), 
                       ("Excel 文件", "*.xlsx"), 
                       ("Excel 97-2003", "*.xls"),
                       ("CSV 文件", "*.csv"),
                       ("所有文件", "*.*")]
        )
        if files:
            self.files = list(files)
            self.update_file_label()
            self.log(f"选择了 {len(self.files)} 个文件：")
            for f in self.files:
                self.log(f"  - {os.path.basename(f)}")
        else:
            self.files = []
            self.update_file_label()

    def clear_files(self):
        """清空已选文件列表"""
        self.files = []
        self.update_file_label()
        self.log("已清空文件列表")

    def update_file_label(self):
        """更新文件显示标签"""
        count = len(self.files)
        if count == 0:
            self.file_label.config(text="未选择任何文件", fg="gray")
        else:
            self.file_label.config(text=f"已选择 {count} 个文件", fg="black")

    def on_drop(self, event):
        """处理拖拽文件"""
        # event.data 可能是带花括号的路径（含空格），或普通路径
        raw = event.data
        # 使用 tkinterdnd2 的 splitlist 方法解析
        files = self.root.tk.splitlist(raw)
        valid_files = []
        for f in files:
            # 去除可能的花括号
            f = f.strip('{}')
            if os.path.isfile(f) and f.lower().endswith(('.xlsx', '.xls', '.csv')):
                valid_files.append(f)
            elif os.path.isdir(f):
                # 如果是文件夹，则遍历其中所有支持的文件
                for root_dir, dirs, filenames in os.walk(f):
                    for fn in filenames:
                        if fn.lower().endswith(('.xlsx', '.xls', '.csv')):
                            valid_files.append(os.path.join(root_dir, fn))
        if valid_files:
            # 合并去重（防止重复拖入同一文件）
            existing = set(self.files)
            new_files = [vf for vf in valid_files if vf not in existing]
            self.files.extend(new_files)
            self.update_file_label()
            self.log(f"拖拽添加了 {len(new_files)} 个文件")
            for f in new_files[:5]:  # 最多显示5个
                self.log(f"  - {os.path.basename(f)}")
            if len(new_files) > 5:
                self.log(f"  ... 等 {len(new_files)} 个文件")
        else:
            self.log("拖拽中没有找到支持的 Excel/CSV 文件")

    def log(self, msg):
        """向日志区域添加消息"""
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def process(self):
        """核心处理逻辑：合并、清洗、汇总"""
        if not self.files:
            messagebox.showwarning("警告", "请先选择 Excel/CSV 文件")
            return

        # 输出文件保存路径
        output_file = filedialog.asksaveasfilename(
            initialdir=self.last_output_dir,
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
            title="保存合并后的文件为"
        )
        if not output_file:
            return
        self.last_output_dir = os.path.dirname(output_file)

        # 显示进度条
        self.progress.pack(fill=tk.X, padx=10, pady=5)
        self.progress.start()
        self.process_btn.config(state='disabled')
        self.log("开始处理...")

        try:
            # ========== 1. 读取所有文件 ==========
            all_dfs = []
            for file in self.files:
                self.log(f"正在读取: {os.path.basename(file)}...")
                # 根据扩展名选择读取方式
                ext = os.path.splitext(file)[1].lower()
                try:
                    if ext == '.csv':
                        # 尝试多种编码
                        try:
                            df = pd.read_csv(file, encoding='utf-8')
                        except UnicodeDecodeError:
                            df = pd.read_csv(file, encoding='gbk')
                    else:
                        # Excel 文件，处理 .xls 可能需要 xlrd
                        try:
                            df = pd.read_excel(file)
                        except ImportError as e:
                            if 'xlrd' in str(e).lower():
                                self.log("⚠️ 读取 .xls 文件需要安装 xlrd，请运行: pip install xlrd")
                                raise
                            else:
                                raise
                except Exception as e:
                    self.log(f"❌ 读取文件失败 {os.path.basename(file)}: {str(e)}")
                    continue

                # 添加来源列
                df['来源文件'] = os.path.basename(file)
                all_dfs.append(df)

            if not all_dfs:
                raise ValueError("没有成功读取任何文件")

            # ========== 2. 合并 ==========
            self.log("正在合并数据...")
            merged = pd.concat(all_dfs, ignore_index=True)
            self.log(f"合并后总行数: {len(merged)}")

            # ========== 3. 清洗 ==========
            # 删除完全空行
            if self.dropna_var.get():
                before = len(merged)
                merged.dropna(how='all', inplace=True)
                after = len(merged)
                self.log(f"删除完全空白行: {before - after} 行")

            # 删除重复行
            if self.drop_dup_var.get():
                before = len(merged)
                merged.drop_duplicates(inplace=True)
                after = len(merged)
                self.log(f"删除重复行: {before - after} 行")

            # 删除指定列
            drop_cols_text = self.drop_cols_entry.get().strip()
            if drop_cols_text:
                drop_cols = [c.strip() for c in drop_cols_text.split(',') if c.strip()]
                existing_cols = [c for c in drop_cols if c in merged.columns]
                if existing_cols:
                    merged.drop(columns=existing_cols, inplace=True)
                    self.log(f"已删除列: {', '.join(existing_cols)}")
                else:
                    self.log(f"警告: 要删除的列 {drop_cols} 在数据中不存在")

            # ========== 4. 分组汇总 ==========
            summary_df = None
            if self.summary_var.get():
                group_col = self.group_col_entry.get().strip()
                if not group_col:
                    raise ValueError("请输入分组列名")
                if group_col not in merged.columns:
                    raise ValueError(f"列 '{group_col}' 不存在于数据中")

                value_col = self.value_col_entry.get().strip()
                if value_col:
                    if value_col not in merged.columns:
                        raise ValueError(f"列 '{value_col}' 不存在于数据中")
                    # 确保数值列是数字
                    merged[value_col] = pd.to_numeric(merged[value_col], errors='coerce')
                    summary_df = merged.groupby(group_col).agg(
                        数据量=('来源文件', 'count'),
                        总和=(value_col, 'sum'),
                        平均值=(value_col, 'mean')
                    ).reset_index()
                    self.log(f"已按 '{group_col}' 分组，对 '{value_col}' 进行汇总")
                else:
                    # 只计数
                    summary_df = merged.groupby(group_col).size().reset_index(name='数据量')
                    self.log(f"已按 '{group_col}' 分组，统计每组的行数")

            # ========== 5. 保存到 Excel ==========
            self.log("正在保存结果...")
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                merged.to_excel(writer, sheet_name='合并数据', index=False)
                if summary_df is not None:
                    summary_df.to_excel(writer, sheet_name='汇总结果', index=False)

            self.log(f"✅ 处理完成！结果保存至：{output_file}")
            messagebox.showinfo("完成", f"处理成功！\n共处理 {len(all_dfs)} 个文件\n最终数据 {len(merged)} 行\n保存路径：{output_file}")

        except Exception as e:
            self.log(f"❌ 发生错误: {str(e)}")
            messagebox.showerror("错误", f"处理失败：\n{str(e)}")
        finally:
            self.progress.stop()
            self.progress.pack_forget()
            self.process_btn.config(state='normal')


def main():
    if DRAG_SUPPORT:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()