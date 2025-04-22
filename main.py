#!/usr/bin/env python3 
"""
PDF智能目录生成器 (GUI版 2025-04-22)
"""
import tkinter as tk
from tempfile import TemporaryDirectory
from tkinter import ttk, filedialog, messagebox
import subprocess
import os
import shutil
import threading
import pdftocgen


class PDFTocGenerator(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF 目录生成器 v2025.4")
        self.geometry("400x600")

        # 依赖检查 
        if not self.check_dependencies():
            messagebox.showerror(" 错误", "请先安装依赖：pdf.tocgen  和 Poppler")
            self.destroy()
            return

            # 文件选择区域
        self.create_file_selector()

        # 目录编辑区域 
        self.create_toc_editor()

        # 控制按钮 
        self.create_control_buttons()

        # 状态栏 
        self.status = ttk.Label(self, text="就绪", relief=tk.SUNKEN)
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

    def create_file_selector(self):
        """创建文件选择组件"""
        frame = ttk.LabelFrame(self, text="文件选择")
        frame.pack(pady=10, fill=tk.X)

        # 输入文件 
        ttk.Button(frame, text="选择输入PDF", command=self.select_input).grid(row=0, column=0)
        self.input_entry = ttk.Entry(frame, width=40)
        self.input_entry.grid(row=0, column=1, padx=5)


    def create_toc_editor(self):
        """创建目录编辑表格"""
        frame = ttk.LabelFrame(self, text="目录条目 (级别≥1，页码≥1)")
        frame.pack(pady=10, fill=tk.BOTH, expand=True)

        # 表格列配置 
        self.tree = ttk.Treeview(frame, columns=("level", "page", "text"), show="headings")
        self.tree.heading("level", text="级别")
        self.tree.heading("page", text="页码")
        self.tree.heading("text", text="标题")
        self.tree.column("level", width=80, anchor='center')
        self.tree.column("page", width=100, anchor='center')
        self.tree.pack(fill=tk.BOTH, expand=True)

        # 编辑工具栏 
        toolbar = ttk.Frame(frame)
        toolbar.pack(pady=5)

        ttk.Button(toolbar, text="添加条目", command=self.add_entry_dialog).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="删除选中", command=self.delete_selected).pack(side=tk.LEFT)

    def create_control_buttons(self):
        """创建操作按钮"""
        frame = ttk.Frame(self)
        frame.pack(pady=10)

        ttk.Button(frame, text="生成目录", command=self.start_generation).pack(side=tk.LEFT)
        ttk.Button(frame, text="退出", command=self.destroy).pack(side=tk.RIGHT)

    def select_input(self):
        """选择输入PDF"""
        path = filedialog.askopenfilename(filetypes=[("PDF 文件", "*.pdf")])
        if path:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, path)



    def add_entry_dialog(self):
        """添加目录条目对话框"""
        dialog = tk.Toplevel(self)
        dialog.title(" 添加目录条目")

        ttk.Label(dialog, text="级别：").grid(row=0, column=0)
        level_entry = ttk.Entry(dialog)
        level_entry.grid(row=0, column=1)

        ttk.Label(dialog, text="页码：").grid(row=1, column=0)
        page_entry = ttk.Entry(dialog)
        page_entry.grid(row=1, column=1)

        ttk.Label(dialog, text="标题：").grid(row=2, column=0)
        text_entry = ttk.Entry(dialog)
        text_entry.grid(row=2, column=1)

        def validate_and_add():
            try:
                level = int(level_entry.get())
                page = int(page_entry.get())
                text = text_entry.get().strip()
                if level < 1 or page < 1 or not text:
                    raise ValueError
                self.tree.insert("", tk.END, values=(level, page, text))
                dialog.destroy()
            except:
                messagebox.showerror(" 错误", "输入无效")

        ttk.Button(dialog, text="确认", command=validate_and_add).grid(row=3, columnspan=2)

    def delete_selected(self):
        """删除选中条目"""
        for item in self.tree.selection():
            self.tree.delete(item)

    def check_dependencies(self):
        """检查系统依赖"""
        required = ['pdfxmeta', 'pdftocgen', 'pdftocio']
        return all(shutil.which(cmd) for cmd in required)

    def start_generation(self):
        """启动生成线程"""
        # 输入验证 
        input_pdf = self.input_entry.get()

        if not all([input_pdf]):
            messagebox.showerror(" 错误", "请先选择输入文件路径")
            return

        if not os.path.exists(input_pdf):
            messagebox.showerror(" 错误", "输入文件不存在")
            return

            # 收集目录条目
        headings = [self.tree.item(item)['values'] for item in self.tree.get_children()]
        if not headings:
            messagebox.showerror(" 错误", "至少需要添加一个目录条目")
            return

            # 在后台线程运行生成过程
        self.status.config(text=" 处理中...")
        threading.Thread(
            target=self.generate_toc,
            args=(input_pdf,  headings),
            daemon=True
        ).start()

    def generate_toc(self, input_pdf, headings):
        """改进版目录生成函数，直接操作目标文件"""
        # 生成输出文件名
        base_name = os.path.splitext(input_pdf)[0]
        recipe_file = f"{base_name}_recipe.toml"
        toc_file = f"{base_name}_toc"
        output_pdf = f"{base_name}_with_toc.pdf"

        try:
            # 阶段1：生成配方文件
            self._write_recipe(recipe_file, input_pdf, headings)

            # 阶段2：生成目录结构
            self._generate_toc_structure(recipe_file, toc_file, input_pdf)

            # 阶段3：嵌入目录到PDF
            self._embed_toc(toc_file, output_pdf, input_pdf)

            messagebox.showinfo(" 成功", f"文件已生成：\n{output_pdf}")
            return output_pdf

        except FileNotFoundError as e:
            messagebox.showerror(" 错误", f"文件路径不存在: {str(e)}")
        except PermissionError:
            messagebox.showerror(" 错误", "文件写入权限被拒绝")
        except subprocess.CalledProcessError as e:
            messagebox.showerror(" 命令执行失败", f"{e.cmd}\n 错误输出: {e.stderr}")
        except Exception as e:
            messagebox.showerror(" 未知错误", str(e))

    def _write_recipe(self, recipe_path, pdf_path, headings):
        """智能写入配方文件"""
        # 自动创建目录
        os.makedirs(os.path.dirname(recipe_path), exist_ok=True)

        with open(recipe_path, 'w+', encoding='utf-8') as f:  # 使用w+模式创建或覆盖
            for level, page, text in headings:
                cmd = ['pdfxmeta','-p', str(page),'-a', str(level),pdf_path,text]
                result = subprocess.run(cmd,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,check=True,timeout=30)
                # 过滤注释行并写入
                filtered = [line for line in result.stdout.split('\n')
                            if not line.strip().startswith('#')]
                f.write('\n'.join(filtered) + '\n\n')

    def _generate_toc_structure(self, recipe_path, toc_path, pdf_path):
        """生成目录中间文件"""
        os.makedirs(os.path.dirname(toc_path), exist_ok=True)

        with open(recipe_path, 'r', encoding='utf-8') as f_in, \
                open(toc_path, 'w', encoding='utf-8') as f_out:
            pdftocgen_cmd = ['pdftocgen', '-v', pdf_path]
            subprocess.run(
                pdftocgen_cmd,
                stdin=f_in,
                stdout=f_out,
                check=True,
                timeout=60
            )

    def _embed_toc(self, toc_path, output_path, input_pdf):
        """嵌入目录到PDF"""
        with open(toc_path, 'r', encoding='utf-8') as f:
            pdftocio_cmd = [
                'pdftocio',
                '-v',
                '-o', output_path,
                input_pdf
            ]
            subprocess.run(
                pdftocio_cmd,
                stdin=f,
                check=True,
                timeout=60
            )


if __name__ == '__main__':
    app = PDFTocGenerator()
    app.mainloop()