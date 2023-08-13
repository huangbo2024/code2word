
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx import Document
from docx.shared import Pt, RGBColor, Inches

def get_files(directory, extensions):
    for root, dirs, files in os.walk(directory):
        for file_name in files:
            _, ext = os.path.splitext(file_name)
            if ext in extensions:
                relative_path = os.path.relpath(os.path.join(root, file_name), directory)
                yield relative_path

def set_font_style(paragraph):
    for run in paragraph.runs:
        font = run.font
        font.name = 'Times New Roman'
        font.size = Pt(10.5)
        font.color.rgb = RGBColor(0, 0, 0)
    paragraph.line_spacing = Pt(12)
    paragraph.space_after = Pt(0)
    paragraph.space_before = Pt(0)

def write_to_docx(directory, file_paths, docx_file_path, progress_var, progress_label_var):
    doc = Document()
    # Set page margins to narrow
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    for idx, file_path in enumerate(file_paths):
        with open(os.path.join(directory, file_path), 'r', encoding='utf-8') as file:
            lines = file.readlines()

            non_empty_lines = [line for line in lines if line.strip()]
            for i in range(0, len(non_empty_lines), 50):
                next_50_lines = non_empty_lines[i: i + 50]
                for line in next_50_lines:
                    para = doc.add_paragraph(line.rstrip())
                    set_font_style(para)

                if i + 50 < len(non_empty_lines):
                    doc.add_page_break()

        progress_var.set(idx + 1)
        progress_label_var.set(f"正在导出文件 {idx + 1} / {len(file_paths)}")

    doc.save(docx_file_path)
    messagebox.showinfo("完成", f"所有文件已成功导入到 {docx_file_path}！")

def main():
    global root  # Making root global to use in write_to_docx function
    root = tk.Tk()
    root.title("代码文件转Word")

    extensions = ['.py', '.c', '.cpp']

    directory = tk.StringVar()
    docx_file_path = tk.StringVar(value=os.path.expanduser("~/Documents/exported_code.docx"))

    progress_var = tk.IntVar()
    progress_var.set(0)

    progress_label_var = tk.StringVar()
    progress_label_var.set("")

    def select_directory():
        dir_path = filedialog.askdirectory()
        directory.set(dir_path)
        dir_label_var.set(f"选择的文件夹: {dir_path}")

    def select_save_path():
        save_path = filedialog.asksaveasfilename(defaultextension=".docx", initialfile="exported_code.docx")
        docx_file_path.set(save_path)
        save_label_var.set(f"保存路径: {save_path}")

    def export_to_word():
        file_paths = list(get_files(directory.get(), extensions))
        if os.path.exists(docx_file_path.get()):
            overwrite = messagebox.askyesno("文件已存在", "所选路径已存在同名文件。是否覆盖？")
            if not overwrite:
                return
        confirm = messagebox.askyesno("确认导出", f"即将导出 {len(file_paths)} 个文件到 {docx_file_path.get()}。是否继续？")
        if confirm:
            progress_bar['maximum'] = len(file_paths)
            threading.Thread(target=write_to_docx, args=(directory.get(), file_paths, docx_file_path.get(), progress_var, progress_label_var)).start()

    dir_select_btn = tk.Button(root, text="选择文件夹", command=select_directory)
    save_select_btn = tk.Button(root, text="选择保存路径", command=select_save_path)
    export_btn = tk.Button(root, text="导出到Word", command=export_to_word)

    progress_bar = ttk.Progressbar(root, length=200, variable=progress_var)
    progress_label = tk.Label(root, textvariable=progress_label_var)

    dir_label_var = tk.StringVar()
    dir_label = tk.Label(root, textvariable=dir_label_var)

    save_label_var = tk.StringVar()
    save_label = tk.Label(root, textvariable=save_label_var)

    dir_select_btn.pack(padx=10, pady=5)
    dir_label.pack(padx=10, pady=5)
    save_select_btn.pack(padx=10, pady=5)
    save_label.pack(padx=10, pady=5)
    export_btn.pack(padx=10, pady=5)
    progress_bar.pack(padx=10, pady=10)
    progress_label.pack(padx=10, pady=10)

    root.mainloop()

if __name__ == '__main__':
    main()
