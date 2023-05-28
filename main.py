import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx import Document

def get_files(directory, extensions):
    for root, dirs, files in os.walk(directory):
        for file_name in files:
            _, ext = os.path.splitext(file_name)
            if ext in extensions:
                yield os.path.join(root, file_name)

def write_to_docx(file_paths, docx_file_path, progress_var, progress_label_var):
    doc = Document()
    for idx, file_path in enumerate(file_paths):
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()

            doc.add_heading(file_path, level=1)

            non_empty_lines = [line for line in lines if line.strip()]
            for i in range(0, len(non_empty_lines), 50):
                next_50_lines = non_empty_lines[i: i + 50]
                for line in next_50_lines:
                    doc.add_paragraph(line.rstrip())

                if i + 50 < len(non_empty_lines):
                    doc.add_page_break()

        progress_var.set(idx + 1)
        progress_label_var.set(f"正在导出文件 {idx + 1} / {len(file_paths)}：{file_path}")

    doc.save(docx_file_path)
    messagebox.showinfo("完成", "所有文件已成功导入Word文档！")

def main():
    root = tk.Tk()
    root.title("代码文件转Word")

    extensions = ['.py', '.c', '.cpp']

    directory = tk.StringVar()
    docx_file_path = tk.StringVar()

    progress_var = tk.IntVar()
    progress_var.set(0)

    progress_label_var = tk.StringVar()
    progress_label_var.set("")

    def select_directory():
        dir_path = filedialog.askdirectory()
        directory.set(dir_path)

    def select_save_path():
        save_path = filedialog.asksaveasfilename(defaultextension=".docx")
        docx_file_path.set(save_path)

    def export_to_word():
        file_paths = list(get_files(directory.get(), extensions))
        progress_bar['maximum'] = len(file_paths)
        threading.Thread(target=write_to_docx, args=(file_paths, docx_file_path.get(), progress_var, progress_label_var)).start()

    dir_select_btn = tk.Button(root, text="选择文件夹", command=select_directory)
    save_select_btn = tk.Button(root, text="选择保存路径", command=select_save_path)
    export_btn = tk.Button(root, text="导出到Word", command=export_to_word)

    progress_bar = ttk.Progressbar(root, length=200, variable=progress_var)
    progress_label = tk.Label(root, textvariable=progress_label_var)

    dir_select_btn.pack(padx=10, pady=10)
    save_select_btn.pack(padx=10, pady=10)
    export_btn.pack(padx=10, pady=10)
    progress_bar.pack(padx=10, pady=10)
    progress_label.pack(padx=10, pady=10)

    root.mainloop()

if __name__ == '__main__':
    main()
