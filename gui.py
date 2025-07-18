import tkinter as tk
from tkinter import filedialog, messagebox
import os

from readnwrite import process_files

class PGDParserGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Chuyển dữ liệu từ Word sang Excel")

        self.selected_type = tk.StringVar(value="thalassemia")
        self.file_list = []

        # Dropdown to select file type
        tk.Label(root, text="Chọn loại xét nghiệm:").pack(pady=5)
        tk.OptionMenu(root, self.selected_type, "thalassemia", "pgd").pack()

        # Button to pick files
        tk.Button(root, text="Chọn file (hiện chỉ hỗ trợ .doc/.docx)", command=self.choose_files).pack(pady=10)

        # File display
        self.text = tk.Text(root, height=10, width=60, state='disabled')
        self.text.pack()

        # Run & quit buttons
        tk.Button(root, text="Chạy", command=self.run).pack(pady=5)
        tk.Button(root, text="Thoát", command=root.quit).pack(pady=5)

    def choose_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Word files", "*.doc *.docx")])
        for f in files:
            ext = os.path.splitext(f)[1].lower()
            if ext not in [".doc", ".docx"]:
                messagebox.showerror("Lỗi", f"File không hợp lệ: {f}")
                continue
            if f not in [fp[0] for fp in self.file_list]:
                self.file_list.append((f, self.selected_type.get()))
        self.update_display()

    def update_display(self):
        self.text.configure(state='normal')
        self.text.delete(1.0, tk.END)
        for f, t in self.file_list:
            self.text.insert(tk.END, f"{f} ({t})\n")
        self.text.configure(state='disabled')

    def run(self):
        if not self.file_list:
            messagebox.showwarning("Chưa chọn file", "Vui lòng chọn ít nhất một file.")
            return

        output = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not output:
            return

        try:
            process_files(self.file_list, output)
            messagebox.showinfo("Thành công", f"Xuất file Excel: {output}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Đã xảy ra lỗi:\n{str(e)}")

        self.file_list = []
        self.update_display()

if __name__ == "__main__":
    root = tk.Tk()
    app = PGDParserGUI(root)
    root.mainloop()