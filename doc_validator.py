import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from docx_validator_rules import DocumentValidator

class DocValidator:
    def __init__(self, root):
        self.root = root
        self.root.title("Автопроверка шаблонов ВКР")
        self.root.geometry("1440x1080")
        self.root.resizable(True, True)
        
        self.file_path = None
        
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.file_label = ttk.Label(file_frame, text="Выберите DOCX файл:")
        self.file_label.pack(side=tk.LEFT, padx=5)
        
        self.file_entry = ttk.Entry(file_frame, width=50)
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        self.browse_button = ttk.Button(file_frame, text="Обзор", command=self.browse_file)
        self.browse_button.pack(side=tk.LEFT, padx=5)
        
        check_button = ttk.Button(main_frame, text="Проверить документ", command=self.validate_document)
        check_button.pack(pady=10)
        
        self.result_frame = ttk.LabelFrame(main_frame, text="Результаты проверки")
        self.result_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.result_scroll = ttk.Scrollbar(self.result_frame)
        self.result_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.result_text = tk.Text(self.result_frame, wrap=tk.WORD, yscrollcommand=self.result_scroll.set)
        self.result_text.pack(fill=tk.BOTH, expand=True)
        
        self.result_scroll.config(command=self.result_text.yview)

        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=100, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=5)
        
        self.status_var = tk.StringVar()
        self.status_var.set("Готов к проверке документа")
        self.status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def browse_file(self):
        """Open a file dialog to select a DOCX file"""
        file_path = filedialog.askopenfilename(
            title="Выберите DOCX файл",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        
        if file_path:
            self.file_path = file_path
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.status_var.set(f"Выбран файл: {os.path.basename(file_path)}")
    
    def validate_document(self):
        """Validate the selected DOCX file against formatting rules"""
        if not self.file_path:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите файл для проверки.")
            return
            
        if not os.path.exists(self.file_path):
            messagebox.showerror("Ошибка", "Файл не найден. Пожалуйста, выберите существующий файл.")
            return
        
        try:
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, "Начало проверки документа...\n\n")
            self.status_var.set("Проверка документа...")
            self.progress.start()
            
            self.root.update()
            
            validator = DocumentValidator(self.file_path)
            issues = validator.validate_all()
            
            self.progress.stop()
            
            if issues:
                self.result_text.insert(tk.END, f"Найдено {len(issues)} проблем:\n\n")
                for i, issue in enumerate(issues, 1):
                    self.result_text.insert(tk.END, f"{i}. {issue}\n")
                self.status_var.set(f"Проверка завершена. Найдено {len(issues)} проблем.")
            else:
                self.result_text.insert(tk.END, "Документ соответствует всем проверенным правилам форматирования!")
                self.status_var.set("Проверка завершена. Проблем не обнаружено.")

        except Exception as e:
            self.progress.stop()
            messagebox.showerror("Ошибка", f"Произошла ошибка при анализе документа: {str(e)}")
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, f"Ошибка: {str(e)}")
            self.status_var.set("Ошибка при проверке документа.")

if __name__ == "__main__":
    root = tk.Tk()
    app = DocValidator(root)
    root.mainloop()
