import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import subprocess
import threading
import xlwt
import datetime

class PrinterInfoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Учет оргтехники")
        self.root.geometry("750x800")
        
        # Настройки приложения
        self.script_folder = "scripts"
        os.makedirs(self.script_folder, exist_ok=True)
        
        # Список всех полей (оригинальный набор)
        self.field_names = [
            "InventoryNumber", 
            "Description", 
            "Owner", 
            "Оргтехника.Цвет_печати",
            "Оргтехника.Производитель", 
            "Оргтехника.Наименование",
            "Оргтехника.Формат_Печати", 
            "Оргтехника.Технология_Печати",
            "Оргтехника.Серийный_номер", 
            "Оргтехника.Тип_подключения",
            "Оргтехника.Формат_сканирования", 
            "Оргтехника.Тип_картриджа",
            "Оргтехника.IP-адрес", 
            "Оргтехника.Тип", 
            "Оргтехника.Модель",
            "Оргтехника.Здание", 
            "Оргтехника.Помещение"
        ]
       
        self.entries = {}
        
        self.create_widgets()


    
    def create_widgets(self):
        """Создание всех элементов интерфейса"""

        mode_frame = ttk.LabelFrame(self.root, text="Источник данных", padding=10)
        mode_frame.pack(pady=5, padx=10, fill="x")
        
        self.mode_var = tk.StringVar(value="manual")
        
        ttk.Radiobutton(mode_frame, text="Ручной ввод", variable=self.mode_var,
                       value="manual", command=self.show_manual_mode).pack(anchor="w")
        ttk.Radiobutton(mode_frame, text="Загрузка с сервера", variable=self.mode_var,
                       value="server", command=self.show_server_mode).pack(anchor="w")
        
        self.manual_frame = ttk.LabelFrame(self.root, text="Данные оргтехники", padding=10)
        self.canvas = tk.Canvas(self.manual_frame)
        self.scrollbar = ttk.Scrollbar(self.manual_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        self.create_manual_fields()
        
        self.server_frame = ttk.LabelFrame(self.root, text="Загрузка с сервера", padding=10)
        self.create_server_fields()
        
        self.show_manual_mode()
        
        self.create_control_buttons()
        
        self.status_var = tk.StringVar(value="Готово")
        ttk.Label(self.root, textvariable=self.status_var, relief="sunken",
                 anchor="w").pack(side="bottom", fill="x", padx=10, pady=5)
    
    def create_manual_fields(self):
        """Создание полей для ручного ввода"""
        for i, field in enumerate(self.field_names):
            label_text = field
            
            label = ttk.Label(self.scrollable_frame, text=label_text + ":")
            label.grid(row=i, column=0, sticky="e", padx=5, pady=2)
            
            entry_var = tk.StringVar()
            entry = ttk.Entry(self.scrollable_frame, textvariable=entry_var, width=40)
            entry.grid(row=i, column=1, sticky="we", padx=5, pady=2)
            
            self.entries[field] = entry_var
    
    def create_server_fields(self):
        """Создание полей для загрузки с сервера"""
        ttk.Label(self.server_frame, text="PowerShell скрипт:").grid(row=0, column=0, sticky="e", padx=5, pady=2)
        
        self.script_path_var = tk.StringVar()
        ttk.Entry(self.server_frame, textvariable=self.script_path_var, width=40).grid(
            row=0, column=1, padx=5, pady=2, sticky="we")
        ttk.Button(self.server_frame, text="Обзор...", command=self.browse_script).grid(
            row=0, column=2, padx=5, pady=2)
        
        ttk.Button(self.server_frame, text="Выполнить скрипт", 
                  command=self.execute_powershell_script).grid(
                      row=1, column=0, columnspan=3, pady=10)
        
        self.output_text = tk.Text(self.server_frame, height=15, width=80, state="disabled")
        self.output_text.grid(row=2, column=0, columnspan=3, padx=5, pady=2)
        
        self.server_frame.grid_columnconfigure(1, weight=1)
    
    def create_control_buttons(self):
        """Создание кнопок управления"""
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)
        
        self.save_btn = ttk.Button(btn_frame, text="Сохранить в XLS", command=self.save_manual_data)
        self.save_btn.pack(side="left", padx=5)
        
        ttk.Button(btn_frame, text="Очистить", command=self.clear_fields).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Выход", command=self.root.quit).pack(side="left", padx=5)
    
    def show_manual_mode(self):
        """Переключение в режим ручного ввода"""
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)
        self.save_btn = ttk.Button(btn_frame, text="Сохранить в XLS", command=self.save_manual_data)
        self.save_btn.pack(side="left", padx=5)
        self.server_frame.pack_forget()
        self.manual_frame.pack(pady=5, padx=10, fill="both", expand=True)
        self.save_btn.config(state="normal")
    
    def show_server_mode(self):
        """Переключение в режим загрузки с сервера"""
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)
        self.save_btn = ttk.Button(btn_frame, text="Сохранить в XLS", command=self.save_manual_data)
        self.save_btn.pack(side="left", padx=5)
        self.manual_frame.pack_forget()
        self.server_frame.pack(pady=5, padx=10, fill="both", expand=True)
        self.save_btn.config(state="disabled")
    
    def browse_script(self):
        """Выбор PowerShell скрипта"""
        initial_dir = self.script_folder if os.path.exists(self.script_folder) else os.getcwd()
        file_path = filedialog.askopenfilename(
            initialdir=initial_dir,
            title="Выберите PowerShell скрипт",
            filetypes=[("PowerShell скрипты", "*.ps1"), ("Все файлы", "*.*")]
        )
        
        if file_path:
            self.script_path_var.set(file_path)
    
    def execute_powershell_script(self):
        """Выполнение PowerShell скрипта"""
        script_path = self.script_path_var.get()
        
        if not script_path:
            messagebox.showwarning("Ошибка", "Не выбран скрипт для выполнения")
            return
        
        if not os.path.exists(script_path):
            messagebox.showerror("Ошибка", f"Файл скрипта не найден:\n{script_path}")
            return
        
        self.output_text.config(state="normal")
        self.output_text.delete(1.0, tk.END)
        self.output_text.insert(tk.END, "Выполнение скрипта...\n")
        self.output_text.see(tk.END)
        self.output_text.config(state="disabled")
        
        threading.Thread(
            target=self._run_powershell_script,
            args=(script_path,),
            daemon=True
        ).start()
    
    def _run_powershell_script(self, script_path):
        """Фоновое выполнение PowerShell скрипта"""
        try:
            process = subprocess.Popen(
                ["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", script_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='windows-1251'
            )
            
            while True:
                output = process.stdout.readline()
                if output == '' and process.poll() is not None:
                    break
                if output:
                    self.update_output(output.strip())
            
            _, stderr = process.communicate()
            
            if process.returncode != 0:
                self.update_output(f"\nОшибка выполнения (код {process.returncode}):")
                if stderr:
                    self.update_output(stderr.strip())
            else:
                self.update_output("\nСкрипт успешно выполнен!")
            
        except Exception as e:
            self.update_output(f"\nОшибка: {str(e)}")
    
    def update_output(self, text):
        """Обновление текстового вывода"""
        self.output_text.config(state="normal")
        self.output_text.insert(tk.END, text + "\n")
        self.output_text.see(tk.END)
        self.output_text.config(state="disabled")
    
    def clear_fields(self):
        """Очистка всех полей ввода"""
        for var in self.entries.values():
            var.set("")
        
        self.entries["Оргтехника.Цвет_печати"].set("Цветной")
        
        self.status_var.set("Поля очищены")
    
    def save_manual_data(self):
        """Сохранение данных ручного ввода в XLS"""
        try:
            workbook = xlwt.Workbook()
            sheet = workbook.add_sheet("Оргтехника")
            
            for col, field in enumerate(self.field_names):
                sheet.write(0, col, field.replace("Оргтехника.", ""))
            
            for col, field in enumerate(self.field_names):
                sheet.write(1, col, self.entries[field].get())
            
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"OrgTech_{timestamp}.xls"
            
            workbook.save(filename)
            
            self.status_var.set(f"Данные сохранены в {filename}")
            messagebox.showinfo("Успех", f"Данные сохранены в файл:\n{os.path.abspath(filename)}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при сохранении:\n{str(e)}")
            self.status_var.set("Ошибка сохранения")

if __name__ == "__main__":
    root = tk.Tk()
    app = PrinterInfoApp(root)
    root.mainloop()
