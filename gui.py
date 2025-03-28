import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import queue
import os
from generate_dict import generate_password_dict
from crack_xlsx import crack_xlsx

class PasswordCrackerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 密碼破解工具")
        self.root.geometry("600x500")
        
        # 建立訊息佇列用於執行緒間通訊
        self.msg_queue = queue.Queue()
        
        # 建立主框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 密碼字典生成區塊
        self.create_dict_frame()
        
        # Excel 破解區塊
        self.create_crack_frame()
        
        # 狀態顯示區
        self.create_status_frame()
        
        # 定期檢查訊息佇列
        self.check_queue()
        
    def create_dict_frame(self):
        dict_frame = ttk.LabelFrame(self.main_frame, text="密碼字典生成", padding="5")
        dict_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # 最小密碼長度
        ttk.Label(dict_frame, text="最小密碼長度:").grid(row=0, column=0, padx=5)
        self.min_length = ttk.Entry(dict_frame, width=10)
        self.min_length.insert(0, "4")
        self.min_length.grid(row=0, column=1, padx=5)
        
        # 最大密碼長度
        ttk.Label(dict_frame, text="最大密碼長度:").grid(row=0, column=2, padx=5)
        self.max_length = ttk.Entry(dict_frame, width=10)
        self.max_length.insert(0, "8")
        self.max_length.grid(row=0, column=3, padx=5)
        
        # 輸出檔案選擇
        ttk.Label(dict_frame, text="輸出檔案:").grid(row=1, column=0, padx=5)
        self.output_file = ttk.Entry(dict_frame, width=40)
        self.output_file.insert(0, "password_dict.txt")
        self.output_file.grid(row=1, column=1, columnspan=2, padx=5)
        
        self.browse_btn = ttk.Button(dict_frame, text="瀏覽", command=self.browse_output)
        self.browse_btn.grid(row=1, column=3, padx=5)
        
        # 生成按鈕
        self.generate_btn = ttk.Button(dict_frame, text="生成密碼字典", command=self.start_generate)
        self.generate_btn.grid(row=2, column=0, columnspan=4, pady=5)
        
    def create_crack_frame(self):
        crack_frame = ttk.LabelFrame(self.main_frame, text="Excel 檔案破解", padding="5")
        crack_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Excel 檔案選擇
        ttk.Label(crack_frame, text="Excel 檔案:").grid(row=0, column=0, padx=5)
        self.excel_file = ttk.Entry(crack_frame, width=40)
        self.excel_file.grid(row=0, column=1, columnspan=2, padx=5)
        
        self.browse_excel_btn = ttk.Button(crack_frame, text="瀏覽", command=self.browse_excel)
        self.browse_excel_btn.grid(row=0, column=3, padx=5)
        
        # 密碼字典選擇
        ttk.Label(crack_frame, text="密碼字典:").grid(row=1, column=0, padx=5)
        self.dict_file = ttk.Entry(crack_frame, width=40)
        self.dict_file.insert(0, "password_dict.txt")
        self.dict_file.grid(row=1, column=1, columnspan=2, padx=5)
        
        self.browse_dict_btn = ttk.Button(crack_frame, text="瀏覽", command=self.browse_dict)
        self.browse_dict_btn.grid(row=1, column=3, padx=5)
        
        # 破解按鈕
        self.crack_btn = ttk.Button(crack_frame, text="開始破解", command=self.start_crack)
        self.crack_btn.grid(row=2, column=0, columnspan=4, pady=5)
        
    def create_status_frame(self):
        status_frame = ttk.LabelFrame(self.main_frame, text="執行狀態", padding="5")
        status_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        self.status_text = tk.Text(status_frame, height=10, width=60)
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar = ttk.Scrollbar(status_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.status_text['yscrollcommand'] = scrollbar.set
        
    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if filename:
            self.output_file.delete(0, tk.END)
            self.output_file.insert(0, filename)
            
    def browse_excel(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.excel_file.delete(0, tk.END)
            self.excel_file.insert(0, filename)
            
    def browse_dict(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if filename:
            self.dict_file.delete(0, tk.END)
            self.dict_file.insert(0, filename)
            
    def log_message(self, message):
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        
    def check_queue(self):
        try:
            while True:
                message = self.msg_queue.get_nowait()
                self.log_message(message)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.check_queue)
            
    def start_generate(self):
        try:
            min_len = int(self.min_length.get())
            max_len = int(self.max_length.get())
            output_file = self.output_file.get()
            
            if min_len < 1:
                messagebox.showerror("錯誤", "最小密碼長度必須大於 0")
                return
            if max_len < min_len:
                messagebox.showerror("錯誤", "最大密碼長度必須大於或等於最小密碼長度")
                return
                
            self.generate_btn.state(['disabled'])
            self.log_message("開始生成密碼字典...")
            
            def generate_thread():
                try:
                    generate_password_dict(min_len, max_len, output_file)
                    self.msg_queue.put("密碼字典生成完成！")
                except Exception as e:
                    self.msg_queue.put(f"錯誤：{str(e)}")
                finally:
                    self.root.after(0, lambda: self.generate_btn.state(['!disabled']))
                    
            threading.Thread(target=generate_thread, daemon=True).start()
            
        except ValueError:
            messagebox.showerror("錯誤", "請輸入有效的數字")
            self.generate_btn.state(['!disabled'])
            
    def start_crack(self):
        excel_file = self.excel_file.get()
        dict_file = self.dict_file.get()
        
        if not excel_file:
            messagebox.showerror("錯誤", "請選擇 Excel 檔案")
            return
        if not os.path.exists(excel_file):
            messagebox.showerror("錯誤", "找不到 Excel 檔案")
            return
        if not os.path.exists(dict_file):
            messagebox.showerror("錯誤", "找不到密碼字典檔案")
            return
            
        self.crack_btn.state(['disabled'])
        self.log_message("開始破解 Excel 檔案...")
        
        def crack_thread():
            try:
                password = crack_xlsx(excel_file, dict_file)
                if password:
                    self.msg_queue.put(f"破解成功！密碼是：{password}")
                else:
                    self.msg_queue.put("未能找到正確的密碼")
            except Exception as e:
                self.msg_queue.put(f"錯誤：{str(e)}")
            finally:
                self.root.after(0, lambda: self.crack_btn.state(['!disabled']))
                
        threading.Thread(target=crack_thread, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = PasswordCrackerGUI(root)
    root.mainloop() 