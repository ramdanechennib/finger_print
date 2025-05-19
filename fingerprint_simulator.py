import tkinter as tk
from tkinter import ttk, messagebox
import socket
import threading
import time

class FingerprintSimulator:
    def __init__(self, root):
        self.root = root
        self.root.title("محاكي جهاز البصمة")
        self.root.geometry("500x400")
        self.root.resizable(False, False)
        
        self.setup_ui()
        self.running = False
        self.server_socket = None
        self.connection = None
        
    def setup_ui(self):
        # إطار العنوان
        title_frame = tk.Frame(self.root, bg="#2c3e50")
        title_frame.pack(fill=tk.X)
        
        tk.Label(title_frame, text="محاكي جهاز البصمة", font=("Arial", 16, "bold"), 
                bg="#2c3e50", fg="white", padx=10, pady=10).pack()
        
        # إطار الإعدادات
        settings_frame = tk.LabelFrame(self.root, text="إعدادات المحاكي", font=("Arial", 12), 
                                     padx=10, pady=10)
        settings_frame.pack(pady=10, padx=10, fill=tk.X)
        
        tk.Label(settings_frame, text="رقم البصمة:").grid(row=0, column=0, padx=5, pady=5)
        self.fingerprint_id = tk.Entry(settings_frame)
        self.fingerprint_id.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(settings_frame, text="الحالة:").grid(row=1, column=0, padx=5, pady=5)
        self.status = ttk.Combobox(settings_frame, values=["AUTO", "IN", "OUT"])
        self.status.grid(row=1, column=1, padx=5, pady=5)
        self.status.current(0)
        
        # أزرار التحكم
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)
        
        self.start_btn = tk.Button(button_frame, text="تشغيل الخادم", command=self.start_server, 
                                  bg="#27ae60", fg="white", width=15)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        
        self.send_btn = tk.Button(button_frame, text="إرسال البيانات", command=self.send_data, 
                                 bg="#3498db", fg="white", width=15, state=tk.DISABLED)
        self.send_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = tk.Button(button_frame, text="إيقاف الخادم", command=self.stop_server, 
                                bg="#e74c3c", fg="white", width=15, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        # سجل الأحداث
        log_frame = tk.LabelFrame(self.root, text="سجل الأحداث", font=("Arial", 12), padx=10, pady=10)
        log_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(log_frame, height=8, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(self.log_text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.log_text.yview)
    
    def log_message(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.config(state=tk.DISABLED)
        self.log_text.see(tk.END)
    
    def start_server(self):
        try:
            self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.server_socket.bind(('127.0.0.1', 9999))
            self.server_socket.listen(1)
            
            self.running = True
            self.start_btn.config(state=tk.DISABLED)
            self.send_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.NORMAL)
            
            self.log_message("تم تشغيل الخادم على 127.0.0.1:9999")
            
            # تشغيل thread لقبول الاتصالات
            threading.Thread(target=self.accept_connections, daemon=True).start()
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في تشغيل الخادم: {str(e)}")
    
    def accept_connections(self):
        while self.running:
            try:
                self.connection, addr = self.server_socket.accept()
                self.log_message(f"تم الاتصال من: {addr}")
            except:
                if self.running:
                    self.log_message("تم إيقاف الخادم")
                break
    
    def send_data(self):
        if not self.connection:
            messagebox.showwarning("تحذير", "لا يوجد اتصال نشط")
            return
            
        fingerprint_id = self.fingerprint_id.get()
        status = self.status.get()
        
        if not fingerprint_id.isdigit():
            messagebox.showerror("خطأ", "يجب أن يكون رقم البصمة رقم صحيح")
            return
            
        message = f"{fingerprint_id},{status}"
        try:
            self.connection.sendall(message.encode('utf-8'))
            self.log_message(f"تم إرسال: {message}")
        except Exception as e:
            self.log_message(f"فشل في الإرسال: {str(e)}")
    
    def stop_server(self):
        self.running = False
        try:
            if self.connection:
                self.connection.close()
            if self.server_socket:
                self.server_socket.close()
        except:
            pass
            
        self.start_btn.config(state=tk.NORMAL)
        self.send_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)
        self.log_message("تم إيقاف الخادم")
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    root = tk.Tk()
    app = FingerprintSimulator(root)
    app.run()