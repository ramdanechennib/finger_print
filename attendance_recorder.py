import tkinter as tk
from tkinter import ttk, messagebox
import mysql.connector
from datetime import datetime, timedelta
import socket
import threading
import pandas as pd
from tkinter import filedialog

class AttendanceRecorder:
    def __init__(self, root):
        self.root = root
        self.root.title("نظام تسجيل الحضور والانصراف")
        self.root.geometry("900x600")
        
        self.setup_ui()
        self.setup_db()
        self.running = False
        self.client_socket = None
        
    def setup_db(self):
        try:
            self.db = mysql.connector.connect(
                user='root',
                password='',
                host='localhost',
                database='attendance_system'
            )
            self.cursor = self.db.cursor()
            messagebox.showinfo("نجاح", "تم الاتصال بقاعدة البيانات بنجاح")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في الاتصال بقاعدة البيانات: {str(e)}")
            self.root.destroy()
    
    def setup_ui(self):
        # إطار العنوان
        title_frame = tk.Frame(self.root, bg="#2c3e50")
        title_frame.pack(fill=tk.X)
        
        tk.Label(title_frame, text="نظام تسجيل الحضور والانصراف بالبصمة", font=("Arial", 16, "bold"), 
                bg="#2c3e50", fg="white", padx=10, pady=10).pack()
        
        # إطار التحكم
        control_frame = tk.Frame(self.root, padx=10, pady=10)
        control_frame.pack(fill=tk.X)
        
        self.connect_btn = tk.Button(control_frame, text="الاتصال بجهاز البصمة", 
                                   command=self.connect_device, bg="#27ae60", fg="white")
        self.connect_btn.pack(side=tk.LEFT, padx=5)
        
        self.disconnect_btn = tk.Button(control_frame, text="قطع الاتصال", 
                                     command=self.disconnect_device, bg="#e74c3c", fg="white", 
                                     state=tk.DISABLED)
        self.disconnect_btn.pack(side=tk.LEFT, padx=5)
        
        self.export_btn = tk.Button(control_frame, text="تصدير إلى Excel", 
                                  command=self.open_export_dialog, bg="#3498db", fg="white")
        self.export_btn.pack(side=tk.RIGHT, padx=5)
        
        self.manual_btn = tk.Button(control_frame, text="إدخال يدوي", 
                                  command=self.open_manual_entry, bg="#3498db", fg="white")
        self.manual_btn.pack(side=tk.RIGHT, padx=5)
        
        # إطار عرض البيانات
        data_frame = tk.Frame(self.root)
        data_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # سجل الأحداث
        log_frame = tk.LabelFrame(data_frame, text="سجل الأحداث", font=("Arial", 12))
        log_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        self.log_text = tk.Text(log_frame, height=15, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = tk.Scrollbar(self.log_text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.log_text.yview)
        
        # سجل الحضور
        attendance_frame = tk.LabelFrame(data_frame, text="سجل الحضور", font=("Arial", 12))
        attendance_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        
        columns = ("id", "employee", "date", "time", "status")
        self.attendance_tree = ttk.Treeview(attendance_frame, columns=columns, show="headings", height=15)
        
        self.attendance_tree.heading("id", text="المعرف")
        self.attendance_tree.heading("employee", text="الموظف")
        self.attendance_tree.heading("date", text="التاريخ")
        self.attendance_tree.heading("time", text="الوقت")
        self.attendance_tree.heading("status", text="الحالة")
        
        self.attendance_tree.column("id", width=50, anchor=tk.CENTER)
        self.attendance_tree.column("employee", width=150, anchor=tk.CENTER)
        self.attendance_tree.column("date", width=100, anchor=tk.CENTER)
        self.attendance_tree.column("time", width=100, anchor=tk.CENTER)
        self.attendance_tree.column("status", width=80, anchor=tk.CENTER)
        
        self.attendance_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(self.attendance_tree, orient=tk.VERTICAL, command=self.attendance_tree.yview)
        self.attendance_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # تحديث سجل الحضور
        self.refresh_attendance()
    
    def get_last_status(self, employee_id):
        try:
            self.cursor.execute("""
                SELECT status FROM attendance 
                WHERE employee_id = %s 
                ORDER BY date DESC, time DESC 
                LIMIT 1
            """, (employee_id,))
            result = self.cursor.fetchone()
            return result[0] if result else None
        except Exception as e:
            self.log_message(f"خطأ في التحقق من آخر حالة: {str(e)}")
            return None
    
    def log_message(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.config(state=tk.DISABLED)
        self.log_text.see(tk.END)
    
    def refresh_attendance(self):
        try:
            self.attendance_tree.delete(*self.attendance_tree.get_children())
            self.cursor.execute("""
                SELECT a.id, e.name, a.date, a.time, a.status 
                FROM attendance a
                JOIN employees e ON a.employee_id = e.id
                ORDER BY a.date DESC, a.time DESC
                LIMIT 100
            """)
            
            for row in self.cursor.fetchall():
                self.attendance_tree.insert("", tk.END, values=row)
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في تحديث سجل الحضور: {str(e)}")
    
    def connect_device(self):
        try:
            self.client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.client_socket.connect(('127.0.0.1', 9999))
            
            self.running = True
            self.connect_btn.config(state=tk.DISABLED)
            self.disconnect_btn.config(state=tk.NORMAL)
            
            self.log_message("تم الاتصال بجهاز البصمة على 127.0.0.1:9999")
            
            # تشغيل thread لاستقبال البيانات
            threading.Thread(target=self.receive_data, daemon=True).start()
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في الاتصال بجهاز البصمة: {str(e)}")
    
    def receive_data(self):
        while self.running:
            try:
                data = self.client_socket.recv(1024)
                if not data:
                    break
                    
                message = data.decode('utf-8').strip()
                self.log_message(f"تم استقبال: {message}")
                
                try:
                    fingerprint_id_str, status = message.split(',')
                    fingerprint_id = int(fingerprint_id_str)
                except ValueError:
                    self.log_message("تنسيق بيانات غير صالح")
                    continue
                
                self.process_attendance(fingerprint_id, status)
                
            except Exception as e:
                if self.running:
                    self.log_message(f"خطأ في استقبال البيانات: {str(e)}")
                break
    
    def process_attendance(self, fingerprint_id, status):
        try:
            self.cursor.execute("SELECT id, name FROM employees WHERE fingerprint_id = %s", (fingerprint_id,))
            result = self.cursor.fetchone()
            
            if result:
                employee_id, employee_name = result
                now = datetime.now()
                date = now.date()
                time_ = now.time()
                
                # الحصول على آخر حالة للموظف
                last_status = self.get_last_status(employee_id)
                
                # تحديد الحالة التلقائية إذا لم يتم تحديدها يدوياً
                if status == "AUTO":
                    if last_status == "IN":
                        status = "OUT"
                    else:
                        status = "IN"
                
                # منع تكرار نفس الحالة مرتين متتاليتين
                if last_status == status:
                    self.log_message(f"تحذير: الموظف {employee_name} لديه بالفعل حالة {status}")
                    return
                
                self.cursor.execute(
                    "INSERT INTO attendance (employee_id, date, time, status) VALUES (%s, %s, %s, %s)",
                    (employee_id, date, time_, status)
                )
                self.db.commit()
                
                self.log_message(f"تم تسجيل {status} للموظف {employee_name}")
                self.refresh_attendance()
            else:
                self.log_message(f"لا يوجد موظف مسجل برقم البصمة {fingerprint_id}")
                
        except Exception as e:
            self.log_message(f"خطأ في تسجيل الحضور: {str(e)}")
    
    def disconnect_device(self):
        self.running = False
        try:
            if self.client_socket:
                self.client_socket.close()
        except:
            pass
            
        self.connect_btn.config(state=tk.NORMAL)
        self.disconnect_btn.config(state=tk.DISABLED)
        self.log_message("تم قطع الاتصال بجهاز البصمة")
    
    def open_manual_entry(self):
        manual_window = tk.Toplevel(self.root)
        manual_window.title("إدخال بصمة يدوي")
        manual_window.geometry("400x300")
        
        tk.Label(manual_window, text="إدخال بصمة يدوي", font=("Arial", 14, "bold")).pack(pady=10)
        
        form_frame = tk.Frame(manual_window, padx=10, pady=10)
        form_frame.pack()
        
        tk.Label(form_frame, text="رقم البصمة:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
        fingerprint_entry = tk.Entry(form_frame)
        fingerprint_entry.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(form_frame, text="الحالة:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
        status_combobox = ttk.Combobox(form_frame, values=["AUTO", "IN", "OUT"])
        status_combobox.grid(row=1, column=1, padx=5, pady=5)
        status_combobox.current(0)
        
        def submit():
            fingerprint_id = fingerprint_entry.get()
            status = status_combobox.get()
            
            if not fingerprint_id.isdigit():
                messagebox.showerror("خطأ", "يجب أن يكون رقم البصمة رقم صحيح")
                return
                
            self.process_attendance(int(fingerprint_id), status)
            manual_window.destroy()
        
        button_frame = tk.Frame(manual_window)
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="إرسال", command=submit, bg="#27ae60", fg="white", width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="إلغاء", command=manual_window.destroy, bg="#e74c3c", fg="white", width=10).pack(side=tk.LEFT, padx=5)
    
    def open_export_dialog(self):
        export_window = tk.Toplevel(self.root)
        export_window.title("تصدير إلى Excel")
        export_window.geometry("400x250")
        
        tk.Label(export_window, text="تصدير سجل الحضور", font=("Arial", 14, "bold")).pack(pady=10)
        
        form_frame = tk.Frame(export_window, padx=10, pady=10)
        form_frame.pack()
        
        tk.Label(form_frame, text="الفترة الزمنية:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
        period_combobox = ttk.Combobox(form_frame, values=["اليوم", "الأسبوع", "الشهر"])
        period_combobox.grid(row=0, column=1, padx=5, pady=5)
        period_combobox.current(0)
        
        def export():
            period = period_combobox.get()
            end_date = datetime.now().date()
            
            if period == "اليوم":
                start_date = end_date
            elif period == "الأسبوع":
                start_date = end_date - timedelta(days=6)
            elif period == "الشهر":
                start_date = end_date - timedelta(days=30)
            else:
                messagebox.showerror("خطأ", "فترة زمنية غير صالحة")
                return
            
            try:
                # استعلام قاعدة البيانات للحصول على البيانات
                self.cursor.execute("""
                    SELECT e.name, a.date, a.time, a.status 
                    FROM attendance a
                    JOIN employees e ON a.employee_id = e.id
                    WHERE a.date BETWEEN %s AND %s
                    ORDER BY a.date, a.time
                """, (start_date, end_date))
                
                data = self.cursor.fetchall()
                
                if not data:
                    messagebox.showinfo("معلومات", "لا توجد بيانات للتصدير في الفترة المحددة")
                    return
                
                # تحويل البيانات إلى DataFrame
                df = pd.DataFrame(data, columns=["الموظف", "التاريخ", "الوقت", "الحالة"])
                
                # فتح حوار لحفظ الملف
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                    title="حفظ ملف Excel"
                )
                
                if file_path:
                    # تصدير إلى Excel
                    df.to_excel(file_path, index=False, engine='openpyxl')
                    messagebox.showinfo("نجاح", f"تم تصدير البيانات بنجاح إلى:\n{file_path}")
                    export_window.destroy()
                
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل في تصدير البيانات: {str(e)}")
        
        button_frame = tk.Frame(export_window)
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="تصدير", command=export, bg="#27ae60", fg="white", width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="إلغاء", command=export_window.destroy, bg="#e74c3c", fg="white", width=10).pack(side=tk.LEFT, padx=5)
    
    def run(self):
        self.root.mainloop()
        try:
            if self.running:
                self.disconnect_device()
            if self.db:
                self.db.close()
        except:
            pass

if __name__ == "__main__":
    root = tk.Tk()
    app = AttendanceRecorder(root)
    app.run()