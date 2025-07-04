import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta, time  # تأكد من استيراد هذه
import socket
import threading
import pandas as pd
from tkinter import filedialog
import pymysql
from pymysql.err import MySQLError
class AttendanceRecorder:
    def __init__(self, root):
        self.root = root
        self.root.title("نظام تسجيل الحضور والانصراف المتقدم")
        self.root.geometry("1200x800")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # إعداد واجهة المستخدم وقاعدة البيانات
        self.setup_ui()
        self.setup_db()
        
        # حالة النظام
        self.running = False
        self.client_socket = None
        self.server_thread = None
        
        # تحديث الواجهة
        self.refresh_attendance()
        self.check_absences()
        
    def setup_db(self):
        """إعداد اتصال قاعدة البيانات وإنشاء الجداول إذا لزم الأمر"""
        try:
            self.db = pymysql.connect(
                user='root',
                password='',
                # host='10.18.101.136',
                database='attendance_system',
                autocommit=True,
                charset='utf8mb4',
                cursorclass=pymysql.cursors.DictCursor
            )
            self.cursor = self.db.cursor()
            self.initialize_tables()
            self.log_message("تم الاتصال بقاعدة البيانات بنجاح")
        except MySQLError as e:
            messagebox.showerror("خطأ في الاتصال", f"فشل في الاتصال بقاعدة البيانات:\n{str(e)}")
            self.root.destroy()

    def initialize_tables(self):
        """إنشاء الجداول اللازمة إذا لم تكن موجودة"""
        queries = [
            """CREATE TABLE IF NOT EXISTS employees (
                id INT PRIMARY KEY AUTO_INCREMENT,
                fingerprint_id INT UNIQUE,
                name VARCHAR(255) NOT NULL,
                department VARCHAR(255),
                position VARCHAR(255),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )""",
            """CREATE TABLE IF NOT EXISTS attendance (
                id INT PRIMARY KEY AUTO_INCREMENT,
                employee_id INT,
                date DATE NOT NULL,
                time TIME NOT NULL,
                status ENUM('IN', 'OUT') NOT NULL,
                FOREIGN KEY (employee_id) REFERENCES employees(id),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                INDEX (employee_id, date)
            )""",
            """CREATE TABLE IF NOT EXISTS work_hours (
                id INT PRIMARY KEY AUTO_INCREMENT,
                employee_id INT,
                date DATE NOT NULL,
                total_hours FLOAT NOT NULL,
                FOREIGN KEY (employee_id) REFERENCES employees(id),
                UNIQUE KEY (employee_id, date)
            )"""
        ]
        
        for query in queries:
            try:
                self.cursor.execute(query)
            except Error as e:
                self.log_message(f"خطأ في إنشاء الجداول: {str(e)}")

    def setup_ui(self):
        """إعداد واجهة المستخدم الرئيسية"""
        # إطار العنوان
        title_frame = tk.Frame(self.root, bg="#2c3e50")
        title_frame.pack(fill=tk.X)
        
        tk.Label(title_frame, 
                text="نظام تسجيل الحضور والانصراف بالبصمة", 
                font=("Arial", 16, "bold"), 
                bg="#2c3e50", fg="white", 
                padx=10, pady=10).pack()
        
        # إطار التحكم
        control_frame = tk.Frame(self.root, padx=10, pady=10)
        control_frame.pack(fill=tk.X)
        
        # أزرار التحكم
        buttons = [
            ("إدارة الموظفين", self.open_employee_management, "#9b59b6"),
            ("الاتصال بجهاز البصمة", self.connect_device, "#27ae60"),
            ("قطع الاتصال", self.disconnect_device, "#e74c3c", tk.DISABLED),
            ("إدخال يدوي", self.open_manual_entry, "#3498db"),
            ("تصدير إلى Excel", self.open_export_dialog, "#3498db"),
            ("التقارير", self.open_report_dialog, "#2ecc71")
        ]
        
        for i, (text, command, color, *state) in enumerate(buttons):
            btn = tk.Button(control_frame, text=text, command=command, 
                           bg=color, fg="white", width=15)
            btn.pack(side=tk.LEFT, padx=5)
            setattr(self, f"btn_{i}", btn)
            if state:
                btn.config(state=state[0])
        
        # إطار عرض البيانات
        data_frame = tk.Frame(self.root)
        data_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # سجل الأحداث
        self.setup_log_frame(data_frame)
        
        # سجل الحضور
        self.setup_attendance_frame(data_frame)
        
        # شريط الحالة
        self.status_bar = tk.Label(self.root, text="جاهز", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def setup_log_frame(self, parent):
        """إعداد إطار سجل الأحداث"""
        log_frame = tk.LabelFrame(parent, text="سجل الأحداث", font=("Arial", 12))
        log_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        self.log_text = tk.Text(log_frame, height=15, state=tk.DISABLED, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = tk.Scrollbar(self.log_text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.log_text.yview)

    def setup_attendance_frame(self, parent):
        """إعداد إطار سجل الحضور"""
        attendance_frame = tk.LabelFrame(parent, text="سجل الحضور", font=("Arial", 12))
        attendance_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        
        columns = ("id", "employee", "date", "time", "status")
        self.attendance_tree = ttk.Treeview(
            attendance_frame, 
            columns=columns, 
            show="headings", 
            height=15,
            selectmode="browse"
        )
        
        # تعريف العناوين والأعمدة
        headings = [
            ("id", "المعرف", 50),
            ("employee", "الموظف", 150),
            ("date", "التاريخ", 100),
            ("time", "الوقت", 100),
            ("status", "الحالة", 80)
        ]
        
        for col, text, width in headings:
            self.attendance_tree.heading(col, text=text)
            self.attendance_tree.column(col, width=width, anchor=tk.CENTER)
        
        self.attendance_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # إضافة شريط تمرير
        scrollbar = ttk.Scrollbar(attendance_frame, orient=tk.VERTICAL, command=self.attendance_tree.yview)
        self.attendance_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # إضافة قائمة السياق (زر يمين الماوس)
        self.attendance_tree.bind("<Button-3>", self.show_attendance_context_menu)

    def show_attendance_context_menu(self, event):
        """عرض قائمة السياق لسجل الحضور"""
        item = self.attendance_tree.identify_row(event.y)
        if not item:
            return
            
        self.attendance_tree.selection_set(item)
        record_id = self.attendance_tree.item(item, "values")[0]
        
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="حذف التسجيل", command=lambda: self.delete_attendance_record(record_id))
        menu.add_command(label="تعديل الحالة", command=lambda: self.edit_attendance_status(record_id))
        menu.tk_popup(event.x_root, event.y_root)

    def delete_attendance_record(self, record_id):
        """حذف تسجيل حضور محدد"""
        if messagebox.askyesno("تأكيد الحذف", "هل أنت متأكد من حذف هذا التسجيل؟"):
            try:
                self.cursor.execute("DELETE FROM attendance WHERE id = %s", (record_id,))
                self.log_message(f"تم حذف تسجيل الحضور رقم {record_id}")
                self.refresh_attendance()
            except Error as e:
                self.log_message(f"خطأ في حذف التسجيل: {str(e)}")

    def edit_attendance_status(self, record_id):
        """تعديل حالة تسجيل حضور"""
        try:
            self.cursor.execute("SELECT status FROM attendance WHERE id = %s", (record_id,))
            current_status = self.cursor.fetchone()["status"]
            new_status = "OUT" if current_status == "IN" else "IN"
            
            self.cursor.execute(
                "UPDATE attendance SET status = %s WHERE id = %s", 
                (new_status, record_id)
            )
            self.log_message(f"تم تغيير حالة التسجيل {record_id} إلى {new_status}")
            self.refresh_attendance()
        except Error as e:
            self.log_message(f"خطأ في تعديل الحالة: {str(e)}")

    def open_employee_management(self):
        """فتح نافذة إدارة الموظفين"""
        management_window = tk.Toplevel(self.root)
        management_window.title("إدارة الموظفين")
        management_window.geometry("1000x600")
        management_window.transient(self.root)
        management_window.grab_set()
        
        # شجرة عرض الموظفين
        self.setup_employee_tree(management_window)
        
        # أزرار التحكم
        control_frame = tk.Frame(management_window)
        control_frame.pack(pady=10)
        
        buttons = [
            ("إضافة موظف", self.open_employee_form, "#27ae60"),
            ("تعديل بيانات", self.edit_employee, "#f1c40f"),
            ("حذف موظف", self.delete_employee, "#e74c3c"),
            ("تحديث القائمة", self.refresh_employee_list, "#3498db")
        ]
        
        for text, command, color in buttons:
            tk.Button(
                control_frame, 
                text=text, 
                command=command, 
                bg=color, 
                fg="white" if color != "#f1c40f" else "black",
                width=15
            ).pack(side=tk.LEFT, padx=5)
        
        self.refresh_employee_list()

    def setup_employee_tree(self, parent):
        """إعداد شجرة عرض الموظفين"""
        columns = [
            ("id", "ID", 50),
            ("fingerprint_id", "رقم البصمة", 100),
            ("name", "الاسم", 200),
            ("department", "القسم", 150),
            ("position", "الوظيفة", 150),
            ("created_at", "تاريخ الإضافة", 120)
        ]
        
        self.employee_tree = ttk.Treeview(
            parent,
            columns=[col[0] for col in columns],
            show="headings",
            height=20,
            selectmode="browse"
        )
        
        for col, text, width in columns:
            self.employee_tree.heading(col, text=text)
            self.employee_tree.column(col, width=width, anchor=tk.CENTER)
        
        self.employee_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # إضافة شريط تمرير
        scrollbar = ttk.Scrollbar(self.employee_tree, orient=tk.VERTICAL, command=self.employee_tree.yview)
        self.employee_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ربط حدث النقر المزدوج
        self.employee_tree.bind("<Double-1>", lambda e: self.edit_employee())

    def open_employee_form(self, employee_data=None):
        """فتح نموذج إضافة/تعديل موظف"""
        form_window = tk.Toplevel(self.root)
        form_window.title("نموذج الموظف" if not employee_data else "تعديل بيانات الموظف")
        form_window.geometry("500x400")
        form_window.transient(self.root)
        form_window.grab_set()
        
        # متغيرات النموذج
        self.emp_form_vars = {
            "id": tk.StringVar(value=employee_data["id"] if employee_data else ""),
            "fingerprint_id": tk.StringVar(value=employee_data["fingerprint_id"] if employee_data else ""),
            "name": tk.StringVar(value=employee_data["name"] if employee_data else ""),
            "department": tk.StringVar(value=employee_data["department"] if employee_data else ""),
            "position": tk.StringVar(value=employee_data["position"] if employee_data else "")
        }
        
        # حقول النموذج
        fields = [
            ("رقم البصمة", "fingerprint_id", True),
            ("الاسم الكامل", "name", True),
            ("القسم", "department", False),
            ("الوظيفة", "position", False)
        ]
        
        for i, (label, var_name, required) in enumerate(fields):
            tk.Label(form_window, text=f"{label}:" + (" *" if required else "")).grid(
                row=i, column=0, padx=10, pady=5, sticky=tk.E
            )
            entry = tk.Entry(form_window, textvariable=self.emp_form_vars[var_name])
            entry.grid(row=i, column=1, padx=10, pady=5, sticky=tk.W)
        
        # زر الحفظ
        save_btn = tk.Button(
            form_window, 
            text="حفظ", 
            command=lambda: self.save_employee(employee_data is not None),
            bg="#27ae60", fg="white", width=15
        )
        save_btn.grid(row=len(fields)+1, column=0, columnspan=2, pady=20)

    def save_employee(self, is_edit=False):
        """حفظ بيانات الموظف في قاعدة البيانات"""
        try:
            data = {k: v.get().strip() for k, v in self.emp_form_vars.items() if k != "id"}
            
            # التحقق من الحقول المطلوبة
            if not data["fingerprint_id"].isdigit():
                messagebox.showerror("خطأ", "رقم البصمة يجب أن يكون رقماً صحيحاً")
                return
                
            if not data["name"]:
                messagebox.showerror("خطأ", "حقل الاسم مطلوب")
                return
            
            if is_edit:
                # عملية التعديل
                self.cursor.execute(
                    """UPDATE employees SET 
                    fingerprint_id = %s, 
                    name = %s, 
                    department = %s, 
                    position = %s 
                    WHERE id = %s""",
                    (data["fingerprint_id"], data["name"], data["department"], 
                     data["position"], self.emp_form_vars["id"].get())
                )
                message = "تم تحديث بيانات الموظف بنجاح"
            else:
                # عملية الإضافة
                self.cursor.execute(
                    """INSERT INTO employees 
                    (fingerprint_id, name, department, position) 
                    VALUES (%s, %s, %s, %s)""",
                    (data["fingerprint_id"], data["name"], data["department"], data["position"])
                )
                message = "تم إضافة الموظف بنجاح"
            
            self.log_message(message)
            self.refresh_employee_list()
            
            # إغلاق النافذة بعد الحفظ
            for window in self.root.winfo_children():
                if isinstance(window, tk.Toplevel) and window.title().startswith("نموذج الموظف"):
                    window.destroy()
                    break
                    
        except mysql.connector.IntegrityError as e:
            if "fingerprint_id" in str(e):
                messagebox.showerror("خطأ", "رقم البصمة مسجل مسبقاً")
            else:
                messagebox.showerror("خطأ", f"فشل في حفظ البيانات: {str(e)}")
        except Error as e:
            messagebox.showerror("خطأ", f"فشل في حفظ البيانات: {str(e)}")

    def edit_employee(self):
        """تعديل بيانات موظف محدد"""
        selected = self.employee_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار موظف للتعديل")
            return
            
        item = self.employee_tree.item(selected[0])
        employee_data = {
            "id": item["values"][0],
            "fingerprint_id": item["values"][1],
            "name": item["values"][2],
            "department": item["values"][3],
            "position": item["values"][4]
        }
        
        self.open_employee_form(employee_data)

    def delete_employee(self):
        """حذف موظف محدد"""
        selected = self.employee_tree.selection()
        if not selected:
            messagebox.showwarning("تحذير", "يرجى اختيار موظف للحذف")
            return
            
        item = self.employee_tree.item(selected[0])
        emp_id, emp_name = item["values"][0], item["values"][2]
        
        if messagebox.askyesno(
            "تأكيد الحذف", 
            f"هل أنت متأكد من حذف الموظف:\n{emp_name}؟\nسيتم حذف جميع سجلات الحضور المرتبطة به"
        ):
            try:
                # حذف سجلات الحضور أولاً بسبب القيود المرجعية
                self.cursor.execute("DELETE FROM attendance WHERE employee_id = %s", (emp_id,))
                self.cursor.execute("DELETE FROM employees WHERE id = %s", (emp_id,))
                
                self.log_message(f"تم حذف الموظف {emp_name} وجميع سجلاته")
                self.refresh_employee_list()
                self.refresh_attendance()
            except Error as e:
                messagebox.showerror("خطأ", f"فشل في حذف الموظف: {str(e)}")

    def refresh_employee_list(self):
        """تحديث قائمة الموظفين"""
        try:
            self.employee_tree.delete(*self.employee_tree.get_children())
            self.cursor.execute("""
                SELECT id, fingerprint_id, name, department, position, 
                DATE_FORMAT(created_at, '%Y-%m-%d') as created_at 
                FROM employees ORDER BY name
            """)
            
            for row in self.cursor.fetchall():
                self.employee_tree.insert("", tk.END, values=(
                    row["id"],
                    row["fingerprint_id"],
                    row["name"],
                    row["department"] or "-",
                    row["position"] or "-",
                    row["created_at"]
                ))
        except Error as e:
            messagebox.showerror("خطأ", f"فشل في تحميل قائمة الموظفين: {str(e)}")

    def calculate_work_hours(self, employee_id, date):
        """حساب ساعات العمل لموظف في تاريخ محدد"""
        try:
            self.cursor.execute("""
                SELECT time, status 
                FROM attendance 
                WHERE employee_id = %s AND date = %s
                ORDER BY time
            """, (employee_id, date))
            print("📌 employee_id before insertion:", employee_id)

            
            records = self.cursor.fetchall()
            total = 0.0
            in_time = None
            
            for record in records:
                time = datetime.strptime(str(record["time"]), "%H:%M:%S").time()
                if record["status"] == "IN":
                    in_time = datetime.combine(date, time)
                elif record["status"] == "OUT" and in_time:
                    out_time = datetime.combine(date, time)
                    delta = out_time - in_time
                    total += delta.total_seconds() / 3600  # تحويل إلى ساعات
                    in_time = None
            
            # حفظ النتائج في جدول work_hours
            self.cursor.execute("""
                INSERT INTO work_hours (employee_id, date, total_hours)
                VALUES (%s, %s, %s)
                ON DUPLICATE KEY UPDATE total_hours = VALUES(total_hours)
            """, (employee_id, date, total))
            print("📌 employee_id before insertion:1", employee_id)

            return total
        except Error as e:
            self.log_message(f"خطأ في حساب ساعات العمل: {str(e)}")
            return 0

    def check_absences(self):
        """فحص الموظفين الغائبين اليوم"""
        try:
            today = datetime.now().date()
            self.cursor.execute("""
                SELECT e.id, e.name 
                FROM employees e
                WHERE NOT EXISTS (
                    SELECT 1 FROM attendance 
                    WHERE employee_id = e.id AND date = %s AND status = 'IN'
                )
            """, (today,))
            
            absentees = self.cursor.fetchall()
            if absentees:
                message = "الموظفون الغائبون اليوم:\n"
                for emp in absentees:
                    message += f"- {emp['name']}\n"
                self.log_message(message)
        except Error as e:
            self.log_message(f"خطأ في فحص الغياب: {str(e)}")

    def open_report_dialog(self):
        """فتح نافذة التقارير"""
        report_window = tk.Toplevel(self.root)
        report_window.title("تقارير الحضور")
        report_window.geometry("800x600")
        report_window.transient(self.root)
        report_window.grab_set()
        
        # عناصر التحكم
        tk.Label(report_window, text="الموظف:", font=("Arial", 12)).pack(pady=5)
        
        # قائمة الموظفين
        self.cursor.execute("SELECT id, name FROM employees ORDER BY name")
        employees = self.cursor.fetchall()
        employee_names = [f"{emp['id']} - {emp['name']}" for emp in employees]
        
        employee_combobox = ttk.Combobox(report_window, values=employee_names, state="readonly")
        employee_combobox.pack(pady=5)
        
        tk.Label(report_window, text="الفترة:", font=("Arial", 12)).pack(pady=5)
        period_combobox = ttk.Combobox(
            report_window, 
            values=["اليوم", "أمس", "الأسبوع الحالي", "الشهر الحالي", "فترة مخصصة"], 
            state="readonly"
        )
        period_combobox.pack(pady=5)
        period_combobox.current(0)
        
        # إطار الفترة المخصصة (يظهر فقط عند اختيار "فترة مخصصة")
        custom_frame = tk.Frame(report_window)
        
        tk.Label(custom_frame, text="من:").grid(row=0, column=0, padx=5)
        self.from_date_entry = tk.Entry(custom_frame, width=12)
        self.from_date_entry.grid(row=0, column=1, padx=5)
        
        tk.Label(custom_frame, text="إلى:").grid(row=0, column=2, padx=5)
        self.to_date_entry = tk.Entry(custom_frame, width=12)
        self.to_date_entry.grid(row=0, column=3, padx=5)
        
        def period_changed(event):
            if period_combobox.get() == "فترة مخصصة":
                custom_frame.pack(pady=10)
            else:
                custom_frame.pack_forget()
        
        period_combobox.bind("<<ComboboxSelected>>", period_changed)
        
        # عرض النتائج
        result_frame = tk.Frame(report_window)
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        columns = ("date", "check_in", "check_out", "hours")
        self.report_tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=10)
        
        headings = [
            ("date", "التاريخ", 100),
            ("check_in", "وقت الحضور", 100),
            ("check_out", "وقت الانصراف", 100),
            ("hours", "الساعات", 80)
        ]
        
        for col, text, width in headings:
            self.report_tree.heading(col, text=text)
            self.report_tree.column(col, width=width, anchor=tk.CENTER)
        
        self.report_tree.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(self.report_tree, orient=tk.VERTICAL, command=self.report_tree.yview)
        self.report_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # زر إنشاء التقرير
        def generate_report():
            employee_id = employee_combobox.get().split(" - ")[0] if employee_combobox.get() else None
            period = period_combobox.get()
            print("📌 employee_id before insertion:2", employee_id)

            if not employee_id:
                messagebox.showwarning("تحذير", "يرجى اختيار موظف")
                return
                
            try:
                # تحديد الفترة الزمنية
                today = datetime.now().date()
                
                if period == "اليوم":
                    start_date = end_date = today
                elif period == "أمس":
                    start_date = end_date = today - timedelta(days=1)
                elif period == "الأسبوع الحالي":
                    start_date = today - timedelta(days=today.weekday())
                    end_date = today
                elif period == "الشهر الحالي":
                    start_date = today.replace(day=1)
                    end_date = today
                else:  # فترة مخصصة
                    try:
                        start_date = datetime.strptime(self.from_date_entry.get(), "%Y-%m-%d").date()
                        end_date = datetime.strptime(self.to_date_entry.get(), "%Y-%m-%d").date()
                    except ValueError:
                        messagebox.showerror("خطأ", "صيغة التاريخ غير صحيحة. استخدم YYYY-MM-DD")
                        return
                
                # استعلام قاعدة البيانات
                query = """
                    SELECT 
                        a.date,
                        MIN(CASE WHEN a.status = 'IN' THEN a.time END) as check_in,
                        MAX(CASE WHEN a.status = 'OUT' THEN a.time END) as check_out,
                        COALESCE(wh.total_hours, 0) as hours
                    FROM 
                        attendance a
                    LEFT JOIN 
                        work_hours wh ON a.employee_id = wh.employee_id AND a.date = wh.date
                    WHERE 
                        a.employee_id = %s AND a.date BETWEEN %s AND %s
                    GROUP BY 
                        a.date, wh.total_hours
                    ORDER BY 
                        a.date
                """
                
                self.cursor.execute(query, (employee_id, start_date, end_date))
                results = self.cursor.fetchall()
                print("📌 employee_id before insertion:", employee_id)

                # عرض النتائج
                self.report_tree.delete(*self.report_tree.get_children())
                
                total_hours = 0
                for row in results:
                    self.report_tree.insert("", tk.END, values=(
                        row["date"],
                        row["check_in"] or "-",
                        row["check_out"] or "-",
                        f"{row['hours']:.2f}" if row["hours"] else "-"
                    ))
                    total_hours += row["hours"] or 0
                
                # عرض إجمالي الساعات
                tk.Label(
                    report_window, 
                    text=f"إجمالي ساعات العمل: {total_hours:.2f} ساعة",
                    font=("Arial", 12, "bold")
                ).pack(pady=5)
                
            except Error as e:
                messagebox.showerror("خطأ", f"فشل في إنشاء التقرير: {str(e)}")
        
        tk.Button(
            report_window, 
            text="إنشاء التقرير", 
            command=generate_report,
            bg="#27ae60", fg="white", width=15
        ).pack(pady=10)

    def get_last_status(self, employee_id):
        """الحصول على آخر حالة حضور للموظف"""
        try:
            self.cursor.execute("""
                SELECT status FROM attendance 
                WHERE employee_id = %s 
                ORDER BY date DESC, time DESC 
                LIMIT 1
            """, (employee_id,))

            result = self.cursor.fetchone()
            
            return result["status"] if result else None
        except Error as e:
            self.log_message(f"خطأ في التحقق من آخر حالة: {str(e)}")
            return None


    def log_message(self, message):
        """إضافة رسالة إلى سجل الأحداث"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(datetime.now())
        print(timestamp)
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.config(state=tk.DISABLED)
        self.log_text.see(tk.END)
        self.status_bar.config(text=message)

    def refresh_attendance(self):
        """تحديث سجل الحضور"""
        try:
            self.attendance_tree.delete(*self.attendance_tree.get_children())
            self.cursor.execute("""
                SELECT a.id, e.name, a.date, a.time, a.status 
                FROM attendance a
                JOIN employees e ON a.employee_id = e.id
                ORDER BY a.date DESC, a.time DESC
                LIMIT 200
            """)
            
            for row in self.cursor.fetchall():
                time_value = row["time"]

                if isinstance(time_value, timedelta):
                    total_seconds = int(time_value.total_seconds())
                    hours = total_seconds // 3600
                    minutes = (total_seconds % 3600) // 60
                    seconds = total_seconds % 60
                    formatted_time = f"{hours:02}:{minutes:02}:{seconds:02}"
                elif isinstance(time_value, time):
                    formatted_time = time_value.strftime("%H:%M:%S")
                else:
                    formatted_time = "-"

                self.attendance_tree.insert("", tk.END, values=(
                    row["id"],
                    row["name"],
                    row["date"],
                    formatted_time,
                    row["status"]
                ))

        except Error as e:
            messagebox.showerror("خطأ", f"فشل في تحديث سجل الحضور: {str(e)}")

    def connect_device(self):
        """الاتصال بجهاز البصمة"""
        try:
            self.client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.client_socket.connect(('127.0.0.1', 9999))
            
            self.running = True
            self.btn_1.config(state=tk.DISABLED)  # تعطيل زر الاتصال
            self.btn_2.config(state=tk.NORMAL)    # تمكين زر قطع الاتصال
            
            self.log_message("تم الاتصال بجهاز البصمة على 127.0.0.1:9999")
            
            # تشغيل thread لاستقبال البيانات
            self.server_thread = threading.Thread(target=self.receive_data, daemon=True)
            self.server_thread.start()
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل في الاتصال بجهاز البصمة: {str(e)}")

    def receive_data(self):
        """استقبال البيانات من جهاز البصمة"""
        while self.running:
            try:
                data = self.client_socket.recv(1024)
                if not data:
                    break
                    
                message = data.decode('utf-8').strip()
                self.log_message(f"تم استقبال: {message}")
                print("data",message)
                
                try:
                    fingerprint_id_str, status = message.split(',')

                    fingerprint_id = int(fingerprint_id_str)
                    print("fingerprint_id",fingerprint_id+3)


                except ValueError:
                    self.log_message("تنسيق بيانات غير صالح")
                    continue
                
                self.process_attendance(fingerprint_id, status)
                
            except Exception as e:
                if self.running:
                    self.log_message(f"خطأ في استقبال البيانات: {str(e)}")
                break

    def process_attendance(self, fingerprint_id, status):
        """معالجة بيانات الحضور"""
        try:
            self.cursor.execute("SELECT id, name FROM employees WHERE fingerprint_id = %s", (fingerprint_id,))
            result = self.cursor.fetchone()
            print("📌 employee_id before insertion:8", result)
            if not result:
                messagebox.showwarning("تحذير", f"لا يوجد موظف مسجل برقم البصمة {fingerprint_id}")
                self.log_message(f"لا يوجد موظف مسجل برقم البصمة {fingerprint_id}")
                return
            if result:
                employee_id = int(result['id'])
                employee_name = result['name']
                print("📌 employee_id before insertion:88", employee_id)
                now = datetime.now()
                date = now.date()
                time_ = now.time()
                
                last_status = self.get_last_status(employee_id)

                if status == "AUTO":
                    status = "OUT" if last_status == "IN" else "IN"

                if last_status == status:
                    self.log_message(f"تحذير: الموظف {employee_name} لديه بالفعل حالة {status}")
                    return

                self.cursor.execute("INSERT INTO attendance (employee_id, date, time, status) VALUES (%s, %s, %s, %s)",
                                    (employee_id, date, time_, status))

                if status == "OUT":
                    self.calculate_work_hours(employee_id, date)

                self.log_message(f"تم تسجيل {status} للموظف {employee_name}")
                self.refresh_attendance()
                
            else:
                messagebox.showwarning("تحذير", f"لا يوجد موظف مسجل برقم البصمة {fingerprint_id}")
                self.log_message(f"لا يوجد موظف مسجل برقم البصمة {fingerprint_id}")
                return
                
        except Error as e:
            self.log_message(f"خطأ في تسجيل الحضور: {str(e)}")
    
    
    def disconnect_device(self):
        """قطع الاتصال بجهاز البصمة"""
        self.running = False
        try:
            if self.client_socket:
                self.client_socket.close()
        except:
            pass
            
        self.btn_1.config(state=tk.NORMAL)  # تمكين زر الاتصال
        self.btn_2.config(state=tk.DISABLED)  # تعطيل زر قطع الاتصال
        self.log_message("تم قطع الاتصال بجهاز البصمة")

    def open_manual_entry(self):
        """فتح نافذة الإدخال اليدوي"""
        manual_window = tk.Toplevel(self.root)
        manual_window.title("إدخال بصمة يدوي")
        manual_window.geometry("400x300")
        manual_window.transient(self.root)
        manual_window.grab_set()

        tk.Label(manual_window, text="إدخال بصمة يدوي", font=("Arial", 14, "bold")).pack(pady=10)

        form_frame = tk.Frame(manual_window, padx=10, pady=10)
        form_frame.pack()

        # حقل رقم البصمة
        tk.Label(form_frame, text="رقم البصمة:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
        fingerprint_entry = tk.Entry(form_frame)
        fingerprint_entry.grid(row=0, column=1, padx=5, pady=5)

        # حقل الحالة
        tk.Label(form_frame, text="الحالة:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
        status_combobox = ttk.Combobox(form_frame, values=["AUTO", "IN", "OUT"], state="readonly")
        status_combobox.grid(row=1, column=1, pady=5)
        status_combobox.current(0)  # AUTO كاختيار افتراضي

        # زر الإرسال
        def submit():
            fingerprint_id = fingerprint_entry.get()
            status = status_combobox.get()
            if not fingerprint_id.isdigit():
                messagebox.showerror("خطأ", "يجب أن يكون رقم البصمة رقمًا صحيحًا")
                return
            self.process_attendance(int(fingerprint_id), status)
            manual_window.destroy()

        button_frame = tk.Frame(manual_window)
        button_frame.pack(pady=10)
        tk.Button(
            button_frame, 
            text="إرسال", 
            command=submit, 
            bg="#27ae60", 
            fg="white", 
            width=10
        ).pack(side=tk.LEFT, padx=5)
        tk.Button(
            button_frame, 
            text="إلغاء", 
            command=manual_window.destroy, 
            bg="#e74c3c", 
            fg="white", 
            width=10
        ).pack(side=tk.LEFT, padx=5)


    def open_export_dialog(self):
        """فتح نافذة تصدير البيانات"""
        export_window = tk.Toplevel(self.root)
        export_window.title("تصدير إلى Excel")
        export_window.geometry("500x300")
        export_window.transient(self.root)
        export_window.grab_set()
        
        tk.Label(export_window, text="تصدير سجل الحضور", font=("Arial", 14, "bold")).pack(pady=10)
        
        form_frame = tk.Frame(export_window, padx=10, pady=10)
        form_frame.pack()
        
        # خيارات التصدير
        tk.Label(form_frame, text="نوع التقرير:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
        report_type = ttk.Combobox(form_frame, values=[
            "سجل الحضور اليومي",
            "إجمالي ساعات العمل",
            "تقرير الغياب"
        ], state="readonly")
        report_type.grid(row=0, column=1, padx=5, pady=5)
        report_type.current(0)
        
        tk.Label(form_frame, text="الفترة الزمنية:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
        period_combobox = ttk.Combobox(form_frame, values=[
            "اليوم",
            "الأسبوع الحالي",
            "الشهر الحالي",
            "فترة مخصصة"
        ], state="readonly")
        period_combobox.grid(row=1, column=1, padx=5, pady=5)
        period_combobox.current(0)
        
        # إطار الفترة المخصصة
        custom_frame = tk.Frame(form_frame)
        
        tk.Label(custom_frame, text="من:").grid(row=0, column=0, padx=5)
        self.export_from_entry = tk.Entry(custom_frame, width=12)
        self.export_from_entry.grid(row=0, column=1, padx=5)
        
        tk.Label(custom_frame, text="إلى:").grid(row=0, column=2, padx=5)
        self.export_to_entry = tk.Entry(custom_frame, width=12)
        self.export_to_entry.grid(row=0, column=3, padx=5)
        
        def period_changed(event):
            if period_combobox.get() == "فترة مخصصة":
                custom_frame.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
            else:
                custom_frame.grid_forget()
        
        period_combobox.bind("<<ComboboxSelected>>", period_changed)
        
        # زر التصدير
        def export():
            report = report_type.get()
            period = period_combobox.get()
            
            # تحديد الفترة الزمنية
            today = datetime.now().date()
            
            if period == "اليوم":
                start_date = end_date = today
            elif period == "الأسبوع الحالي":
                start_date = today - timedelta(days=today.weekday())
                end_date = today
            elif period == "الشهر الحالي":
                start_date = today.replace(day=1)
                end_date = today
            else:  # فترة مخصصة
                try:
                    start_date = datetime.strptime(self.export_from_entry.get(), "%Y-%m-%d").date()
                    end_date = datetime.strptime(self.export_to_entry.get(), "%Y-%m-%d").date()
                except ValueError:
                    messagebox.showerror("خطأ", "صيغة التاريخ غير صحيحة. استخدم YYYY-MM-DD")
                    return
            
            try:
                # بناء استعلام حسب نوع التقرير
                if report == "سجل الحضور اليومي":
                    query = """
                        SELECT 
                            e.name AS employee_name,
                            a.date,
                            a.time,
                            a.status
                        FROM 
                            attendance a
                        JOIN 
                            employees e ON a.employee_id = e.id
                        WHERE 
                            a.date BETWEEN %s AND %s
                        ORDER BY 
                            a.date, a.time
                    """
                    columns = ["الموظف", "التاريخ", "الوقت", "الحالة"]
                elif report == "إجمالي ساعات العمل":
                    query = """
                        SELECT 
                            e.name AS employee_name,
                            wh.date,
                            wh.total_hours
                        FROM 
                            work_hours wh
                        JOIN 
                            employees e ON wh.employee_id = e.id
                        WHERE 
                            wh.date BETWEEN %s AND %s
                        ORDER BY 
                            wh.date
                    """
                    columns = ["الموظف", "التاريخ", "ساعات العمل"]
                else:  # تقرير الغياب
                    query = """
                        SELECT 
                            e.name AS employee_name,
                            d.date
                        FROM 
                            (SELECT DISTINCT date FROM attendance 
                            WHERE date BETWEEN %s AND %s) d
                        CROSS JOIN 
                            employees e
                        WHERE 
                            NOT EXISTS (
                                SELECT 1 FROM attendance 
                                WHERE employee_id = e.id AND date = d.date AND status = 'IN'
                            )
                        ORDER BY 
                            d.date, e.name
                    """
                    columns = ["الموظف", "تاريخ الغياب"]
                
                self.cursor.execute(query, (start_date, end_date))
                data = self.cursor.fetchall()
                
                if not data:
                    messagebox.showinfo("معلومات", "لا توجد بيانات للتصدير في الفترة المحددة")
                    return
                
                # تحويل البيانات إلى DataFrame
                df = pd.DataFrame(data, columns=columns)
                
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
                
            except Error as e:
                messagebox.showerror("خطأ", f"فشل في تصدير البيانات: {str(e)}")
        
        button_frame = tk.Frame(export_window)
        button_frame.pack(pady=10)
        
        tk.Button(
            button_frame, 
            text="تصدير", 
            command=export, 
            bg="#27ae60", 
            fg="white", 
            width=10
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            button_frame, 
            text="إلغاء", 
            command=export_window.destroy, 
            bg="#e74c3c", 
            fg="white", 
            width=10
        ).pack(side=tk.LEFT, padx=5)

    def on_close(self):
        """إغلاق التطبيق بشكل آمن"""
        if messagebox.askokcancel("خروج", "هل أنت متأكد من رغبتك في الخروج؟"):
            try:
                if self.running:
                    self.disconnect_device()
                if self.db and self.db.is_connected():
                    self.db.close()
            except:
                pass
            finally:
                self.root.destroy()

    def run(self):
        """تشغيل التطبيق"""
        self.root.mainloop()
if __name__ == "__main__":
    root = tk.Tk()
    app = AttendanceRecorder(root)
    app.run()