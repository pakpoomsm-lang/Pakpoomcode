import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
import sqlite3
import pandas as pd
from tkcalendar import DateEntry
import json

class OTRecordingSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("ระบบบันทึกการทำ OT พนักงาน - Heat Exchange Indoor")
        self.root.geometry("1400x800")

        # Initialize database
        self.init_database()

        # Create main UI
        self.create_widgets()

        # Load employees
        self.load_employees()

    def init_database(self):
        """Initialize SQLite database"""
        self.conn = sqlite3.connect('ot_records.db')
        self.cursor = self.conn.cursor()

        # Create employees table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                emp_id TEXT UNIQUE NOT NULL,
                first_name TEXT NOT NULL,
                last_name TEXT NOT NULL,
                workplace TEXT,
                emp_type TEXT,
                shift TEXT,
                supervisor TEXT,
                gl_number INTEGER
            )
        ''')

        # Create OT records table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS ot_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                work_date DATE NOT NULL,
                emp_id TEXT NOT NULL,
                workplace TEXT,
                work_description TEXT,
                remark TEXT,
                status TEXT DEFAULT 'NT',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (emp_id) REFERENCES employees(emp_id)
            )
        ''')

        self.conn.commit()

    def create_widgets(self):
        """Create main UI widgets"""
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # Tab 1: Record OT
        self.tab_record = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_record, text='บันทึกการทำ OT')
        self.create_record_tab()

        # Tab 2: View/Print Reports
        self.tab_report = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_report, text='รายงาน OT')
        self.create_report_tab()

        # Tab 3: Import/Manage Employees
        self.tab_employee = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_employee, text='จัดการพนักงาน')
        self.create_employee_tab()

    def create_record_tab(self):
        """Create OT recording tab"""
        # Top frame for filters
        filter_frame = ttk.LabelFrame(self.tab_record, text="ข้อมูลการทำงาน", padding=10)
        filter_frame.pack(fill='x', padx=10, pady=10)

        # Work date
        ttk.Label(filter_frame, text="วันที่ทำงาน:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.work_date = DateEntry(filter_frame, width=20, date_pattern='dd/mm/yyyy')
        self.work_date.grid(row=0, column=1, sticky='w', padx=5, pady=5)

        # GL Selection
        ttk.Label(filter_frame, text="GL:").grid(row=0, column=2, sticky='w', padx=5, pady=5)
        self.gl_combo = ttk.Combobox(filter_frame, width=25, state='readonly')
        self.gl_combo.grid(row=0, column=3, sticky='w', padx=5, pady=5)
        self.gl_combo.bind('<<ComboboxSelected>>', self.on_gl_selected)

        # Process Selection
        ttk.Label(filter_frame, text="Process:").grid(row=0, column=4, sticky='w', padx=5, pady=5)
        self.process_combo = ttk.Combobox(filter_frame, width=25, state='readonly')
        self.process_combo.grid(row=0, column=5, sticky='w', padx=5, pady=5)
        self.process_combo.bind('<<ComboboxSelected>>', self.on_process_selected)

        # Shift Selection
        ttk.Label(filter_frame, text="Shift:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.shift_combo = ttk.Combobox(filter_frame, values=['All', 'A', 'B', 'D'], width=20, state='readonly')
        self.shift_combo.set('A')
        self.shift_combo.grid(row=1, column=1, sticky='w', padx=5, pady=5)
        self.shift_combo.bind('<<ComboboxSelected>>', self.filter_employees)

        # Employee list frame
        emp_frame = ttk.LabelFrame(self.tab_record, text="รายชื่อพนักงาน", padding=10)
        emp_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Employee treeview
        columns = ('emp_id', 'name', 'workplace', 'type', 'shift', 'supervisor')
        self.emp_tree = ttk.Treeview(emp_frame, columns=columns, show='headings', height=10)

        self.emp_tree.heading('emp_id', text='รหัสพนักงาน')
        self.emp_tree.heading('name', text='ชื่อ-นามสกุล')
        self.emp_tree.heading('workplace', text='สถานที่ทำงาน')
        self.emp_tree.heading('type', text='ประเภท')
        self.emp_tree.heading('shift', text='Shift')
        self.emp_tree.heading('supervisor', text='หัวหน้างาน')

        self.emp_tree.column('emp_id', width=100)
        self.emp_tree.column('name', width=200)
        self.emp_tree.column('workplace', width=250)
        self.emp_tree.column('type', width=80)
        self.emp_tree.column('shift', width=60)
        self.emp_tree.column('supervisor', width=200)

        # Scrollbar for employee tree
        emp_scrollbar = ttk.Scrollbar(emp_frame, orient='vertical', command=self.emp_tree.yview)
        self.emp_tree.configure(yscrollcommand=emp_scrollbar.set)

        self.emp_tree.pack(side='left', fill='both', expand=True)
        emp_scrollbar.pack(side='right', fill='y')

        # Bind double-click to add work entry
        self.emp_tree.bind('<Double-1>', self.add_work_entry)

        # Work entries frame
        work_frame = ttk.LabelFrame(self.tab_record, text="รายการงานที่บันทึก", padding=10)
        work_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Work entries treeview
        work_columns = ('emp_id', 'name', 'workplace', 'work_desc', 'remark', 'status')
        self.work_tree = ttk.Treeview(work_frame, columns=work_columns, show='headings', height=8)

        self.work_tree.heading('emp_id', text='รหัส')
        self.work_tree.heading('name', text='ชื่อ-นามสกุล')
        self.work_tree.heading('workplace', text='สถานที่ทำงาน')
        self.work_tree.heading('work_desc', text='รายการงาน')
        self.work_tree.heading('remark', text='Remark')
        self.work_tree.heading('status', text='สถานะ')

        self.work_tree.column('emp_id', width=80)
        self.work_tree.column('name', width=180)
        self.work_tree.column('workplace', width=200)
        self.work_tree.column('work_desc', width=300)
        self.work_tree.column('remark', width=200)
        self.work_tree.column('status', width=80)

        work_scrollbar = ttk.Scrollbar(work_frame, orient='vertical', command=self.work_tree.yview)
        self.work_tree.configure(yscrollcommand=work_scrollbar.set)

        self.work_tree.pack(side='left', fill='both', expand=True)
        work_scrollbar.pack(side='right', fill='y')

        # Buttons frame
        btn_frame = ttk.Frame(self.tab_record)
        btn_frame.pack(fill='x', padx=10, pady=10)

        ttk.Button(btn_frame, text="ลบรายการ", command=self.delete_work_entry).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Preview", command=self.preview_work).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="บันทึกข้อมูล", command=self.save_work_entries).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="ล้างข้อมูล", command=self.clear_work_entries).pack(side='left', padx=5)

    def create_report_tab(self):
        """Create report viewing tab"""
        # Filter frame
        filter_frame = ttk.LabelFrame(self.tab_report, text="กรองข้อมูล", padding=10)
        filter_frame.pack(fill='x', padx=10, pady=10)

        # Date range
        ttk.Label(filter_frame, text="จากวันที่:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.report_from_date = DateEntry(filter_frame, width=15, date_pattern='dd/mm/yyyy')
        self.report_from_date.grid(row=0, column=1, sticky='w', padx=5, pady=5)

        ttk.Label(filter_frame, text="ถึงวันที่:").grid(row=0, column=2, sticky='w', padx=5, pady=5)
        self.report_to_date = DateEntry(filter_frame, width=15, date_pattern='dd/mm/yyyy')
        self.report_to_date.grid(row=0, column=3, sticky='w', padx=5, pady=5)

        # Process filter
        ttk.Label(filter_frame, text="Process:").grid(row=0, column=4, sticky='w', padx=5, pady=5)
        self.report_process_combo = ttk.Combobox(filter_frame, values=['All'], width=20, state='readonly')
        self.report_process_combo.set('All')
        self.report_process_combo.grid(row=0, column=5, sticky='w', padx=5, pady=5)

        ttk.Button(filter_frame, text="ค้นหา", command=self.search_reports).grid(row=0, column=6, padx=5, pady=5)
        ttk.Button(filter_frame, text="Export Excel", command=self.export_to_excel).grid(row=0, column=7, padx=5, pady=5)

        # Report treeview
        report_frame = ttk.Frame(self.tab_report)
        report_frame.pack(fill='both', expand=True, padx=10, pady=10)

        report_columns = ('date', 'emp_id', 'name', 'workplace', 'work_desc', 'remark', 'status')
        self.report_tree = ttk.Treeview(report_frame, columns=report_columns, show='headings')

        self.report_tree.heading('date', text='วันที่')
        self.report_tree.heading('emp_id', text='รหัสพนักงาน')
        self.report_tree.heading('name', text='ชื่อ-นามสกุล')
        self.report_tree.heading('workplace', text='สถานที่ทำงาน')
        self.report_tree.heading('work_desc', text='รายการงาน')
        self.report_tree.heading('remark', text='Remark')
        self.report_tree.heading('status', text='สถานะ')

        self.report_tree.column('date', width=100)
        self.report_tree.column('emp_id', width=100)
        self.report_tree.column('name', width=180)
        self.report_tree.column('workplace', width=200)
        self.report_tree.column('work_desc', width=300)
        self.report_tree.column('remark', width=200)
        self.report_tree.column('status', width=80)

        report_scrollbar = ttk.Scrollbar(report_frame, orient='vertical', command=self.report_tree.yview)
        self.report_tree.configure(yscrollcommand=report_scrollbar.set)

        self.report_tree.pack(side='left', fill='both', expand=True)
        report_scrollbar.pack(side='right', fill='y')

    def create_employee_tab(self):
        """Create employee management tab"""
        # Import frame
        import_frame = ttk.LabelFrame(self.tab_employee, text="นำเข้าข้อมูลพนักงาน", padding=10)
        import_frame.pack(fill='x', padx=10, pady=10)

        ttk.Label(import_frame, text="นำเข้าข้อมูลพนักงานจากไฟล์ Excel:").pack(side='left', padx=5)
        ttk.Button(import_frame, text="เลือกไฟล์ Excel", command=self.import_employees).pack(side='left', padx=5)
        ttk.Button(import_frame, text="ลบข้อมูลพนักงานทั้งหมด", command=self.clear_all_employees).pack(side='left', padx=5)

        # Employee list frame
        emp_list_frame = ttk.LabelFrame(self.tab_employee, text="รายชื่อพนักงานในระบบ", padding=10)
        emp_list_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Employee treeview
        emp_columns = ('emp_id', 'name', 'workplace', 'type', 'shift', 'supervisor', 'gl')
        self.emp_mgmt_tree = ttk.Treeview(emp_list_frame, columns=emp_columns, show='headings')

        self.emp_mgmt_tree.heading('emp_id', text='รหัสพนักงาน')
        self.emp_mgmt_tree.heading('name', text='ชื่อ-นามสกุล')
        self.emp_mgmt_tree.heading('workplace', text='สถานที่ทำงาน')
        self.emp_mgmt_tree.heading('type', text='ประเภท')
        self.emp_mgmt_tree.heading('shift', text='Shift')
        self.emp_mgmt_tree.heading('supervisor', text='หัวหน้างาน')
        self.emp_mgmt_tree.heading('gl', text='GL')

        self.emp_mgmt_tree.column('emp_id', width=100)
        self.emp_mgmt_tree.column('name', width=200)
        self.emp_mgmt_tree.column('workplace', width=250)
        self.emp_mgmt_tree.column('type', width=80)
        self.emp_mgmt_tree.column('shift', width=60)
        self.emp_mgmt_tree.column('supervisor', width=180)
        self.emp_mgmt_tree.column('gl', width=80)

        emp_mgmt_scrollbar = ttk.Scrollbar(emp_list_frame, orient='vertical', command=self.emp_mgmt_tree.yview)
        self.emp_mgmt_tree.configure(yscrollcommand=emp_mgmt_scrollbar.set)

        self.emp_mgmt_tree.pack(side='left', fill='both', expand=True)
        emp_mgmt_scrollbar.pack(side='right', fill='y')

        # Status label
        self.status_label = ttk.Label(self.tab_employee, text="", foreground='blue')
        self.status_label.pack(pady=5)

    def import_employees(self):
        """Import employees from Excel file"""
        file_path = filedialog.askopenfilename(
            title="เลือกไฟล์ Excel รายชื่อพนักงาน",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )

        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)

            # Clear existing employees
            self.cursor.execute('DELETE FROM employees')

            # Import employees
            count = 0
            for _, row in df.iterrows():
                try:
                    emp_id = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ''
                    first_name = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ''
                    last_name = str(row.iloc[3]) if pd.notna(row.iloc[3]) else ''
                    workplace = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ''
                    emp_type = str(row.iloc[4]) if pd.notna(row.iloc[4]) else ''
                    shift = str(row.iloc[5]) if pd.notna(row.iloc[5]) else ''
                    supervisor = str(row.iloc[6]) if pd.notna(row.iloc[6]) else ''
                    gl_number = int(row.iloc[10]) if pd.notna(row.iloc[10]) else 0

                    if emp_id and first_name:
                        self.cursor.execute('''
                            INSERT INTO employees (emp_id, first_name, last_name, workplace, emp_type, shift, supervisor, gl_number)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (emp_id, first_name, last_name, workplace, emp_type, shift, supervisor, gl_number))
                        count += 1
                except Exception as e:
                    continue

            self.conn.commit()
            self.status_label.config(text=f"นำเข้าข้อมูลพนักงานสำเร็จ จำนวน {count} คน")

            # Reload employees
            self.load_employees()
            self.load_employee_management_list()

            messagebox.showinfo("สำเร็จ", f"นำเข้าข้อมูลพนักงาน {count} คน")

        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถนำเข้าข้อมูลได้: {str(e)}")

    def load_employees(self):
        """Load employees into combo boxes"""
        # Load GLs
        self.cursor.execute('SELECT DISTINCT supervisor FROM employees WHERE supervisor != "" ORDER BY supervisor')
        supervisors = [row[0] for row in self.cursor.fetchall()]
        self.gl_combo['values'] = ['All'] + supervisors
        if supervisors:
            self.gl_combo.set('All')

        # Load Processes (workplaces)
        self.cursor.execute('SELECT DISTINCT workplace FROM employees WHERE workplace != "" ORDER BY workplace')
        workplaces = [row[0] for row in self.cursor.fetchall()]
        self.process_combo['values'] = ['All'] + workplaces
        if workplaces:
            self.process_combo.set('All')

        # Load for report tab
        self.report_process_combo['values'] = ['All'] + workplaces

        # Initial filter
        self.filter_employees()

    def load_employee_management_list(self):
        """Load all employees in management tab"""
        # Clear existing items
        for item in self.emp_mgmt_tree.get_children():
            self.emp_mgmt_tree.delete(item)

        # Load employees
        self.cursor.execute('''
            SELECT emp_id, first_name, last_name, workplace, emp_type, shift, supervisor, gl_number
            FROM employees
            ORDER BY gl_number, emp_id
        ''')

        for row in self.cursor.fetchall():
            emp_id, first, last, workplace, emp_type, shift, supervisor, gl = row
            name = f"{first} {last}"
            self.emp_mgmt_tree.insert('', 'end', values=(emp_id, name, workplace, emp_type, shift, supervisor, gl))

    def on_gl_selected(self, event=None):
        """Handle GL selection"""
        self.filter_employees()

    def on_process_selected(self, event=None):
        """Handle process selection"""
        self.filter_employees()

    def filter_employees(self, event=None):
        """Filter employees based on selections"""
        # Clear existing items
        for item in self.emp_tree.get_children():
            self.emp_tree.delete(item)

        # Build query
        query = '''
            SELECT emp_id, first_name, last_name, workplace, emp_type, shift, supervisor
            FROM employees WHERE 1=1
        '''
        params = []

        # Filter by GL
        gl = self.gl_combo.get()
        if gl and gl != 'All':
            query += ' AND supervisor = ?'
            params.append(gl)

        # Filter by Process
        process = self.process_combo.get()
        if process and process != 'All':
            query += ' AND workplace = ?'
            params.append(process)

        # Filter by Shift
        shift = self.shift_combo.get()
        if shift and shift != 'All':
            query += ' AND shift = ?'
            params.append(shift)

        query += ' ORDER BY emp_id'

        self.cursor.execute(query, params)

        for row in self.cursor.fetchall():
            emp_id, first, last, workplace, emp_type, shift, supervisor = row
            name = f"{first} {last}"
            self.emp_tree.insert('', 'end', values=(emp_id, name, workplace, emp_type, shift, supervisor))

    def add_work_entry(self, event=None):
        """Add work entry for selected employee"""
        selection = self.emp_tree.selection()
        if not selection:
            return

        item = self.emp_tree.item(selection[0])
        emp_id = item['values'][0]
        name = item['values'][1]
        workplace = item['values'][2]

        # Create dialog for work details
        dialog = tk.Toplevel(self.root)
        dialog.title("เพิ่มรายการงาน")
        dialog.geometry("600x300")
        dialog.transient(self.root)
        dialog.grab_set()

        # Work description
        ttk.Label(dialog, text="รายการงาน:").grid(row=0, column=0, sticky='w', padx=10, pady=10)
        work_desc_entry = tk.Text(dialog, width=50, height=5)
        work_desc_entry.grid(row=0, column=1, padx=10, pady=10)

        # Remark
        ttk.Label(dialog, text="Remark:").grid(row=1, column=0, sticky='w', padx=10, pady=10)
        remark_entry = tk.Text(dialog, width=50, height=3)
        remark_entry.grid(row=1, column=1, padx=10, pady=10)

        # Status
        ttk.Label(dialog, text="สถานะ:").grid(row=2, column=0, sticky='w', padx=10, pady=10)
        status_combo = ttk.Combobox(dialog, values=['NT', 'Leave', 'MCP', 'Subcon'], width=20, state='readonly')
        status_combo.set('NT')
        status_combo.grid(row=2, column=1, sticky='w', padx=10, pady=10)

        def save_entry():
            work_desc = work_desc_entry.get('1.0', 'end-1c').strip()
            remark = remark_entry.get('1.0', 'end-1c').strip()
            status = status_combo.get()

            if not work_desc:
                messagebox.showwarning("คำเตือน", "กรุณาระบุรายการงาน")
                return

            # Add to work tree
            self.work_tree.insert('', 'end', values=(emp_id, name, workplace, work_desc, remark, status))
            dialog.destroy()

        ttk.Button(dialog, text="บันทึก", command=save_entry).grid(row=3, column=1, sticky='e', padx=10, pady=10)

    def delete_work_entry(self):
        """Delete selected work entry"""
        selection = self.work_tree.selection()
        if not selection:
            messagebox.showwarning("คำเตือน", "กรุณาเลือกรายการที่ต้องการลบ")
            return

        if messagebox.askyesno("ยืนยัน", "ต้องการลบรายการที่เลือกใช่หรือไม่?"):
            for item in selection:
                self.work_tree.delete(item)

    def preview_work(self):
        """Preview work entries before saving"""
        items = self.work_tree.get_children()
        if not items:
            messagebox.showwarning("คำเตือน", "ไม่มีรายการงานที่จะแสดง")
            return

        # Create preview window
        preview = tk.Toplevel(self.root)
        preview.title("ตรวจสอบข้อมูลก่อนบันทึก")
        preview.geometry("1200x600")

        # Date label
        work_date = self.work_date.get_date()
        ttk.Label(preview, text=f"วันที่ทำงาน: {work_date.strftime('%d/%m/%Y')}",
                 font=('TH Sarabun New', 14, 'bold')).pack(pady=10)

        # Preview treeview
        preview_frame = ttk.Frame(preview)
        preview_frame.pack(fill='both', expand=True, padx=10, pady=10)

        columns = ('emp_id', 'name', 'workplace', 'work_desc', 'remark', 'status')
        preview_tree = ttk.Treeview(preview_frame, columns=columns, show='headings')

        preview_tree.heading('emp_id', text='รหัส')
        preview_tree.heading('name', text='ชื่อ-นามสกุล')
        preview_tree.heading('workplace', text='สถานที่ทำงาน')
        preview_tree.heading('work_desc', text='รายการงาน')
        preview_tree.heading('remark', text='Remark')
        preview_tree.heading('status', text='สถานะ')

        preview_tree.column('emp_id', width=80)
        preview_tree.column('name', width=180)
        preview_tree.column('workplace', width=200)
        preview_tree.column('work_desc', width=350)
        preview_tree.column('remark', width=200)
        preview_tree.column('status', width=80)

        preview_scrollbar = ttk.Scrollbar(preview_frame, orient='vertical', command=preview_tree.yview)
        preview_tree.configure(yscrollcommand=preview_scrollbar.set)

        preview_tree.pack(side='left', fill='both', expand=True)
        preview_scrollbar.pack(side='right', fill='y')

        # Load data
        for item in items:
            values = self.work_tree.item(item)['values']
            preview_tree.insert('', 'end', values=values)

        ttk.Label(preview, text=f"จำนวนรายการทั้งหมด: {len(items)} รายการ",
                 font=('TH Sarabun New', 12)).pack(pady=5)

        ttk.Button(preview, text="ปิด", command=preview.destroy).pack(pady=10)

    def save_work_entries(self):
        """Save work entries to database"""
        items = self.work_tree.get_children()
        if not items:
            messagebox.showwarning("คำเตือน", "ไม่มีรายการงานที่จะบันทึก")
            return

        if not messagebox.askyesno("ยืนยัน", "ต้องการบันทึกข้อมูลทั้งหมดใช่หรือไม่?"):
            return

        work_date = self.work_date.get_date().strftime('%Y-%m-%d')

        try:
            for item in items:
                values = self.work_tree.item(item)['values']
                emp_id, name, workplace, work_desc, remark, status = values

                self.cursor.execute('''
                    INSERT INTO ot_records (work_date, emp_id, workplace, work_description, remark, status)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (work_date, emp_id, workplace, work_desc, remark, status))

            self.conn.commit()
            messagebox.showinfo("สำเร็จ", f"บันทึกข้อมูล {len(items)} รายการเรียบร้อยแล้ว")

            # Clear work entries
            self.clear_work_entries()

        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถบันทึกข้อมูลได้: {str(e)}")

    def clear_work_entries(self):
        """Clear all work entries"""
        for item in self.work_tree.get_children():
            self.work_tree.delete(item)

    def search_reports(self):
        """Search and display reports"""
        # Clear existing items
        for item in self.report_tree.get_children():
            self.report_tree.delete(item)

        from_date = self.report_from_date.get_date().strftime('%Y-%m-%d')
        to_date = self.report_to_date.get_date().strftime('%Y-%m-%d')

        query = '''
            SELECT ot.work_date, ot.emp_id,
                   e.first_name || ' ' || e.last_name as name,
                   ot.workplace, ot.work_description, ot.remark, ot.status
            FROM ot_records ot
            JOIN employees e ON ot.emp_id = e.emp_id
            WHERE ot.work_date BETWEEN ? AND ?
        '''
        params = [from_date, to_date]

        process = self.report_process_combo.get()
        if process and process != 'All':
            query += ' AND ot.workplace = ?'
            params.append(process)

        query += ' ORDER BY ot.work_date, e.gl_number, ot.emp_id'

        self.cursor.execute(query, params)

        for row in self.cursor.fetchall():
            work_date = datetime.strptime(row[0], '%Y-%m-%d').strftime('%d/%m/%Y')
            self.report_tree.insert('', 'end', values=(work_date,) + row[1:])

    def export_to_excel(self):
        """Export report to Excel"""
        items = self.report_tree.get_children()
        if not items:
            messagebox.showwarning("คำเตือน", "ไม่มีข้อมูลที่จะ Export")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"OT_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        if not file_path:
            return

        try:
            data = []
            for item in items:
                data.append(self.report_tree.item(item)['values'])

            df = pd.DataFrame(data, columns=['วันที่', 'รหัสพนักงาน', 'ชื่อ-นามสกุล',
                                            'สถานที่ทำงาน', 'รายการงาน', 'Remark', 'สถานะ'])

            df.to_excel(file_path, index=False, engine='openpyxl')
            messagebox.showinfo("สำเร็จ", f"Export ข้อมูลไปยัง {file_path} เรียบร้อยแล้ว")

        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถ Export ข้อมูลได้: {str(e)}")

    def clear_all_employees(self):
        """Clear all employees from database"""
        if messagebox.askyesno("ยืนยัน", "ต้องการลบข้อมูลพนักงานทั้งหมดใช่หรือไม่?\n(ข้อมูล OT จะยังคงอยู่)"):
            self.cursor.execute('DELETE FROM employees')
            self.conn.commit()

            self.load_employees()
            self.load_employee_management_list()

            messagebox.showinfo("สำเร็จ", "ลบข้อมูลพนักงานทั้งหมดเรียบร้อยแล้ว")

    def __del__(self):
        """Close database connection"""
        if hasattr(self, 'conn'):
            self.conn.close()

if __name__ == '__main__':
    root = tk.Tk()
    app = OTRecordingSystem(root)
    root.mainloop()
