from flask import Flask, render_template, request, jsonify, send_file
import sqlite3
import pandas as pd
from datetime import datetime
import os
import json

OT_DIR   = os.path.dirname(os.path.abspath(__file__))
DB_PATH  = os.path.join(OT_DIR, 'ot_records.db')
STATIC_DIR = os.path.join(OT_DIR, 'static')

app = Flask(__name__)
app.config['SECRET_KEY'] = 'ot-recording-system-2026'

# Database initialization
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Create employees table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id TEXT UNIQUE NOT NULL,
            first_name TEXT NOT NULL,
            last_name TEXT NOT NULL,
            workplace TEXT,
            group_name TEXT,
            emp_type TEXT,
            shift TEXT,
            supervisor TEXT
        )
    ''')

    # Create OT records table
    cursor.execute('''
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

    # Create attendance records table (NEW)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS attendance_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            work_date DATE NOT NULL,
            emp_id TEXT NOT NULL,
            attendance_type TEXT NOT NULL,
            workplace_substitute TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (emp_id) REFERENCES employees(emp_id),
            UNIQUE(work_date, emp_id, attendance_type)
        )
    ''')

    # Add workplace_substitute column if it doesn't exist (for existing databases)
    try:
        cursor.execute("SELECT workplace_substitute FROM attendance_records LIMIT 1")
    except sqlite3.OperationalError:
        cursor.execute("ALTER TABLE attendance_records ADD COLUMN workplace_substitute TEXT")

    try:
        cursor.execute("SELECT leave_period FROM attendance_records LIMIT 1")
    except sqlite3.OperationalError:
        cursor.execute("ALTER TABLE attendance_records ADD COLUMN leave_period TEXT")

    try:
        cursor.execute("SELECT leave_time FROM attendance_records LIMIT 1")
    except sqlite3.OperationalError:
        cursor.execute("ALTER TABLE attendance_records ADD COLUMN leave_time TEXT")

    # Create holidays table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS holidays (
            date TEXT PRIMARY KEY,
            description TEXT,
            type TEXT DEFAULT 'holiday',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Employees excluded from working-hour calculation (permanent until removed)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS hour_exclusions (
            emp_id TEXT PRIMARY KEY,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    # Add type column for existing databases
    try:
        cursor.execute("SELECT type FROM holidays LIMIT 1")
    except sqlite3.OperationalError:
        cursor.execute("ALTER TABLE holidays ADD COLUMN type TEXT DEFAULT 'holiday'")

    # ── Indexes (CREATE IF NOT EXISTS — ปลอดภัย รันซ้ำได้) ──────────────
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_att_date     ON attendance_records(work_date)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_att_emp      ON attendance_records(emp_id)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_att_date_emp ON attendance_records(work_date, emp_id)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_emp_shift    ON employees(shift)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_emp_super    ON employees(supervisor)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_emp_workplace ON employees(workplace)")
    # ────────────────────────────────────────────────────────────────────

    conn.commit()
    conn.close()

init_db()

@app.route('/')
def index():
    """Main page - Record OT"""
    return render_template('index.html')

@app.route('/report')
def report():
    """Report page"""
    return render_template('report.html')

@app.route('/employees')
def employees():
    """Employee management page"""
    return render_template('employees.html')

# API Endpoints

@app.route('/api/supervisors')
def get_supervisors():
    """Get all supervisors/GLs filtered by shift"""
    shift = request.args.get('shift', 'All')
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    if shift and shift != 'All':
        cursor.execute('SELECT DISTINCT supervisor FROM employees WHERE supervisor != "" AND shift = ? ORDER BY supervisor', (shift,))
    else:
        cursor.execute('SELECT DISTINCT supervisor FROM employees WHERE supervisor != "" ORDER BY supervisor')
    supervisors = [row[0] for row in cursor.fetchall()]
    conn.close()
    return jsonify(supervisors)

@app.route('/api/workplaces')
def get_workplaces():
    """Get all workplaces/processes"""
    excluded_lower = {'พนักงาน day', 'พนักงานป่วย', 'พนักงานป่วย day', 'พนักงานท้อง'}
    excluded_prefixes = ('gl', 'sub gl', 'fm')
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('SELECT DISTINCT workplace FROM employees WHERE workplace != "" ORDER BY workplace')
    workplaces = [
        row[0] for row in cursor.fetchall()
        if row[0].strip().lower() not in excluded_lower
        and not row[0].strip().lower().startswith(excluded_prefixes)
    ]
    conn.close()
    return jsonify(workplaces)

@app.route('/api/groups')
def get_groups():
    """Get distinct group names for OT substitute selection"""
    excluded_prefixes = ('gl', 'sub gl', 'fm')
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('SELECT DISTINCT group_name FROM employees WHERE group_name != "" ORDER BY group_name')
    groups = [
        row[0] for row in cursor.fetchall()
        if not row[0].strip().lower().startswith(excluded_prefixes)
    ]
    conn.close()
    return jsonify(groups)

@app.route('/api/employees/filter')
def filter_employees():
    """Filter employees by GL, Process, and Shift"""
    gl = request.args.get('gl', 'All')
    process = request.args.get('process', 'All')
    shift = request.args.get('shift', 'All')

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    query = '''
        SELECT emp_id, first_name, last_name, workplace, group_name, emp_type, shift, supervisor
        FROM employees WHERE 1=1
    '''
    params = []

    if gl and gl != 'All':
        query += ' AND supervisor = ?'
        params.append(gl)

    if process and process != 'All':
        query += ' AND workplace = ?'
        params.append(process)

    if shift and shift != 'All':
        query += ' AND shift = ?'
        params.append(shift)

    query += ' ORDER BY emp_id'

    cursor.execute(query, params)
    employees = []
    for row in cursor.fetchall():
        employees.append({
            'emp_id': row[0],
            'first_name': row[1],
            'last_name': row[2],
            'name': f"{row[1]} {row[2]}",
            'workplace': row[3],
            'group_name': row[4],
            'emp_type': row[5],
            'shift': row[6],
            'supervisor': row[7]
        })

    conn.close()
    return jsonify(employees)

@app.route('/api/ot/save', methods=['POST'])
def save_ot_records():
    """Save OT records"""
    data = request.json
    work_date = data.get('work_date')
    records = data.get('records', [])

    if not work_date or not records:
        return jsonify({'success': False, 'message': 'ข้อมูลไม่ครบถ้วน'}), 400

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    try:
        for record in records:
            # Get workplace from employee data
            cursor.execute('SELECT workplace FROM employees WHERE emp_id = ?', (record['emp_id'],))
            result = cursor.fetchone()
            workplace = result[0] if result else ''

            cursor.execute('''
                INSERT INTO ot_records (work_date, emp_id, workplace, work_description, remark, status)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (work_date, record['emp_id'], workplace,
                  record['work_description'], record['remark'], record['status']))

        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': f'บันทึกข้อมูล {len(records)} รายการเรียบร้อยแล้ว'})

    except Exception as e:
        conn.rollback()
        conn.close()
        return jsonify({'success': False, 'message': f'เกิดข้อผิดพลาด: {str(e)}'}), 500

@app.route('/api/reports/search')
def search_reports():
    """Search OT reports"""
    from_date = request.args.get('from_date')
    to_date = request.args.get('to_date')
    process = request.args.get('process', 'All')

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    query = '''
        SELECT ot.work_date, ot.emp_id,
               e.first_name || ' ' || e.last_name as name,
               ot.workplace, ot.work_description, ot.remark, ot.status
        FROM ot_records ot
        JOIN employees e ON ot.emp_id = e.emp_id
        WHERE ot.work_date BETWEEN ? AND ?
    '''
    params = [from_date, to_date]

    if process and process != 'All':
        query += ' AND ot.workplace = ?'
        params.append(process)

    query += ' ORDER BY ot.work_date, e.gl_number, ot.emp_id'

    cursor.execute(query, params)

    reports = []
    for row in cursor.fetchall():
        reports.append({
            'work_date': row[0],
            'emp_id': row[1],
            'name': row[2],
            'workplace': row[3],
            'work_description': row[4],
            'remark': row[5],
            'status': row[6]
        })

    conn.close()
    return jsonify(reports)

@app.route('/api/reports/export')
def export_reports():
    """Export reports to Excel"""
    from_date = request.args.get('from_date')
    to_date = request.args.get('to_date')
    process = request.args.get('process', 'All')

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    query = '''
        SELECT ot.work_date as 'วันที่',
               ot.emp_id as 'รหัสพนักงาน',
               e.first_name || ' ' || e.last_name as 'ชื่อ-นามสกุล',
               ot.workplace as 'สถานที่ทำงาน',
               ot.work_description as 'รายการงาน',
               ot.remark as 'Remark',
               ot.status as 'สถานะ'
        FROM ot_records ot
        JOIN employees e ON ot.emp_id = e.emp_id
        WHERE ot.work_date BETWEEN ? AND ?
    '''
    params = [from_date, to_date]

    if process and process != 'All':
        query += ' AND ot.workplace = ?'
        params.append(process)

    query += ' ORDER BY ot.work_date, e.gl_number, ot.emp_id'

    df = pd.read_sql_query(query, conn, params=params)
    conn.close()

    # Create Excel file
    filename = f"OT_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    filepath = os.path.join(STATIC_DIR, filename)
    df.to_excel(filepath, index=False, engine='openpyxl')

    return send_file(filepath, as_attachment=True, download_name=filename)

@app.route('/api/employees/import', methods=['POST'])
def import_employees():
    """Import employees from Excel"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'ไม่พบไฟล์'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'success': False, 'message': 'ไม่ได้เลือกไฟล์'}), 400

    try:
        df = pd.read_excel(file)

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        # Clear existing employees
        cursor.execute('DELETE FROM employees')

        # Import employees
        count = 0
        for _, row in df.iterrows():
            try:
                workplace = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ''
                group_name = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ''
                emp_id = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ''
                first_name = str(row.iloc[3]) if pd.notna(row.iloc[3]) else ''
                last_name = str(row.iloc[4]) if pd.notna(row.iloc[4]) else ''
                emp_type = str(row.iloc[5]) if pd.notna(row.iloc[5]) else ''
                shift = str(row.iloc[6]) if pd.notna(row.iloc[6]) else ''
                supervisor = str(row.iloc[7]) if pd.notna(row.iloc[7]) else ''

                if emp_id and first_name:
                    cursor.execute('''
                        INSERT INTO employees (emp_id, first_name, last_name, workplace, group_name, emp_type, shift, supervisor)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (emp_id, first_name, last_name, workplace, group_name, emp_type, shift, supervisor))
                    count += 1
            except Exception as e:
                continue

        conn.commit()
        conn.close()

        return jsonify({'success': True, 'message': f'นำเข้าข้อมูลพนักงาน {count} คน'})

    except Exception as e:
        return jsonify({'success': False, 'message': f'เกิดข้อผิดพลาด: {str(e)}'}), 500

@app.route('/api/employees/list')
def list_employees():
    """List all employees"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute('''
        SELECT emp_id, first_name, last_name, workplace, group_name, emp_type, shift, supervisor
        FROM employees
        ORDER BY group_name, emp_id
    ''')

    employees = []
    for row in cursor.fetchall():
        employees.append({
            'emp_id': row[0],
            'first_name': row[1],
            'last_name': row[2],
            'name': f"{row[1]} {row[2]}",
            'workplace': row[3],
            'group_name': row[4],
            'emp_type': row[5],
            'shift': row[6],
            'supervisor': row[7]
        })

    conn.close()
    return jsonify(employees)

@app.route('/api/employees/save', methods=['POST'])
def save_employee():
    """Create or update one employee from the editable table."""
    data = request.json or {}
    original_emp_id = (data.get('original_emp_id') or data.get('emp_id') or '').strip()
    emp_id = (data.get('emp_id') or '').strip()
    first_name = (data.get('first_name') or '').strip()
    last_name = (data.get('last_name') or '').strip()
    workplace = (data.get('workplace') or '').strip()
    group_name = (data.get('group_name') or '').strip()
    emp_type = (data.get('emp_type') or '').strip()
    shift = (data.get('shift') or '').strip()
    supervisor = (data.get('supervisor') or '').strip()

    if not emp_id or not first_name:
        return jsonify({'success': False, 'message': 'กรุณาระบุรหัสพนักงานและชื่อ'}), 400

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    try:
        if original_emp_id:
            cursor.execute('SELECT COUNT(*) FROM employees WHERE emp_id = ?', (original_emp_id,))
            exists = cursor.fetchone()[0] > 0
        else:
            exists = False

        if exists:
            if emp_id != original_emp_id:
                cursor.execute(
                    'SELECT COUNT(*) FROM employees WHERE emp_id = ? AND emp_id != ?',
                    (emp_id, original_emp_id)
                )
                if cursor.fetchone()[0] > 0:
                    return jsonify({'success': False, 'message': 'รหัสพนักงานนี้มีอยู่แล้ว'}), 409

            cursor.execute('''
                UPDATE employees
                SET emp_id = ?, first_name = ?, last_name = ?, workplace = ?,
                    group_name = ?, emp_type = ?, shift = ?, supervisor = ?
                WHERE emp_id = ?
            ''', (emp_id, first_name, last_name, workplace, group_name, emp_type, shift, supervisor, original_emp_id))

            if emp_id != original_emp_id:
                cursor.execute('UPDATE ot_records SET emp_id = ? WHERE emp_id = ?', (emp_id, original_emp_id))
                cursor.execute('UPDATE attendance_records SET emp_id = ? WHERE emp_id = ?', (emp_id, original_emp_id))

            message = 'บันทึกการแก้ไขพนักงานเรียบร้อยแล้ว'
        else:
            cursor.execute('''
                INSERT INTO employees (emp_id, first_name, last_name, workplace, group_name, emp_type, shift, supervisor)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (emp_id, first_name, last_name, workplace, group_name, emp_type, shift, supervisor))
            message = 'เพิ่มพนักงานใหม่เรียบร้อยแล้ว'

        conn.commit()
        return jsonify({'success': True, 'message': message})
    except sqlite3.IntegrityError:
        conn.rollback()
        return jsonify({'success': False, 'message': 'รหัสพนักงานนี้มีอยู่แล้ว'}), 409
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': f'เกิดข้อผิดพลาด: {str(e)}'}), 500
    finally:
        conn.close()

@app.route('/api/employees/delete', methods=['POST'])
def delete_employee():
    """Delete one employee from the roster."""
    data = request.json or {}
    emp_id = (data.get('emp_id') or '').strip()

    if not emp_id:
        return jsonify({'success': False, 'message': 'ไม่พบรหัสพนักงาน'}), 400

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM employees WHERE emp_id = ?', (emp_id,))
    deleted = cursor.rowcount
    conn.commit()
    conn.close()

    if deleted == 0:
        return jsonify({'success': False, 'message': 'ไม่พบพนักงานที่ต้องการลบ'}), 404

    return jsonify({'success': True, 'message': 'ลบพนักงานเรียบร้อยแล้ว'})

@app.route('/api/employees/supervisors/delete', methods=['POST'])
def delete_supervisor_name():
    """Remove a supervisor name from all employees that currently use it."""
    data = request.json or {}
    supervisor = (data.get('supervisor') or '').strip()

    if not supervisor:
        return jsonify({'success': False, 'message': 'กรุณาเลือกชื่อหัวหน้างานที่ต้องการลบ'}), 400

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('UPDATE employees SET supervisor = "" WHERE supervisor = ?', (supervisor,))
    updated = cursor.rowcount
    conn.commit()
    conn.close()

    return jsonify({
        'success': True,
        'message': f'ลบชื่อหัวหน้างาน "{supervisor}" ออกจาก {updated} รายการเรียบร้อยแล้ว',
        'updated': updated
    })

@app.route('/api/employees/export')
def export_employees():
    """Export employee master data to Excel."""
    conn = sqlite3.connect(DB_PATH)
    query = '''
        SELECT
            workplace as 'สถานที่ทำงาน',
            group_name as 'Group',
            emp_id as 'รหัสพนักงาน',
            first_name as 'ชื่อ',
            last_name as 'นามสกุล',
            emp_type as 'ประเภท',
            shift as 'Shift',
            supervisor as 'หัวหน้างาน'
        FROM employees
        ORDER BY group_name, emp_id
    '''
    df = pd.read_sql_query(query, conn)
    conn.close()

    filename = f"Employee_Master_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    filepath = os.path.join(STATIC_DIR, filename)
    df.to_excel(filepath, index=False, engine='openpyxl')

    return send_file(filepath, as_attachment=True, download_name=filename)

@app.route('/api/employees/template')
def download_employee_template():
    """Download an Excel template for employee import."""
    columns = [
        'สถานที่ทำงาน',
        'Group',
        'รหัสพนักงาน',
        'ชื่อ',
        'นามสกุล',
        'ประเภท',
        'Shift',
        'หัวหน้างาน'
    ]
    example_rows = [
        {
            'สถานที่ทำงาน': 'Fin press machine no.01',
            'Group': 'Fin press',
            'รหัสพนักงาน': '10001',
            'ชื่อ': 'สมชาย',
            'นามสกุล': 'ตัวอย่าง',
            'ประเภท': 'MCP',
            'Shift': 'A',
            'หัวหน้างาน': 'ภาคภูมิ พรมชา'
        },
        {
            'สถานที่ทำงาน': 'Hair pin bender no.01',
            'Group': 'Hairpin bender',
            'รหัสพนักงาน': 'S0001',
            'ชื่อ': 'อารี',
            'นามสกุล': 'ตัวอย่าง',
            'ประเภท': 'SUB',
            'Shift': 'B',
            'หัวหน้างาน': 'วัชรา ศรีบาลชื่น'
        }
    ]
    df = pd.DataFrame(example_rows, columns=columns)
    filename = 'Employee_Import_Template.xlsx'
    filepath = os.path.join(STATIC_DIR, filename)

    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Employees')
        worksheet = writer.sheets['Employees']
        widths = [24, 18, 16, 18, 20, 12, 10, 24]
        for index, width in enumerate(widths, start=1):
            worksheet.column_dimensions[chr(64 + index)].width = width

    return send_file(filepath, as_attachment=True, download_name=filename)

@app.route('/api/employees/clear', methods=['POST'])
def clear_employees():
    """Clear all employees"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM employees')
    conn.commit()
    conn.close()

    return jsonify({'success': True, 'message': 'ลบข้อมูลพนักงานทั้งหมดเรียบร้อยแล้ว'})

@app.route('/api/employees/count')
def count_employees():
    """Count total employees"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('SELECT COUNT(*) FROM employees')
    count = cursor.fetchone()[0]
    conn.close()

    return jsonify({'count': count})

# NEW: Attendance Recording APIs
@app.route('/api/attendance/record', methods=['POST'])
def record_attendance():
    """Record attendance (work/leave/ot)"""
    data = request.json
    work_date = data.get('work_date')
    emp_id = data.get('emp_id')
    attendance_type = data.get('attendance_type')  # 'work', 'leave', 'ot'
    workplace_substitute = data.get('workplace_substitute')  # For OT substitute
    leave_period = data.get('leave_period')
    leave_time = data.get('leave_time')

    if not work_date or not emp_id or not attendance_type:
        return jsonify({'success': False, 'message': 'ข้อมูลไม่ครบถ้วน'}), 400

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    try:
        # Insert or replace attendance record
        cursor.execute('''
            INSERT OR REPLACE INTO attendance_records
            (work_date, emp_id, attendance_type, workplace_substitute, leave_period, leave_time)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (work_date, emp_id, attendance_type, workplace_substitute, leave_period, leave_time))

        conn.commit()
        conn.close()

        # Get employee name for success message
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT first_name, last_name FROM employees WHERE emp_id = ?', (emp_id,))
        result = cursor.fetchone()
        conn.close()

        emp_name = f"{result[0]} {result[1]}" if result else emp_id

        type_text = {
            'work': 'มาทำงาน',
            'leave': 'ลางาน',
            'ot': 'ทำ OT'
        }

        return jsonify({
            'success': True,
            'message': f'บันทึก {type_text.get(attendance_type, attendance_type)} สำหรับ {emp_name} เรียบร้อยแล้ว'
        })

    except Exception as e:
        conn.rollback()
        conn.close()
        return jsonify({'success': False, 'message': f'เกิดข้อผิดพลาด: {str(e)}'}), 500

@app.route('/api/attendance/check')
def check_attendance():
    """Check attendance records for a specific date"""
    work_date = request.args.get('work_date')

    if not work_date:
        return jsonify({'success': False, 'message': 'กรุณาระบุวันที่'}), 400

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute('''
        SELECT emp_id, attendance_type, workplace_substitute, leave_period, leave_time
        FROM attendance_records
        WHERE work_date = ?
    ''', (work_date,))

    records = {}
    for row in cursor.fetchall():
        emp_id = row[0]
        attendance_type = row[1]
        workplace_substitute = row[2]
        leave_period = row[3]
        leave_time = row[4]

        if emp_id not in records:
            records[emp_id] = {
                'types': [],
                'workplace_substitute': None,
                'leave_period': None,
                'leave_time': None
            }
        records[emp_id]['types'].append(attendance_type)

        # Store workplace_substitute if it's an OT record
        if attendance_type == 'ot' and workplace_substitute:
            records[emp_id]['workplace_substitute'] = workplace_substitute
        if attendance_type == 'leave':
            records[emp_id]['leave_period'] = leave_period
            records[emp_id]['leave_time'] = leave_time

    conn.close()
    return jsonify(records)

@app.route('/api/attendance/leave-details')
def get_leave_details():
    """Get detailed leave information for a specific date"""
    work_date = request.args.get('work_date')

    if not work_date:
        return jsonify({'success': False, 'message': 'กรุณาระบุวันที่'}), 400

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Get all employees on leave with their details
    cursor.execute('''
        SELECT
            ar.emp_id,
            e.first_name,
            e.last_name,
            e.workplace,
            ar.leave_period,
            ar.leave_time
        FROM attendance_records ar
        LEFT JOIN employees e ON ar.emp_id = e.emp_id
        WHERE ar.work_date = ? AND ar.attendance_type = 'leave'
        ORDER BY ar.emp_id
    ''', (work_date,))

    leave_records = []
    for row in cursor.fetchall():
        first_name = row[1] or ''
        last_name = row[2] or ''
        full_name = f"{first_name} {last_name}".strip() if first_name or last_name else '-'

        leave_records.append({
            'emp_id': row[0],
            'emp_name': full_name,
            'workplace': row[3] or '-',
            'leave_period': row[4] or '-',
            'leave_time': row[5] or '-'
        })

    conn.close()

    return jsonify({
        'success': True,
        'work_date': work_date,
        'count': len(leave_records),
        'records': leave_records
    })

@app.route('/api/attendance/delete', methods=['POST'])
def delete_attendance():
    """Delete attendance record (for correction)"""
    data = request.json
    work_date = data.get('work_date')
    emp_id = data.get('emp_id')
    attendance_type = data.get('attendance_type')

    if not work_date or not emp_id or not attendance_type:
        return jsonify({'success': False, 'message': 'ข้อมูลไม่ครบถ้วน'}), 400

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    try:
        cursor.execute('''
            DELETE FROM attendance_records
            WHERE work_date = ? AND emp_id = ? AND attendance_type = ?
        ''', (work_date, emp_id, attendance_type))

        conn.commit()
        conn.close()

        # Get employee name for success message
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT first_name, last_name FROM employees WHERE emp_id = ?', (emp_id,))
        result = cursor.fetchone()
        conn.close()

        emp_name = f"{result[0]} {result[1]}" if result else emp_id

        type_text = {
            'work': 'มาทำงาน',
            'leave': 'ลางาน',
            'ot': 'ทำ OT'
        }

        return jsonify({
            'success': True,
            'message': f'ยกเลิก {type_text.get(attendance_type, attendance_type)} สำหรับ {emp_name} เรียบร้อยแล้ว'
        })

    except Exception as e:
        conn.rollback()
        conn.close()
        return jsonify({'success': False, 'message': f'เกิดข้อผิดพลาด: {str(e)}'}), 500

@app.route('/api/attendance/history')
def get_attendance_history():
    """Get attendance history for a specific date"""
    work_date = request.args.get('work_date')

    if not work_date:
        return jsonify({'success': False, 'message': 'กรุณาระบุวันที่'}), 400

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Get all attendance records with employee details
    cursor.execute('''
        SELECT
            a.emp_id,
            e.first_name || ' ' || e.last_name as name,
            e.workplace,
            e.group_name,
            e.emp_type,
            e.shift,
            a.attendance_type,
            a.created_at,
            a.leave_period,
            a.leave_time
        FROM attendance_records a
        JOIN employees e ON a.emp_id = e.emp_id
        WHERE a.work_date = ?
        ORDER BY a.attendance_type, e.group_name, a.emp_id
    ''', (work_date,))

    records = []
    for row in cursor.fetchall():
        records.append({
            'emp_id': row[0],
            'name': row[1],
            'workplace': row[2],
            'group_name': row[3],
            'emp_type': row[4],
            'shift': row[5],
            'attendance_type': row[6],
            'created_at': row[7],
            'leave_period': row[8],
            'leave_time': row[9]
        })

    # Count by type
    cursor.execute('''
        SELECT attendance_type, COUNT(*)
        FROM attendance_records
        WHERE work_date = ?
        GROUP BY attendance_type
    ''', (work_date,))

    summary = {}
    for row in cursor.fetchall():
        summary[row[0]] = row[1]

    conn.close()

    return jsonify({
        'success': True,
        'work_date': work_date,
        'records': records,
        'summary': summary
    })

@app.route('/api/holidays')
def get_holidays():
    """Get holidays/special workdays for a year or date range"""
    year      = request.args.get('year')
    from_date = request.args.get('from_date')
    to_date   = request.args.get('to_date')
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    if from_date and to_date:
        cursor.execute('SELECT date, description, type FROM holidays WHERE date BETWEEN ? AND ? ORDER BY date', (from_date, to_date))
    elif year:
        cursor.execute("SELECT date, description, type FROM holidays WHERE date LIKE ? ORDER BY date", (f"{year}-%",))
    else:
        cursor.execute('SELECT date, description, type FROM holidays ORDER BY date')
    rows = [{'date': r[0], 'description': r[1], 'type': r[2] or 'holiday'} for r in cursor.fetchall()]
    conn.close()
    return jsonify(rows)

@app.route('/api/holidays', methods=['POST'])
def add_holiday():
    """Add a holiday or special workday"""
    data = request.json
    date        = data.get('date')
    description = data.get('description', 'วันหยุดพิเศษ')
    htype       = data.get('type', 'holiday')  # 'holiday' or 'workday'
    if not date:
        return jsonify({'success': False, 'message': 'กรุณาระบุวันที่'}), 400
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('INSERT OR REPLACE INTO holidays (date, description, type) VALUES (?, ?, ?)', (date, description, htype))
        conn.commit()
        return jsonify({'success': True, 'message': f'บันทึก {date} เรียบร้อย'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500
    finally:
        conn.close()

@app.route('/api/holidays/<date>', methods=['DELETE'])
def delete_holiday(date):
    """Delete a holiday"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM holidays WHERE date = ?', (date,))
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': f'ลบวันหยุด {date} เรียบร้อย'})


@app.route('/api/employees/lookup')
def lookup_employee():
    """Look up a single employee by code (for auto-filling the name field)."""
    emp_id = (request.args.get('emp_id') or '').strip()
    if not emp_id:
        return jsonify({'success': False, 'message': 'กรุณาระบุรหัสพนักงาน'}), 400
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        'SELECT emp_id, first_name, last_name, shift, workplace FROM employees WHERE emp_id = ?',
        (emp_id,))
    row = cursor.fetchone()
    conn.close()
    if not row:
        return jsonify({'success': False, 'message': 'ไม่พบรหัสพนักงานนี้'}), 404
    return jsonify({'success': True, 'employee': {
        'emp_id': row[0],
        'name': f"{row[1]} {row[2]}",
        'shift': row[3],
        'workplace': row[4]
    }})


@app.route('/api/hour-exclusions')
def get_hour_exclusions():
    """List employees currently excluded from working-hour calculation."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        SELECT x.emp_id, e.first_name, e.last_name, e.shift, e.workplace
        FROM hour_exclusions x
        LEFT JOIN employees e ON e.emp_id = x.emp_id
        ORDER BY x.created_at DESC
    ''')
    items = []
    for row in cursor.fetchall():
        name = f"{row[1]} {row[2]}".strip() if row[1] else '(ไม่พบในฐานข้อมูลพนักงาน)'
        items.append({
            'emp_id': row[0],
            'name': name,
            'shift': row[3] or '-',
            'workplace': row[4] or '-'
        })
    conn.close()
    return jsonify({'success': True, 'exclusions': items})


@app.route('/api/hour-exclusions', methods=['POST'])
def add_hour_exclusion():
    """Add an employee to the working-hour exclusion list (permanent until removed)."""
    data = request.json or {}
    emp_id = (data.get('emp_id') or '').strip()
    if not emp_id:
        return jsonify({'success': False, 'message': 'กรุณาระบุรหัสพนักงาน'}), 400
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        'SELECT first_name, last_name, shift, workplace FROM employees WHERE emp_id = ?',
        (emp_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': 'ไม่พบรหัสพนักงานนี้'}), 404
    cursor.execute('INSERT OR IGNORE INTO hour_exclusions (emp_id) VALUES (?)', (emp_id,))
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': f'ยกเว้นการคิดชั่วโมงของ {emp_id} แล้ว', 'employee': {
        'emp_id': emp_id,
        'name': f"{row[0]} {row[1]}",
        'shift': row[2] or '-',
        'workplace': row[3] or '-'
    }})


@app.route('/api/hour-exclusions/<emp_id>', methods=['DELETE'])
def delete_hour_exclusion(emp_id):
    """Remove an employee from the working-hour exclusion list."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM hour_exclusions WHERE emp_id = ?', (emp_id,))
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': f'ยกเลิกการยกเว้นของ {emp_id} แล้ว'})


@app.route('/api/attendance/incoming-substitutes')
def get_incoming_substitutes():
    """พนักงานจากทีมอื่นที่มาทำ OT แทนให้กับทีมของ GL นี้"""
    work_date = request.args.get('work_date')
    gl        = request.args.get('gl')
    shift     = request.args.get('shift', 'All')

    if not work_date or not gl or gl == 'All':
        return jsonify([])

    conn   = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # หา groups ทั้งหมดของทีม GL นี้
    q      = 'SELECT DISTINCT group_name FROM employees WHERE supervisor = ? AND group_name != ""'
    params = [gl]
    if shift != 'All':
        q += ' AND shift = ?'
        params.append(shift)
    cursor.execute(q, params)
    gl_groups = [row[0] for row in cursor.fetchall()]

    if not gl_groups:
        conn.close()
        return jsonify([])

    # หา shift จริงของทีม GL นี้ใน groups เหล่านั้น
    if shift != 'All':
        required_shifts = [shift]
    else:
        gps_ph = ','.join(['?' for _ in gl_groups])
        cursor.execute(
            f'SELECT DISTINCT shift FROM employees WHERE supervisor = ? AND group_name IN ({gps_ph})',
            [gl] + gl_groups
        )
        required_shifts = [row[0] for row in cursor.fetchall()]

    if not required_shifts:
        conn.close()
        return jsonify([])

    # ดึงพนักงานจากทีมอื่นที่ workplace_substitute (group name) ตรงกับ groups ของทีม GL นี้
    # และต้องเป็น shift เดียวกันด้วย
    gps_ph    = ','.join(['?' for _ in gl_groups])
    shifts_ph = ','.join(['?' for _ in required_shifts])
    cursor.execute(f'''
        SELECT ar.emp_id,
               e.first_name || ' ' || e.last_name AS name,
               e.group_name AS original_group,
               e.supervisor AS original_gl,
               e.shift, e.emp_type,
               ar.workplace_substitute
        FROM attendance_records ar
        JOIN employees e ON ar.emp_id = e.emp_id
        WHERE ar.work_date = ?
          AND ar.attendance_type = 'ot'
          AND ar.workplace_substitute IN ({gps_ph})
          AND (e.supervisor != ? OR e.supervisor IS NULL OR e.supervisor = '')
          AND e.shift IN ({shifts_ph})
        ORDER BY ar.workplace_substitute, e.supervisor, ar.emp_id
    ''', [work_date] + gl_groups + [gl] + required_shifts)

    result = []
    for row in cursor.fetchall():
        result.append({
            'emp_id':            row[0],
            'name':              row[1],
            'original_group':    row[2],
            'original_gl':       row[3] or '-',
            'shift':             row[4],
            'emp_type':          row[5],
            'substitute_group':  row[6]
        })

    conn.close()
    return jsonify(result)


@app.route('/api/attendance/range')
def get_attendance_range():
    """Get attendance records for a date range (max 31 days) — for grid report"""
    from datetime import date as date_cls, timedelta
    from_date = request.args.get('from_date')
    to_date   = request.args.get('to_date')
    process   = request.args.get('process', 'All')
    shift     = request.args.get('shift',   'All')
    gl        = request.args.get('gl',      'All')

    if not from_date or not to_date:
        return jsonify({'success': False, 'message': 'กรุณาระบุช่วงวันที่'}), 400

    start = date_cls.fromisoformat(from_date)
    end   = date_cls.fromisoformat(to_date)

    if (end - start).days > 30:
        return jsonify({'success': False, 'message': 'ช่วงวันที่ต้องไม่เกิน 31 วัน'}), 400

    # Build date list
    dates, cur = [], start
    while cur <= end:
        dates.append(cur.isoformat())
        cur += timedelta(days=1)

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Get employees with filters
    q = 'SELECT emp_id, first_name, last_name, workplace, shift, supervisor FROM employees WHERE 1=1'
    params = []
    if process != 'All':
        q += ' AND workplace = ?'; params.append(process)
    if shift != 'All':
        q += ' AND shift = ?';     params.append(shift)
    if gl != 'All':
        q += ' AND supervisor = ?'; params.append(gl)
    q += ' ORDER BY workplace, emp_id'

    cursor.execute(q, params)
    employees = cursor.fetchall()

    # Get attendance records for date range
    cursor.execute('''
        SELECT emp_id, work_date, attendance_type, workplace_substitute, leave_period, leave_time
        FROM attendance_records
        WHERE work_date BETWEEN ? AND ?
    ''', (from_date, to_date))

    att = {}
    for emp_id, work_date, att_type, workplace_sub, leave_period, leave_time in cursor.fetchall():
        if emp_id not in att:
            att[emp_id] = {}
        if work_date not in att[emp_id]:
            att[emp_id][work_date] = {
                'types': [],
                'workplace_substitute': None,
                'leave_period': None,
                'leave_time': None
            }
        att[emp_id][work_date]['types'].append(att_type)
        if att_type == 'ot' and workplace_sub:
            att[emp_id][work_date]['workplace_substitute'] = workplace_sub
        if att_type == 'leave':
            att[emp_id][work_date]['leave_period'] = leave_period
            att[emp_id][work_date]['leave_time'] = leave_time

    conn.close()

    result = []
    for emp in employees:
        emp_id = emp[0]
        records = {}
        for d in dates:
            if emp_id in att and d in att[emp_id]:
                records[d] = att[emp_id][d]
            else:
                records[d] = {
                    'types': [],
                    'workplace_substitute': None,
                    'leave_period': None,
                    'leave_time': None
                }
        result.append({
            'emp_id':    emp_id,
            'name':      f"{emp[1]} {emp[2]}",
            'workplace': emp[3],
            'shift':     emp[4],
            'supervisor':emp[5],
            'records':   records
        })

    return jsonify({'success': True, 'dates': dates, 'employees': result})


@app.route('/api/attendance/range/export')
def export_attendance_range():
    """Export attendance grid to Excel"""
    from datetime import date as date_cls, timedelta
    from_date = request.args.get('from_date')
    to_date   = request.args.get('to_date')
    process   = request.args.get('process', 'All')
    shift     = request.args.get('shift',   'All')
    gl        = request.args.get('gl',      'All')

    # Reuse range API logic
    with app.test_request_context(
        f'/api/attendance/range?from_date={from_date}&to_date={to_date}'
        f'&process={process}&shift={shift}&gl={gl}'
    ):
        pass

    start = date_cls.fromisoformat(from_date)
    end   = date_cls.fromisoformat(to_date)
    dates, cur = [], start
    while cur <= end:
        dates.append(cur.isoformat())
        cur += timedelta(days=1)

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    q = 'SELECT emp_id, first_name, last_name, workplace, shift, supervisor FROM employees WHERE 1=1'
    params = []
    if process != 'All': q += ' AND workplace = ?'; params.append(process)
    if shift   != 'All': q += ' AND shift = ?';     params.append(shift)
    if gl      != 'All': q += ' AND supervisor = ?'; params.append(gl)
    q += ' ORDER BY workplace, emp_id'
    cursor.execute(q, params)
    employees = cursor.fetchall()

    cursor.execute('SELECT emp_id, work_date, attendance_type FROM attendance_records WHERE work_date BETWEEN ? AND ?', (from_date, to_date))
    att = {}
    for emp_id, work_date, att_type in cursor.fetchall():
        att.setdefault(emp_id, {}).setdefault(work_date, []).append(att_type)
    conn.close()

    OT_WEEKDAY = 2.2
    OT_WEEKEND = 10.32

    # Load custom holidays/workdays for this range
    conn2 = sqlite3.connect(DB_PATH)
    cur2 = conn2.cursor()
    cur2.execute('SELECT date, type FROM holidays WHERE date BETWEEN ? AND ?', (from_date, to_date))
    holiday_set  = set()
    workday_set  = set()
    for r in cur2.fetchall():
        (holiday_set if (r[1] or 'holiday') == 'holiday' else workday_set).add(r[0])
    conn2.close()

    rows = []
    for emp in employees:
        emp_id = emp[0]
        row = {'รหัส': emp_id, 'ชื่อ-นามสกุล': f"{emp[1]} {emp[2]}", 'Process': emp[3], 'Shift': emp[4]}
        work_count = leave_count = ot_count = 0
        ot_hours = 0.0
        for d in dates:
            types = att.get(emp_id, {}).get(d, [])
            label = d[8:] + '/' + d[5:7]
            dow = date_cls.fromisoformat(d).weekday()  # 0=จ ... 6=อา
            is_weekend = (dow >= 5 or d in holiday_set) and d not in workday_set

            if 'leave' in types:
                row[label] = 'X'
                leave_count += 1
            elif 'work' in types and 'ot' in types:
                row[label] = 'OT'
                work_count += 1
                ot_count   += 1
                ot_hours   += OT_WEEKEND if is_weekend else OT_WEEKDAY
            elif 'work' in types:
                row[label] = '✓'
                work_count += 1
            elif 'ot' in types:
                row[label] = 'OT'
                ot_count   += 1
                ot_hours   += OT_WEEKEND if is_weekend else OT_WEEKDAY
            else:
                row[label] = '-'

        row['มา (วัน)']      = work_count  if work_count  else '-'
        row['ลา (วัน)']      = leave_count if leave_count else '-'
        row['OT (วัน)']      = ot_count    if ot_count    else '-'
        row['ชั่วโมง OT']    = round(ot_hours, 2) if ot_hours else '-'
        rows.append(row)

    df = pd.DataFrame(rows)
    filename = f"Attendance_{from_date}_to_{to_date}.xlsx"
    filepath = os.path.join(STATIC_DIR, filename)
    df.to_excel(filepath, index=False, engine='openpyxl')
    return send_file(filepath, as_attachment=True, download_name=filename)


@app.route('/api/attendance/reset-all', methods=['POST'])
def reset_all_attendance():
    data = request.json or {}
    pin = data.get('pin', '')
    if pin != '47117257':
        return jsonify({'success': False, 'message': 'รหัสยืนยันไม่ถูกต้อง'}), 403
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM attendance_records')
    count = cursor.rowcount
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': f'ลบข้อมูลการบันทึกทั้งหมด {count} รายการเรียบร้อยแล้ว'})


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
