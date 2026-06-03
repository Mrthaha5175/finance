import os
import sqlite3
from datetime import date
from io import BytesIO
from functools import wraps

from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = os.urandom(24)

DB_PATH = os.path.join(os.path.dirname(__file__), 'finance.db')
USERS = {
    'admin': 'password',
}

SAMPLE_RECORDS = [
    ('Alice', '2026-06', 1200.0, 800.0, 'Rent payment'),
    ('Bob', '2026-06', 950.0, 950.0, 'Consulting fee'),
    ('Charlie', '2026-06', 670.0, 0.0, 'Utility bills'),
]


def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    os.makedirs(os.path.dirname(DB_PATH) or '.', exist_ok=True)
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS finance_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            transaction_date TEXT NOT NULL,
            month TEXT NOT NULL,
            due_amount REAL NOT NULL DEFAULT 0,
            payment_amount REAL NOT NULL DEFAULT 0,
            notes TEXT
        )
        """
    )
    conn.commit()

    cursor.execute('PRAGMA table_info(finance_records)')
    columns = [row[1] for row in cursor.fetchall()]
    if 'due_amount' not in columns or 'transaction_date' not in columns:
        old_rows = cursor.execute('SELECT id, name, month, monthly_amount, paid, notes FROM finance_records').fetchall()
        cursor.execute('ALTER TABLE finance_records RENAME TO finance_records_old')
        cursor.execute(
            """
            CREATE TABLE finance_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                transaction_date TEXT NOT NULL,
                month TEXT NOT NULL,
                due_amount REAL NOT NULL DEFAULT 0,
                payment_amount REAL NOT NULL DEFAULT 0,
                notes TEXT
            )
            """
        )
        for row in old_rows:
            _, name, month, monthly_amount, paid, notes = row
            transaction_date = f"{month}-01"
            cursor.execute(
                'INSERT INTO finance_records (id, name, transaction_date, month, due_amount, payment_amount, notes) VALUES (?, ?, ?, ?, ?, ?, ?)',
                (row[0], name, transaction_date, month, monthly_amount, paid, notes),
            )
        conn.commit()
        cursor.execute('DROP TABLE finance_records_old')
        conn.commit()

    cursor.execute('SELECT COUNT(1) FROM finance_records')
    if cursor.fetchone()[0] == 0:
        today = date.today().isoformat()
        example_records = [
            ('Alice', today, today[:7], 1200.0, 800.0, 'Rent payment'),
            ('Bob', today, today[:7], 950.0, 950.0, 'Consulting fee'),
            ('Charlie', today, today[:7], 670.0, 0.0, 'Utility bills'),
        ]
        cursor.executemany(
            'INSERT INTO finance_records (name, transaction_date, month, due_amount, payment_amount, notes) VALUES (?, ?, ?, ?, ?, ?)',
            example_records,
        )
        conn.commit()

    conn.close()


def fetch_all_records():
    conn = get_db_connection()
    rows = conn.execute('SELECT * FROM finance_records ORDER BY id').fetchall()
    conn.close()
    return [dict(row) for row in rows]


def fetch_record(record_id):
    conn = get_db_connection()
    row = conn.execute('SELECT * FROM finance_records WHERE id = ?', (record_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def create_record(name, transaction_date, month, due_amount, payment_amount, notes):
    conn = get_db_connection()
    conn.execute(
        'INSERT INTO finance_records (name, transaction_date, month, due_amount, payment_amount, notes) VALUES (?, ?, ?, ?, ?, ?)',
        (name, transaction_date, month, due_amount, payment_amount, notes),
    )
    conn.commit()
    conn.close()


def update_record(record_id, name, transaction_date, month, due_amount, payment_amount, notes):
    conn = get_db_connection()
    conn.execute(
        'UPDATE finance_records SET name = ?, transaction_date = ?, month = ?, due_amount = ?, payment_amount = ?, notes = ? WHERE id = ?',
        (name, transaction_date, month, due_amount, payment_amount, notes, record_id),
    )
    conn.commit()
    conn.close()


def login_required(view):
    @wraps(view)
    def wrapped_view(*args, **kwargs):
        if not session.get('user'):
            return redirect(url_for('login'))
        return view(*args, **kwargs)

    return wrapped_view


@app.route('/', methods=['GET', 'POST'])
def login():
    if session.get('user'):
        return redirect(url_for('dashboard'))

    error = None
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')

        if USERS.get(username) == password:
            session['user'] = username
            return redirect(url_for('dashboard'))

        error = 'Invalid username or password'

    return render_template('login.html', error=error)


@app.route('/dashboard')
@login_required
def dashboard():
    records = fetch_all_records()
    total_due = sum(record['due_amount'] for record in records)
    total_paid = sum(record['payment_amount'] for record in records)
    total_pending = total_due - total_paid

    person_summary = {}
    for record in records:
        person = record['name']
        summary = person_summary.setdefault(person, {'due': 0, 'paid': 0})
        summary['due'] += record['due_amount']
        summary['paid'] += record['payment_amount']

    for summary in person_summary.values():
        summary['pending'] = summary['due'] - summary['paid']

    return render_template(
        'dashboard.html',
        records=records,
        total_due=total_due,
        total_paid=total_paid,
        total_pending=total_pending,
        person_summary=person_summary,
    )


@app.route('/update', methods=['GET', 'POST'])
@login_required
def update():
    edit_id = request.args.get('edit_id', type=int)
    edit_record = fetch_record(edit_id) if edit_id else None

    if request.method == 'POST':
        record_id = request.form.get('record_id')
        record_id = int(record_id) if record_id else None
        name = request.form.get('name', '').strip()
        month = request.form.get('month', '').strip()
        transaction_date = request.form.get('transaction_date', '').strip()
        month = request.form.get('month', '').strip()
        transaction_type = request.form.get('transaction_type', 'due')
        amount = request.form.get('amount', type=float, default=0)
        notes = request.form.get('notes', '').strip()

        if not name or not transaction_date or not month:
            flash('Name, date, and month are required.', 'error')
            return redirect(url_for('update'))

        due_amount = amount if transaction_type == 'due' else 0
        payment_amount = amount if transaction_type == 'payment' else 0

        if record_id:
            update_record(record_id, name, transaction_date, month, due_amount, payment_amount, notes)
            flash('Record updated successfully.', 'success')
        else:
            create_record(name, transaction_date, month, due_amount, payment_amount, notes)
            flash('New finance record added.', 'success')

        return redirect(url_for('update'))

    records = fetch_all_records()
    return render_template('update.html', records=records, edit_record=edit_record)


@app.route('/customer/<name>', methods=['GET', 'POST'])
@login_required
def customer(name):
    if request.method == 'POST':
        transaction_date = request.form.get('transaction_date', '').strip()
        transaction_type = request.form.get('transaction_type', 'due')
        amount = request.form.get('amount', type=float, default=0)
        notes = request.form.get('notes', '').strip()
        month = transaction_date[:7] if transaction_date else date.today().strftime('%Y-%m')

        due_amount = amount if transaction_type == 'due' else 0
        payment_amount = amount if transaction_type == 'payment' else 0
        create_record(name, transaction_date, month, due_amount, payment_amount, notes)
        flash('Customer transaction updated.', 'success')
        return redirect(url_for('customer', name=name))

    records = fetch_all_records()
    customer_records = [record for record in records if record['name'] == name]
    total_due = sum(record['due_amount'] for record in customer_records)
    total_paid = sum(record['payment_amount'] for record in customer_records)
    total_pending = total_due - total_paid
    return render_template(
        'customer.html',
        name=name,
        records=customer_records,
        total_due=total_due,
        total_paid=total_paid,
        total_pending=total_pending,
    )


@app.route('/calendar')
@login_required
def calendar():
    selected_month = request.args.get('month')
    if not selected_month:
        selected_month = date.today().strftime('%Y-%m')

    records = fetch_all_records()
    month_records = [record for record in records if record['month'] == selected_month]
    total_due = sum(record['due_amount'] for record in month_records)
    total_paid = sum(record['payment_amount'] for record in month_records)
    total_pending = total_due - total_paid

    return render_template(
        'calendar.html',
        records=month_records,
        selected_month=selected_month,
        total_due=total_due,
        total_paid=total_paid,
        total_pending=total_pending,
    )


@app.route('/dues')
@login_required
def dues():
    records = fetch_all_records()
    pending = [record for record in records if record['due_amount'] - record['payment_amount'] > 0]
    paid = [record for record in records if record['due_amount'] - record['payment_amount'] <= 0]
    return render_template('dues.html', pending=pending, paid=paid)


@app.route('/download-excel')
@login_required
def download_excel():
    records = fetch_all_records()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 'Client Dues'
    sheet.append(['ID', 'Name', 'Date', 'Month', 'Due Amount', 'Payment Amount', 'Pending', 'Notes'])

    for record in records:
        sheet.append([
            record['id'],
            record['name'],
            record['transaction_date'],
            record['month'],
            record['due_amount'],
            record['payment_amount'],
            record['due_amount'] - record['payment_amount'],
            record['notes'],
        ])

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return send_file(
        output,
        download_name='finance_dues.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@app.route('/logout')
@login_required
def logout():
    session.clear()
    return redirect(url_for('login'))


if __name__ == '__main__':
    init_db()
    app.run(debug=True)
