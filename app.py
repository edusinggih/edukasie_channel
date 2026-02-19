from flask import Flask, render_template, request, jsonify, send_file, redirect, session, url_for, flash, make_response
from datetime import datetime, timedelta
from io import StringIO, BytesIO
import csv, io
import psycopg2
import psycopg2.extras   # ⬅️ ini penting
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from reportlab.lib.pagesizes import A4, landscape
from werkzeug.security import generate_password_hash, check_password_hash
from functools import wraps



app = Flask(__name__)
app.secret_key = "kunci_rahasia_flask"

# Auto logout setelah 10 menit idle
app.permanent_session_lifetime = timedelta(minutes=10)

# -------- Middleware untuk password halaman manage
@app.before_request
def protect_manage():
    if request.endpoint == "manage_mesin":
        if not session.get("authorized_manage"):
            # Cek password di query string
            pwd = request.args.get("pwd")
            if pwd == "admin123":  # password disini
                session["authorized_manage"] = True
            else:
                return """
                    <script>
                        var pass = prompt("Masukkan password untuk akses Manage:");
                        if (pass) {
                            window.location.href = window.location.pathname + "?pwd=" + pass;
                        } else {
                            window.location.href = "/";
                        }
                    </script>
                """



# ----------------- MANAGE USER -----------------


# =====================================================
# DECORATORS
# =====================================================
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            flash("Silakan login terlebih dahulu!", "danger")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session or session.get("role") != "admin":
            flash("Anda tidak punya akses ke halaman ini!", "danger")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function

# =====================================================
# MIDDLEWARE UNTUK UPDATE SESSION
# =====================================================
@app.before_request
def before_request():
    session.permanent = True  # aktifkan aturan auto timeout 10 menit

# =====================================================
# LOGIN / LOGOUT
# =====================================================
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        conn = get_db_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
        cur.execute("SELECT * FROM users WHERE username = %s", (username,))
        user = cur.fetchone()
        conn.close()

        if user and check_password_hash(user["password"], password):
            session.permanent = True
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["role"] = user["role"]

            flash("Login berhasil", "success")
            return redirect(url_for("metal_input"))
        else:
            flash("Username atau password salah", "danger")

    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    flash("Anda sudah logout", "info")
    return redirect(url_for("login"))

# =====================================================
# MANAGE USER (HANYA ADMIN)
# =====================================================
@app.route("/manage-user")
@admin_required
def manage_user():
    conn = get_db_connection()
    cur = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
    cur.execute("SELECT * FROM users ORDER BY id ASC")
    users = cur.fetchall()
    cur.close()
    conn.close()
    return render_template("user.html", users=users)

@app.route("/manage-user/add", methods=["POST"])
@admin_required
def add_user():
    username = request.form["username"]
    password = request.form["password"]
    role = request.form["role"]

    hashed_pw = generate_password_hash(password)

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO users (username, password, role) VALUES (%s, %s, %s)",
        (username, hashed_pw, role),
    )
    conn.commit()
    cur.close()
    conn.close()

    flash("User berhasil ditambahkan!", "success")
    return redirect(url_for("manage_user"))

@app.route("/manage-user/edit/<int:id>", methods=["POST"])
@admin_required
def edit_user(id):
    username = request.form["username"]
    password = request.form["password"]
    role = request.form["role"]

    conn = get_db_connection()
    cur = conn.cursor()

    if password.strip() == "":
        cur.execute(
            "UPDATE users SET username=%s, role=%s WHERE id=%s",
            (username, role, id),
        )
    else:
        hashed_pw = generate_password_hash(password)
        cur.execute(
            "UPDATE users SET username=%s, password=%s, role=%s WHERE id=%s",
            (username, hashed_pw, role, id),
        )

    conn.commit()
    cur.close()
    conn.close()

    flash("User berhasil diupdate!", "success")
    return redirect(url_for("manage_user"))

@app.route("/manage-user/delete/<int:id>", methods=["POST"])
@admin_required
def delete_user(id):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("DELETE FROM users WHERE id=%s", (id,))
    conn.commit()
    cur.close()
    conn.close()

    flash("User berhasil dihapus!", "danger")
    return redirect(url_for("manage_user"))

# =====================================================
# METAL INPUT (UNTUK USER & ADMIN)
# =====================================================





# ==================== KONEKSI DATABASE ====================
def get_db_connection():
    return psycopg2.connect(
        dbname='modif_monitoring',
        user='postgres',
        password='rosululloh',
        host='localhost',
        port='5432'
    )


reset_flag = False

# ==================== ROUTES ====================
@app.route('/')
def home():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT nama_mesin, status, waktu_awal_status, terakhir_update FROM mesin")
    rows = cur.fetchall()
    conn.close()

    now = datetime.now()
    status = {}
    for nama_mesin, stat, waktu_awal, last_upd in rows:
        durasi = int((now - waktu_awal).total_seconds())
        if last_upd is None or (now - last_upd).total_seconds() > 10:
            stat = 'POWER_OFF'
        status[nama_mesin] = {
            'status': stat,
            'durasi': durasi
        }

    return render_template('home.html', status=status, now=now)

@app.route('/status')
def get_status():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT nama_mesin, status, waktu_awal_status, terakhir_update FROM mesin")
    rows = cur.fetchall()
    conn.close()

    now = datetime.now()
    result = {}
    for nama_mesin, stat, waktu_awal, last_upd in rows:
        durasi = int((now - waktu_awal).total_seconds())
        if last_upd is None or (now - last_upd).total_seconds() > 10:
            stat = 'POWER_OFF'
        result[nama_mesin] = {
            'status': stat,
            'durasi': durasi
        }
    return jsonify(result)

@app.route('/update/<mesin>', methods=['POST'])
def update_status(mesin):
    conn = get_db_connection()
    cur = conn.cursor()
    new_status = request.form['status']
    now = datetime.now()

    cur.execute("SELECT id, status, waktu_awal_status FROM mesin WHERE nama_mesin = %s", (mesin,))
    row = cur.fetchone()

    if row:
        id_mesin, current_status, waktu_awal = row

        if current_status != new_status:
            durasi = int((now - waktu_awal).total_seconds())
            cur.execute("INSERT INTO log_status (id_mesin, status, waktu, durasi) VALUES (%s, %s, %s, %s)",
                        (id_mesin, new_status, now, durasi))
            cur.execute("UPDATE mesin SET status = %s, waktu_awal_status = %s, terakhir_update = %s WHERE id = %s",
                        (new_status, now, now, id_mesin))
        else:
            cur.execute("UPDATE mesin SET terakhir_update = %s WHERE id = %s", (now, id_mesin))
    else:
        cur.execute("INSERT INTO mesin (nama_mesin, status, waktu_awal_status, terakhir_update, halaman) VALUES (%s, %s, %s, %s, %s)",
                    (mesin, new_status, now, now, 'pc32'))  # default halaman jika tidak ada

    conn.commit()
    conn.close()
    return 'OK'

@app.route('/log')
def get_log():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT l.waktu, m.nama_mesin, l.status, l.durasi
        FROM log_status l
        JOIN mesin m ON l.id_mesin = m.id
        ORDER BY l.waktu DESC
        LIMIT 100
    """)
    rows = cur.fetchall()
    conn.close()

    logs = [{
        'waktu': r[0].strftime('%Y-%m-%d %H:%M:%S'),
        'mesin': r[1],
        'status': r[2],
        'durasi': r[3]
    } for r in rows]

    return jsonify(logs)

@app.route('/export-csv')
def export_csv():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT l.waktu, m.nama_mesin, l.status, l.durasi
        FROM log_status l
        JOIN mesin m ON l.id_mesin = m.id
        ORDER BY l.waktu DESC
    """)
    rows = cur.fetchall()
    conn.close()

    si = StringIO()
    writer = csv.writer(si)
    writer.writerow(['Waktu', 'Mesin', 'Status', 'Durasi (detik)'])
    for row in rows:
        writer.writerow([row[0].strftime('%Y-%m-%d %H:%M:%S'), row[1], row[2], row[3]])

    mem = BytesIO()
    mem.write(si.getvalue().encode('utf-8'))
    mem.seek(0)

    return send_file(mem, mimetype='text/csv', download_name='log_mesin.csv', as_attachment=True)

@app.route('/reset-durasi', methods=['POST'])
def reset_durasi():
    now = datetime.now()
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("UPDATE mesin SET waktu_awal_status = %s", (now,))
    conn.commit()
    conn.close()
    return 'Durasi direset'

@app.route('/reset-esp')
def reset_esp_route():
    global reset_flag
    reset_flag = True
    return 'ESP reset triggered'

@app.route('/check-reset')
def check_reset():
    global reset_flag
    if reset_flag:
        reset_flag = False
        return "RESET"
    return "OK"

@app.route('/log-filtered')
def log_filtered():
    mesin = request.args.get('mesin')
    bulan = request.args.get('bulan')
    tahun = request.args.get('tahun')

    conn = get_db_connection()
    cur = conn.cursor()
    query = """
        SELECT l.waktu, m.nama_mesin, l.status, l.durasi
        FROM log_status l
        JOIN mesin m ON l.id_mesin = m.id
        WHERE 1=1
    """
    params = []

    if mesin and mesin != 'ALL':
        query += " AND m.nama_mesin = %s"
        params.append(mesin)

    if bulan and tahun:
        query += " AND EXTRACT(MONTH FROM l.waktu) = %s AND EXTRACT(YEAR FROM l.waktu) = %s"
        params.extend([int(bulan), int(tahun)])

    query += " ORDER BY l.waktu DESC LIMIT 100"
    cur.execute(query, tuple(params))
    rows = cur.fetchall()
    conn.close()

    result = [{
        'waktu': r[0].strftime('%Y-%m-%d %H:%M:%S'),
        'mesin': r[1],
        'status': r[2],
        'durasi': r[3]
    } for r in rows]

    return jsonify(result)


@app.route('/manage')
def manage_mesin():
    if not session.get("authorized_manage"):
        return render_template("manage.html", mesins=[], need_password=True)

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT id, nama_mesin, status, halaman FROM mesin ORDER BY id")
    rows = cur.fetchall()
    conn.close()

    mesins = [{'id': r[0], 'nama': r[1], 'status': r[2], 'halaman': r[3]} for r in rows]
    return render_template("manage.html", mesins=mesins, need_password=False)

@app.route('/check-manage-password', methods=['POST'])
def check_manage_password():
    data = request.get_json()
    if data.get("password") == "admin123":  # password ganti di sini
        session["authorized_manage"] = True
        return jsonify({"success": True})
    return jsonify({"success": False})


@app.route('/mesin/tambah', methods=['POST'])
def tambah_mesin():
    nama = request.form['nama'].strip().replace(' ', '_').upper()
    halaman = request.form['halaman'].strip().lower()
    now = datetime.now()
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO mesin (nama_mesin, status, waktu_awal_status, terakhir_update, halaman)
        VALUES (%s, %s, %s, %s, %s)
    """, (nama, 'POWER_OFF', now, now, halaman))
    conn.commit()
    conn.close()
    return redirect(f'/{halaman}')

@app.route('/mesin/edit/<int:id>', methods=['POST'])
def edit_mesin(id):
    nama_baru = request.form['nama'].strip().replace(' ', '_').upper()
    halaman = request.form.get('halaman', 'pc32')
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("UPDATE mesin SET nama_mesin = %s, halaman = %s WHERE id = %s", (nama_baru, halaman, id))
    conn.commit()
    conn.close()
    return redirect(f'/{halaman}')

@app.route('/mesin/hapus/<int:id>', methods=['POST'])
def hapus_mesin(id):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("DELETE FROM mesin WHERE id = %s", (id,))
    conn.commit()
    conn.close()
    return redirect('/manage')

@app.route('/clear-log', methods=['POST'])
def clear_log():
    data = request.get_json()
    if not data or data.get('password') != "admin123":
        return "Unauthorized", 403

    try:
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("DELETE FROM log_status")
        conn.commit()
        cur.close()
        conn.close()
        return "Log cleared", 200
    except Exception as e:
        return f"Error: {str(e)}", 500

# ==================== FILTER PER HALAMAN ====================
def get_all_status_by_page(halaman):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT nama_mesin, status, waktu_awal_status, terakhir_update FROM mesin WHERE halaman = %s", (halaman,))
    rows = cur.fetchall()
    conn.close()

    now = datetime.now()
    status = {}
    for nama_mesin, stat, waktu_awal, last_upd in rows:
        durasi = int((now - waktu_awal).total_seconds())
        if last_upd is None or (now - last_upd).total_seconds() > 10:
            stat = 'POWER_OFF'
        status[nama_mesin] = {
            'status': stat,
            'durasi': durasi
        }
    return status



# ===============================
# Monitoring Metal Detector
# ===============================

# Input Metal Detector
from functools import wraps
from flask import session, redirect, url_for, flash, request, render_template

# ===== Middleware wajib login =====
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:   # belum login
            flash("Silakan login terlebih dahulu", "danger")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function


# ===== Route Metal Input =====
from functools import wraps
from flask import session, redirect, url_for, flash

# 🔒 Decorator login_required
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            flash("Silakan login terlebih dahulu!", "danger")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function


@app.route("/metal_input", methods=["GET", "POST"])
@login_required   # ⬅️ WAJIB LOGIN
def metal_input():
    if request.method == "POST":
        line = request.form["line"]
        no_mesin = request.form["no_mesin"]
        type_mesin = request.form["type_mesin"]
        type_md = request.form["type_md"]
        standard_range = request.form["standard_range"]
        note = request.form.get("note", "")

        # Ambil semua kemungkinan input sensitivitas
        product_phase = request.form.get("product_phase")
        fe_phase = request.form.get("fe_phase")
        sus_phase = request.form.get("sus_phase")
        analog_gain = request.form.get("analog_gain")
        digital_gain = request.form.get("digital_gain")
        phase = request.form.get("phase")
        sensitivitas_actual = request.form.get("sensitivitas_actual")

        # Gabungkan menjadi satu string sesuai tipe MD
        if product_phase or fe_phase or sus_phase:
            sens = f"PP:{product_phase}; FE:{fe_phase}; SUS:{sus_phase}"
        elif analog_gain or digital_gain or phase:
            sens = f"AG:{analog_gain}; DG:{digital_gain}; PH:{phase}"
        else:
            sens = sensitivitas_actual

        # Simpan ke DB
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO metal_log 
            (line, no_mesin, type_mesin, type_md, standard_range, sensitivitas_actual, note, user_id, created_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, NOW())
        """, (line, no_mesin, type_mesin, type_md, standard_range, sens, note, session["user_id"]))
        conn.commit()
        conn.close()

        flash("Data berhasil disimpan", "success")
        return redirect(url_for("metal_input"))

    # GET request → ambil daftar line
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT line FROM mesin_master ORDER BY line")
    lines = [{"line": r[0]} for r in cur.fetchall()]
    conn.close()

    return render_template("metal_input.html", lines=lines)





from datetime import datetime

@app.route("/metal/report")
@login_required
def metal_report():
    # Ambil bulan dari query parameter atau default ke bulan saat ini
    bulan = request.args.get("bulan")
    if not bulan:
        bulan = datetime.now().strftime("%Y-%m")  # format 'YYYY-MM'

    line = request.args.get("line", "all")

    # Query data log dengan join user
    with get_db_connection() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
        query = """
            SELECT m.*, u.username AS user_input
            FROM metal_log m
            LEFT JOIN users u ON m.user_id = u.id
            WHERE to_char(m.created_at, 'YYYY-MM') = %s
        """
        params = [bulan]
        if line != "all":
            query += " AND m.line = %s"
            params.append(line)
        query += " ORDER BY m.created_at DESC"
        cur.execute(query, params)
        records = cur.fetchall()

    # Ambil daftar line untuk filter
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("SELECT DISTINCT line FROM metal_log ORDER BY line")
        lines = [row[0] for row in cur.fetchall()]

    # Konversi bulan ke format nama bulan
    bulan_obj = datetime.strptime(bulan, "%Y-%m")
    periode_text = bulan_obj.strftime("%B %Y")  # Contoh: "Agustus 2025"

    return render_template("metal_report.html",
                           records=records,
                           bulan=bulan,
                           line=line,
                           lines=lines,
                           periode_text=periode_text)





@app.route("/metal_master/export")
def export_metal_master():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT line,no_mesin,type_mesin,type_md,standard FROM mesin_master ORDER BY line,no_mesin")
    si = io.StringIO()
    cw = csv.writer(si)
    for r in cur.fetchall():
        cw.writerow(r)
    output = make_response(si.getvalue())
    output.headers["Content-Disposition"] = "attachment; filename=mesin_master.csv"
    output.headers["Content-type"] = "text/csv"
    return output

@app.route("/metal_master/import", methods=["POST"])
def import_metal_master():
    file = request.files.get('file')
    if not file:
        flash("Tidak ada file yang diunggah.", "danger")
        return redirect(url_for('metal_master'))

    # Baca file mentah
    raw = file.stream.read()

    # Coba beberapa encoding umum
    encodings = ["utf-8", "latin-1", "windows-1252"]
    for enc in encodings:
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            text = None
    if text is None:
        flash("Gagal membaca file CSV: encoding tidak dikenali.", "danger")
        return redirect(url_for('metal_master'))

    # Parsing CSV
    stream = io.StringIO(text, newline=None)
    reader = csv.reader(stream)

    # Siapkan koneksi database
    conn = get_db_connection()
    cur = conn.cursor()

    imported_rows = 0
    for i, row in enumerate(reader):
        # Deteksi header: skip jika baris pertama berisi teks
        if i == 0:
            if "line" in row[0].lower() or "no" in row[1].lower():
                continue  # skip header

        # Validasi minimal jumlah kolom
        if len(row) < 5:
            flash(f"Baris {i+1} tidak valid: kurang kolom.", "warning")
            continue

        # Masukkan ke database
        cur.execute("""
            INSERT INTO mesin_master (line, no_mesin, type_mesin, type_md, standard)
            VALUES (%s, %s, %s, %s, %s)
        """, row[:5])
        imported_rows += 1

    conn.commit()
    conn.close()

    flash(f"Import selesai. {imported_rows} baris berhasil dimasukkan.", "success")
    return redirect(url_for('metal_master'))

@app.route('/metal/export/excel/<string:bulan>/<string:line>')
def export_metal_excel(bulan, line):
    conn = get_db_connection()
    cur = conn.cursor()
    query = """
        SELECT tanggal, line, no_mesin, sensitivitas_actual, note
        FROM metal_log
        WHERE TO_CHAR(tanggal, 'YYYY-MM') = %s
    """
    params = [bulan]
    if line != "all":
        query += " AND line = %s"
        params.append(line)

    cur.execute(query, params)
    rows = cur.fetchall()
    cur.close()
    conn.close()

    df = pd.DataFrame(rows, columns=["Tanggal & Jam", "Line", "No. Mesin", "Sensitivitas Actual", "Note"])
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Report Metal Detector")

    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name=f"metal_report_{bulan}_{line}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route('/metal/export/pdf/<string:bulan>/<string:line>')
def export_metal_pdf(bulan, line):
    conn = get_db_connection()
    cur = conn.cursor()
    query = """
        SELECT tanggal, line, no_mesin, sensitivitas_actual, note
        FROM metal_log
        WHERE TO_CHAR(tanggal, 'YYYY-MM') = %s
    """
    params = [bulan]
    if line != "all":
        query += " AND line = %s"
        params.append(line)

    cur.execute(query, params)
    rows = cur.fetchall()
    cur.close()
    conn.close()

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [Paragraph(
        f"Laporan Metal Detector - {bulan} ({'Semua Line' if line=='all' else 'Line ' + line})",
        styles['Title']
    )]

    data = [["Tanggal & Jam", "Line", "No. Mesin", "Sensitivitas Actual", "Note"]]
    for r in rows:
        data.append(list(r))

    table = Table(data, repeatRows=1, hAlign="CENTER")
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.darkblue),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("BACKGROUND", (0,1), (-1,-1), colors.whitesmoke)
    ]))
    elements.append(table)
    doc.build(elements)

    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"metal_report_{bulan}_{line}.pdf",
        mimetype="application/pdf"
    )

@app.route('/metal/export/rekap/excel/<int:tahun>')
def export_metal_rekap_excel(tahun):
    conn = get_db_connection()
    cur = conn.cursor()

    # Definisi bulan sesuai format to_char (Mon/YY)
    months = ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Aug","Sep","Okt","Nov","Des"]
    selects = []
    for m in months:
        selects.append(
            f"MAX(CASE WHEN to_char(ml.tanggal,'Mon/YY')='{m}/{str(tahun)[-2:]}' "
            f"THEN ml.sensitivitas_actual END) AS \"{m}/{str(tahun)[-2:]}\""
        )
    select_clause = ",\n    ".join(selects)

    query = f"""
        SELECT 
            mm.no_mesin AS "NO MESIN",
            mm.line AS "LINE",
            mm.type_mesin AS "TYPE MESIN",
            mm.type_md AS "TYPE MD",
            mm.standard AS "STANDARD (RANGE)",
            {select_clause},
            '' AS "KETERANGAN"
        FROM mesin_master mm
        LEFT JOIN metal_log ml 
          ON mm.line=ml.line AND mm.no_mesin=ml.no_mesin 
         AND EXTRACT(YEAR FROM ml.tanggal)=%s
        GROUP BY mm.no_mesin, mm.line, mm.type_mesin, mm.type_md, mm.standard
        ORDER BY mm.line, mm.no_mesin;
    """
    cur.execute(query, (tahun,))
    rows = cur.fetchall()
    cols = [desc[0] for desc in cur.description]
    cur.close()
    conn.close()

    df = pd.DataFrame(rows, columns=cols)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=f"Rekap {tahun}")
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name=f"rekap_metal_{tahun}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route('/metal/export/rekap/pdf/<int:tahun>')
def export_metal_rekap_pdf(tahun):
    conn = get_db_connection()
    cur = conn.cursor()

    months = ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Aug","Sep","Okt","Nov","Des"]
    selects = []
    for m in months:
        selects.append(
            f"MAX(CASE WHEN to_char(ml.tanggal,'Mon/YY')='{m}/{str(tahun)[-2:]}' "
            f"THEN ml.sensitivitas_actual END) AS \"{m}/{str(tahun)[-2:]}\""
        )
    select_clause = ",\n    ".join(selects)

    query = f"""
        SELECT 
            mm.no_mesin,
            mm.line,
            mm.type_mesin,
            mm.type_md,
            mm.standard,
            {select_clause},
            '' AS keterangan
        FROM mesin_master mm
        LEFT JOIN metal_log ml 
          ON mm.line=ml.line AND mm.no_mesin=ml.no_mesin 
         AND EXTRACT(YEAR FROM ml.tanggal)=%s
        GROUP BY mm.no_mesin, mm.line, mm.type_mesin, mm.type_md, mm.standard
        ORDER BY mm.line, mm.no_mesin;
    """
    cur.execute(query, (tahun,))
    rows = cur.fetchall()
    cols = [desc[0] for desc in cur.description]
    cur.close()
    conn.close()

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
    styles = getSampleStyleSheet()
    elements = []
    elements.append(Paragraph(f"Rekap Metal Detector Tahun {tahun}", styles['Title']))

    # Buat tabel PDF
    data = [cols] + rows
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.darkblue),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
        ("FONTSIZE", (0,0), (-1,-1), 6),
        ("BACKGROUND", (0,1), (-1,-1), colors.whitesmoke)
    ]))
    elements.append(table)
    doc.build(elements)

    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"rekap_metal_{tahun}.pdf",
        mimetype="application/pdf"
    )



# CRUD Metal Master (tidak diubah)
@app.route("/metal_master")
def metal_master():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT id,line,no_mesin,type_mesin,type_md,standard FROM mesin_master ORDER BY line,no_mesin")
    rows = [dict(id=r[0], line=r[1], no_mesin=r[2], type_mesin=r[3], type_md=r[4], standard=r[5]) for r in cur.fetchall()]
    conn.close()
    return render_template("metal_master.html", mesin=rows)


@app.route("/metal_master/add", methods=["POST"])
def add_metal_master():
    data = (request.form['line'], request.form['no_mesin'], request.form['type_mesin'], request.form['type_md'], request.form['standard'])
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("INSERT INTO mesin_master (line,no_mesin,type_mesin,type_md,standard) VALUES (%s,%s,%s,%s,%s)", data)
    conn.commit()
    conn.close()
    return redirect(url_for('metal_master'))


@app.route("/metal_master/edit/<int:id>", methods=["POST"])
def edit_metal_master(id):
    line = request.form['line']
    no_mesin = request.form['no_mesin']
    type_mesin = request.form['type_mesin']
    type_md = request.form['type_md']
    standard = request.form['standard']
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        UPDATE mesin_master 
        SET line=%s, no_mesin=%s, type_mesin=%s, type_md=%s, standard=%s
        WHERE id=%s
    """, (line, no_mesin, type_mesin, type_md, standard, id))
    conn.commit()
    conn.close()
    flash("Data mesin berhasil diupdate", "success")
    return redirect(url_for('metal_master'))


@app.route("/metal_master/delete/<int:id>")
def delete_metal_master(id):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("DELETE FROM mesin_master WHERE id=%s", (id,))
    conn.commit()
    conn.close()
    flash("Data mesin berhasil dihapus", "warning")
    return redirect(url_for('metal_master'))


# Edit & Delete Metal Log
@app.route("/metal/edit/<int:id>", methods=["POST"])
def edit_metal_log(id):
    now = datetime.now()
    line = request.form["line"]
    no_mesin = request.form["no_mesin"]
    sensitivitas_actual = request.form["sensitivitas_actual"]
    note = request.form.get("note", "")
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        UPDATE metal_log
        SET tanggal = %s,
            line = %s,
            no_mesin = %s,
            sensitivitas_actual = %s,
            note = %s
        WHERE id = %s
    """, (now, line, no_mesin, sensitivitas_actual, note, id))
    conn.commit()
    conn.close()
    flash("Data berhasil diupdate dengan jam terbaru", "success")
    return redirect(url_for("metal_report"))


@app.route("/metal/delete/<int:id>", methods=["POST"])
def delete_metal_log(id):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("DELETE FROM metal_log WHERE id = %s", (id,))
    conn.commit()
    conn.close()
    flash("Data berhasil dihapus", "success")
    return redirect(url_for("metal_report"))

@app.route("/api/mesin")
def api_mesin():
    line = request.args.get("line")
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT no_mesin,type_mesin,type_md,standard FROM mesin_master WHERE line=%s ORDER BY no_mesin", (line,))
    rows = cur.fetchall()
    conn.close()
    return jsonify([
        {"no_mesin": r[0], "type_mesin": r[1], "type_md": r[2], "standard": r[3]}
        for r in rows
    ])

@app.context_processor
def inject_metal_status():
    conn = get_db_connection()
    cur = conn.cursor()

    # Hitung total mesin dari tabel master
    cur.execute("SELECT COUNT(*) FROM mesin_master")
    total_mesin = cur.fetchone()[0]

    # Hitung mesin yang sudah diinput bulan ini (DISTINCT per mesin)
    cur.execute("""
        SELECT COUNT(DISTINCT no_mesin) 
        FROM metal_log 
        WHERE DATE_TRUNC('month', tanggal) = DATE_TRUNC('month', CURRENT_DATE)
    """)
    sudah_input = cur.fetchone()[0]

    belum_input = total_mesin - sudah_input

    conn.close()
    return dict(
        total_mesin=total_mesin,
        sudah_input=sudah_input,
        belum_input=belum_input
    )






@app.route('/pc32')
def pc32():
    status = get_all_status_by_page('pc32')
    return render_template('PC32.html', status=status, now=datetime.now())

@app.route('/pc14')
def pc14():
    status = get_all_status_by_page('pc14')
    return render_template('PC14.html', status=status, now=datetime.now())

@app.route('/TS')
def TS():
    status = get_all_status_by_page('TS')
    return render_template('TS.html', status=status, now=datetime.now())

# ==================== RUN SERVER ====================
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True)
