from flask import Flask, request, send_file, jsonify, render_template, session, redirect, url_for
from flask_cors import CORS
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from pdf2docx import Converter
from docx2pdf import convert

from functools import wraps
# Import tambahan untuk fitur ringkas dokumen
from flask_sqlalchemy import SQLAlchemy
from docx import Document
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lex_rank import LexRankSummarizer
# Import tambahan untuk meningkatkan akurasi Bahasa Indonesia
from Sastrawi.Stemmer.StemmerFactory import StemmerFactory
from Sastrawi.StopWordRemover.StopWordRemoverFactory import StopWordRemoverFactory
import datetime

import os

# Beri tahu Flask di mana folder template berada (yaitu, direktori saat ini, '.')
app = Flask(__name__, template_folder='.')
CORS(app) # Mengizinkan frontend mengambil data dari backend

# Kunci rahasia untuk mengamankan sesi. Ganti dengan string acak yang kuat.
app.secret_key = 'ganti-dengan-kunci-rahasia-yang-sangat-aman'

# ==========================================
# KONFIGURASI DATABASE MYSQL
# ==========================================
# Ganti 'username', 'password', 'localhost', dan 'portal_karyawan'
# sesuai dengan konfigurasi database MySQL Anda.
app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql+pymysql://root:@localhost/portal_karyawan'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# Debugging koneksi database saat startup
with app.app_context():
    try:
        db.engine.connect()
        print("Database connected successfully!")
    except Exception as e:
        print(f"Error connecting to database: {e}")

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ==========================================
# MODEL DATABASE (Cetak Biru Tabel)
# ==========================================
class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    fullname = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='user')
    superior_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    subordinates = db.relationship('User', backref=db.backref('superior', remote_side=[id]), lazy='dynamic')

class JadwalRapat(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    judul = db.Column(db.String(200), nullable=False)
    tanggal = db.Column(db.Date, nullable=False)
    waktu = db.Column(db.Time, nullable=False)
    peserta = db.Column(db.Text, nullable=True)

class Announcement(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    content = db.Column(db.Text, nullable=False)
    category = db.Column(db.String(50), nullable=False, default='indigo')
    date = db.Column(db.String(50), nullable=False)


# ==========================================
# DECORATORS UNTUK KEAMANAN
# ==========================================
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            # Jika ini adalah permintaan API, kirim error JSON. Jika tidak, alihkan.
            if request.path.startswith('/api/'):
                return jsonify(error="Sesi telah berakhir, silakan login kembali."), 401
            return redirect(url_for('serve_login_page'))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            if request.path.startswith('/api/'):
                return jsonify(error="Sesi telah berakhir, silakan login kembali."), 401
            return redirect(url_for('serve_login_page'))
        if session.get('user_role') != 'admin':
            if request.path.startswith('/api/'):
                return jsonify(error="Akses admin diperlukan."), 403
            return "Akses Ditolak", 403
        return f(*args, **kwargs)
    return decorated_function

# API 1: Konversi PDF ke Word
@app.route('/api/pdf-to-word', methods=['POST'])
def pdf_to_word():
    if 'file' not in request.files:
        return jsonify({"error": "Tidak ada file"}), 400
    
    file = request.files['file']
    filename = secure_filename(file.filename)
    input_path = os.path.join(UPLOAD_FOLDER, filename)
    output_filename = filename.rsplit('.', 1)[0] + '.docx'
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)
    
    file.save(input_path)

    try:
        # Menggunakan engine pdf2docx untuk konversi akurat
        cv = Converter(input_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# API 2: Konversi Word ke PDF
@app.route('/api/word-to-pdf', methods=['POST'])
def word_to_pdf():
    if 'file' not in request.files:
        return jsonify({"error": "Tidak ada file"}), 400
    
    file = request.files['file']
    filename = secure_filename(file.filename)
    input_path = os.path.join(UPLOAD_FOLDER, filename)
    output_filename = filename.rsplit('.', 1)[0] + '.pdf'
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)
    
    file.save(input_path)

    try:
        # Membutuhkan MS Word terinstal di OS server (Windows/Mac)
        convert(input_path, output_path)
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# API 3: Ringkas Dokumen (dari file .docx)
@app.route('/api/summarize', methods=['POST'])
def summarize_document():
    if 'file' not in request.files:
        return jsonify({"error": "Tidak ada file yang diunggah"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "Tidak ada file yang dipilih"}), 400

    filename = secure_filename(file.filename)
    input_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(input_path)

    try:
        # 1. Ekstrak teks dari file .docx
        doc = Document(input_path)
        full_text = [para.text for para in doc.paragraphs if para.text.strip() != '']
        text_content = '\n'.join(full_text)

        if not text_content:
            return jsonify({"summary": "Dokumen tidak berisi teks untuk diringkas."})

        # 2. Lakukan peringkasan menggunakan pustaka 'sumy'
        # Menggunakan tokenizer, stemmer, dan stopword remover Bahasa Indonesia untuk akurasi yang lebih baik.
        # Pastikan data NLTK 'punkt' sudah diunduh.
        parser = PlaintextParser.from_string(text_content, Tokenizer("indonesian"))
        
        # Buat stemmer dari Sastrawi
        stemmer_factory = StemmerFactory()
        stemmer = stemmer_factory.create_stemmer()

        # Beralih ke LexRankSummarizer yang lebih robust untuk Bahasa Indonesia
        summarizer = LexRankSummarizer(stemmer.stem)
        
        # Tambahkan daftar stopword (kata umum yang diabaikan) dari Sastrawi
        stopword_factory = StopWordRemoverFactory()
        summarizer.stop_words = stopword_factory.get_stop_words()
        
        # Tentukan jumlah kalimat dalam ringkasan (contoh: 5 kalimat)
        summary_sentences = summarizer(parser.document, 5)

        if not summary_sentences:
            summary = "Gagal membuat ringkasan. Teks di dalam dokumen mungkin terlalu pendek atau tidak memiliki cukup variasi untuk diringkas."
        else:
            summary = '\n\n'.join([str(sentence) for sentence in summary_sentences])

        return jsonify({"summary": summary})
    except Exception as e:
        return jsonify({"error": f"Gagal memproses dokumen: {str(e)}"}), 500


# ==========================================
# API UNTUK OTENTIKASI PENGGUNA
# ==========================================

@app.route('/login', methods=['POST'])
def handle_login():
    email = request.form.get('email')
    password = request.form.get('password')

    user = User.query.filter_by(email=email).first()

    if user and check_password_hash(user.password_hash, password):
        # Simpan informasi pengguna ke dalam sesi
        session['user_id'] = user.id
        session['user_name'] = user.fullname
        session['user_email'] = user.email # Menambahkan email ke sesi
        session['user_role'] = user.role

        role = user.role
        if role == 'admin':
            redirect_url = url_for('serve_admin_dashboard')
        elif role == 'atasan':
            redirect_url = url_for('serve_atasan_dashboard')
        else:  # 'user'
            redirect_url = url_for('serve_user_dashboard')
        
        return jsonify({"message": "Login berhasil!", "redirect_url": redirect_url})
    
    return jsonify({"error": "Email atau kata sandi salah"}), 401

@app.route('/register', methods=['POST'])
def handle_register():
    fullname = request.form.get('fullname')
    email = request.form.get('email')
    password = request.form.get('password')

    if not all([fullname, email, password]):
        return jsonify({"error": "Semua kolom harus diisi."}), 400
    
    if User.query.filter_by(email=email).first():
        return jsonify({"error": "Email sudah terdaftar."}), 409

    # Hash password untuk keamanan sebelum disimpan
    hashed_password = generate_password_hash(password)
    new_user = User(fullname=fullname, email=email, password_hash=hashed_password, role='user')
    
    db.session.add(new_user)
    db.session.commit()
    
    return jsonify({"message": "Registrasi berhasil! Anda akan diarahkan ke halaman login.", "redirect_url": url_for('serve_login_page')})

@app.route('/api/users', methods=['GET'])
@admin_required
def get_users():
    try:
        users = User.query.all()
        users_list = []
        for user in users:
            users_list.append({
                'id': user.id,
                'fullname': user.fullname,
                'email': user.email,
                'role': user.role,
                'superior_id': user.superior_id
            })
        return jsonify(users_list)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/change-password', methods=['POST'])
@login_required
def change_password():
    user_id = session.get('user_id')
    old_password = request.form.get('old_password')
    new_password = request.form.get('new_password')
    confirm_new_password = request.form.get('confirm_new_password')

    if not all([old_password, new_password, confirm_new_password]):
        return jsonify({"error": "Semua kolom harus diisi."}), 400

    if new_password != confirm_new_password:
        return jsonify({"error": "Kata sandi baru dan konfirmasi tidak cocok."}), 400

    if len(new_password) < 6: # Contoh validasi minimal panjang password
        return jsonify({"error": "Kata sandi baru minimal 6 karakter."}), 400

    user = User.query.get(user_id)
    if not user:
        session.clear() # Clear session if user not found (shouldn't happen with login_required)
        return jsonify({"error": "Pengguna tidak ditemukan."}), 401

    if not check_password_hash(user.password_hash, old_password):
        return jsonify({"error": "Kata sandi lama salah."}), 401

    user.password_hash = generate_password_hash(new_password)
    db.session.commit()

    return jsonify({"message": "Kata sandi berhasil diubah."}), 200

@app.route('/api/user/delete/<int:user_id>', methods=['DELETE'])
@admin_required
def delete_user(user_id):
    # Admin tidak bisa menghapus akunnya sendiri
    if user_id == session.get('user_id'):
        return jsonify({"error": "Anda tidak dapat menghapus akun Anda sendiri."}), 403

    user = User.query.get(user_id)
    if not user:
        return jsonify({"error": "Pengguna tidak ditemukan."}), 404

    try:
        db.session.delete(user)
        db.session.commit()
        return jsonify({"message": "Pengguna berhasil dihapus."}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": f"Gagal menghapus pengguna: {str(e)}"}), 500

@app.route('/api/user/update/<int:user_id>', methods=['POST'])
@admin_required
def update_user(user_id):
    user = User.query.get(user_id)
    if not user:
        return jsonify({"error": "Pengguna tidak ditemukan."}), 404

    fullname = request.form.get('fullname')
    role = request.form.get('role')
    superior_id_str = request.form.get('superior_id')

    if not fullname or not role:
        return jsonify({"error": "Nama lengkap dan peran harus diisi."}), 400

    # Admin tidak bisa mengubah perannya sendiri menjadi bukan admin
    if user.id == session.get('user_id') and role != 'admin':
        return jsonify({"error": "Anda tidak dapat mengubah peran Anda sendiri dari admin."}), 403

    try:
        user.fullname = fullname
        user.role = role
        if superior_id_str and superior_id_str.isdigit():
            user.superior_id = int(superior_id_str)
        else:
            user.superior_id = None
        db.session.commit()
        return jsonify({"message": "Data pengguna berhasil diperbarui."}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": f"Gagal memperbarui pengguna: {str(e)}"}), 500


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('serve_login_page'))

# ==========================================
# ROUTES UNTUK MENYAJIKAN HALAMAN HTML
# ==========================================
@app.route('/')
def serve_login_page():
    return render_template('LOGIN.html')

@app.route('/register.html')
def serve_register_page():
    return render_template('register.html')

@app.route('/admin/dashboard.html')
@admin_required
def serve_admin_dashboard():
    return render_template('admin/dashboard.html', user_name=session.get('user_name', 'Admin'))

@app.route('/admin/kelola-pengguna.html')
@admin_required
def serve_admin_kelola_pengguna():
    return render_template('admin/kelola-pengguna.html', user_name=session.get('user_name', 'Admin'))

@app.route('/atasan/dashboard.html')
@login_required
def serve_atasan_dashboard():
    return render_template('atasan/dashboard.html', user_name=session.get('user_name', 'Atasan'))

@app.route('/atasan/tim-saya.html')
@login_required
def serve_atasan_tim_saya():
    if session.get('user_role') != 'atasan':
        return "Akses Ditolak", 403
    return render_template('atasan/tim-saya.html', user_name=session.get('user_name', 'Atasan'))

@app.route('/atasan/jadwal-rapat.html')
@login_required
def serve_atasan_jadwal_rapat():
    if session.get('user_role') != 'atasan':
        return "Akses Ditolak", 403
    return render_template('atasan/jadwal-rapat.html', user_name=session.get('user_name', 'Atasan'))

@app.route('/atasan/kalender-kegiatan.html')
@login_required
def serve_atasan_kalender_kegiatan():
    if session.get('user_role') != 'atasan':
        return "Akses Ditolak", 403
    return render_template('atasan/kalender-kegiatan.html', user_name=session.get('user_name', 'Atasan'))

@app.route('/atasan/kelola-pengumuman.html')
@login_required
def serve_atasan_kelola_pengumuman():
    if session.get('user_role') != 'atasan':
        return "Akses Ditolak", 403
    return render_template('atasan/kelola-pengumuman.html', user_name=session.get('user_name', 'Atasan'))

@app.route('/atasan/persetujuan.html')
@login_required
def serve_atasan_persetujuan():
    if session.get('user_role') != 'atasan':
        return "Akses Ditolak", 403
    return render_template('atasan/persetujuan.html', user_name=session.get('user_name', 'Atasan'))

@app.route('/USER/dashboard.html')
@login_required
def serve_user_dashboard():
    return render_template('USER/dashboard.html', user_name=session.get('user_name', 'Pengguna'))

@app.route('/USER/direktori-staf.html')
@login_required
def serve_user_direktori():
    return render_template('USER/direktori-staf.html', user_name=session.get('user_name', 'Pengguna'))

@app.route('/USER/bantuan.html')
@login_required
def serve_user_bantuan():
    return render_template('USER/bantuan.html', user_name=session.get('user_name', 'Pengguna'))

@app.route('/USER/alat-bantuan.html')
@login_required
def serve_user_alat():
    return render_template('USER/alat-bantuan.html', user_name=session.get('user_name', 'Pengguna'))

@app.route('/USER/profile.html')
@login_required
def serve_user_profile():
    return render_template('USER/profile.html', user_name=session.get('user_name', 'Pengguna'), user_email=session.get('user_email', ''))

@app.route('/USER/jadwal-rapat.html')
@login_required
def serve_user_jadwal_rapat():
    return render_template('USER/jadwal-rapat.html', user_name=session.get('user_name', 'Pengguna'))

@app.route('/USER/kalender-kegiatan.html')
@login_required
def serve_user_kalender_kegiatan():
    return render_template('USER/kalender-kegiatan.html', user_name=session.get('user_name', 'Pengguna'))

@app.route('/admin/jadwal-rapat.html')
@admin_required
def serve_admin_jadwal_rapat():
    return render_template('admin/jadwal-rapat.html', user_name=session.get('user_name', 'Admin'))

@app.route('/admin/kalender-kegiatan.html')
@admin_required
def serve_admin_kalender_kegiatan():
    return render_template('admin/kalender-kegiatan.html', user_name=session.get('user_name', 'Admin'))

@app.route('/admin/lapor-kendala.html')
@admin_required
def serve_admin_lapor_kendala():
    return render_template('admin/lapor-kendala.html', user_name=session.get('user_name', 'Admin'))

@app.route('/admin/kelola-pengumuman.html')
@admin_required
def serve_admin_kelola_pengumuman():

   return render_template('admin/kelola-pengumuman.html', user_name=session.get('user_name', 'Admin'))

@app.route('/api/pengumuman', methods=['GET'])
@login_required
def get_announcements():
    announcements = Announcement.query.order_by(Announcement.id.desc()).all()
    return jsonify([{
        'id': a.id,
        'title': a.title,
        'content': a.content,
        'category': a.category,
        'date': a.date
    } for a in announcements])

@app.route('/api/pengumuman', methods=['POST'])
@login_required
def add_announcement():
    data = request.get_json(silent=True) or {}
    title = data.get('title')
    content = data.get('content')
    category = data.get('category', 'indigo')

    if not title or not content:
        return jsonify({"error": "Judul dan konten wajib diisi."}), 400

    new_announcement = Announcement(
        title=title,
        content=content,
        category=category,
        date=datetime.datetime.now().strftime('%d %b %Y %H:%M')
    )

    try:
        db.session.add(new_announcement)
        db.session.commit()
        return jsonify({"message": "Pengumuman berhasil ditambahkan!"}), 201
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": f"Gagal menambah pengumuman: {str(e)}"}), 500

@app.route('/api/pengumuman/<int:ann_id>', methods=['DELETE'])
@login_required
def delete_announcement(ann_id):
    announcement = Announcement.query.get(ann_id)
    if not announcement:
        return jsonify({"error": "Pengumuman tidak ditemukan."}), 404

    try:
        db.session.delete(announcement)
        db.session.commit()
        return jsonify({"message": "Pengumuman berhasil dihapus!"}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": f"Gagal menghapus pengumuman: {str(e)}"}), 500

@app.route('/api/meetings', methods=['GET'])
@login_required
def get_meetings():
    meetings = JadwalRapat.query.order_by(JadwalRapat.tanggal, JadwalRapat.waktu).all()
    return jsonify([{
        'id': m.id,
        'judul': m.judul,
        'tanggal': m.tanggal.strftime('%d %b %Y'),
        'waktu': m.waktu.strftime('%H:%M'),
        'peserta': m.peserta
    } for m in meetings])

@app.route('/api/meetings', methods=['POST'])
@admin_required
def add_meeting():
    judul = request.form.get('judul_rapat')
    tanggal_str = request.form.get('tanggal_rapat')
    waktu_str = request.form.get('waktu_rapat')
    peserta = request.form.get('peserta_rapat')

    if not all([judul, tanggal_str, waktu_str]):
        return jsonify({"error": "Judul, tanggal, dan waktu harus diisi."}), 400

    try:
        tanggal = datetime.datetime.strptime(tanggal_str, '%Y-%m-%d').date()
        waktu = datetime.datetime.strptime(waktu_str, '%H:%M').time()

        new_meeting = JadwalRapat(judul=judul, tanggal=tanggal, waktu=waktu, peserta=peserta)
        db.session.add(new_meeting)
        db.session.commit()

        return jsonify({"message": "Jadwal rapat berhasil ditambahkan."}), 201
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": f"Gagal menambah jadwal: {str(e)}"}), 500

@app.route('/api/meetings/<int:meeting_id>', methods=['DELETE'])
@admin_required
def delete_meeting(meeting_id):
    meeting = JadwalRapat.query.get_or_404(meeting_id)
    try:
        db.session.delete(meeting)
        db.session.commit()
        return jsonify({"message": "Jadwal rapat berhasil dihapus."}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": f"Gagal menghapus jadwal: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)