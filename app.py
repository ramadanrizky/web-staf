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

# ===================== TAMBAHAN: AUTO-DOWNLOAD DATA NLTK =====================
# Sumy (peringkasan dokumen & notulensi) butuh data tokenizer NLTK untuk
# memecah teks jadi kalimat. Blok ini otomatis mengunduh data tersebut kalau
# belum ada, supaya tidak perlu jalankan perintah manual tiap pindah komputer
# atau deploy ke server baru.
import nltk

def _pastikan_data_nltk_tersedia():
    paket_dibutuhkan = ['punkt_tab', 'punkt']
    for paket in paket_dibutuhkan:
        try:
            nltk.data.find(f'tokenizers/{paket}')
        except LookupError:
            print(f"[NLTK] Data '{paket}' belum ada, mengunduh otomatis...")
            try:
                nltk.download(paket, quiet=True)
                print(f"[NLTK] Berhasil mengunduh '{paket}'.")
            except Exception as e:
                print(f"[PERINGATAN] Gagal mengunduh data NLTK '{paket}': {e}")
                print("[PERINGATAN] Fitur Ringkas Dokumen & Notulensi Rapat mungkin tidak berfungsi.")

_pastikan_data_nltk_tersedia()
# ===============================================================================

# ===================== TAMBAHAN: NOTULENSI VIDEO RAPAT =====================
# Import untuk transkripsi audio/video rapat menjadi teks
import speech_recognition as sr
from pydub import AudioSegment
import math
import tempfile

# Path LANGSUNG ke ffmpeg.exe (Windows). Ini menghindari masalah PATH yang
# sering gagal. Kalau nanti pindah komputer atau deploy ke server Linux,
# ganti baris _FFMPEG_PATH ini sesuai lokasi ffmpeg di server tersebut
# (atau kosongkan jadi None agar otomatis pakai "ffmpeg" dari PATH sistem Linux).
_FFMPEG_PATH = r"C:\Users\MyBook Pro Max Army\Documents\OJT\ffmpeg-9.0.1-essentials_build\bin\ffmpeg.exe"
_FFMPEG_BIN_DIR = os.path.dirname(_FFMPEG_PATH)
_FFMPEG_READY = os.path.exists(_FFMPEG_PATH)

if _FFMPEG_READY:
    # PENTING: pydub juga butuh ffprobe.exe (file pendamping ffmpeg) untuk
    # membaca info durasi/format video. pydub mencari ffprobe lewat PATH
    # sistem, bukan lewat AudioSegment.converter. Supaya tidak perlu edit
    # System PATH Windows secara manual, kita tambahkan folder bin ini ke
    # PATH milik proses Python ini saja (tidak permanen, aman, tidak
    # mengubah pengaturan Windows Anda).
    os.environ["PATH"] = _FFMPEG_BIN_DIR + os.pathsep + os.environ.get("PATH", "")

    AudioSegment.converter = _FFMPEG_PATH
    _FFPROBE_PATH = os.path.join(_FFMPEG_BIN_DIR, 'ffprobe.exe')
    if os.path.exists(_FFPROBE_PATH):
        AudioSegment.ffprobe = _FFPROBE_PATH
        print(f"[OK] ffmpeg siap dipakai: {_FFMPEG_PATH}")
        print(f"[OK] ffprobe siap dipakai: {_FFPROBE_PATH}")
    else:
        print(f"[PERINGATAN] ffprobe.exe TIDAK ditemukan di folder: {_FFMPEG_BIN_DIR}")
        print("[PERINGATAN] Cek isi folder tersebut, harus ada ffmpeg.exe, ffplay.exe, DAN ffprobe.exe.")
else:
    print(f"[PERINGATAN] File ffmpeg TIDAK ditemukan di: {_FFMPEG_PATH}")
    print("[PERINGATAN] Cek kembali apakah path di atas sudah sesuai lokasi ffmpeg.exe Anda.")
    print("[PERINGATAN] Fitur Notulensi Rapat tidak akan berfungsi sampai ffmpeg tersedia.")
# =============================================================================


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

# ===================== TAMBAHAN: BATAS UKURAN UPLOAD =====================
# File video rapat bisa besar, default Flask tidak membatasi tapi server (mis. Werkzeug dev)
# bisa lambat/timeout untuk file sangat besar. Batas ini bisa disesuaikan (contoh: 500 MB).
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024
# ===========================================================================

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

# ===================== TAMBAHAN: FOLDER SEMENTARA UNTUK AUDIO =====================
TEMP_AUDIO_FOLDER = 'temp_audio'
os.makedirs(TEMP_AUDIO_FOLDER, exist_ok=True)
# Format video/audio yang diterima untuk fitur Notulensi Rapat
ALLOWED_MEDIA_EXTENSIONS = {'mp4', 'mov', 'avi', 'mkv', 'webm', 'mp3', 'wav', 'm4a', 'aac', 'ogg'}

def is_allowed_media(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_MEDIA_EXTENSIONS
# ====================================================================================

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


# ===================== TAMBAHAN: MODEL NOTULENSI RAPAT =====================
class Notulensi(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    judul = db.Column(db.String(255), nullable=False)
    peserta = db.Column(db.Text, nullable=True)
    transkrip = db.Column(db.Text, nullable=False)
    ringkasan = db.Column(db.Text, nullable=True)
    durasi_detik = db.Column(db.Integer, nullable=True)
    tanggal = db.Column(db.String(50), nullable=False)
    dibuat_oleh = db.Column(db.String(100), nullable=True)
# =============================================================================


# ===================== TAMBAHAN: MODEL TUGAS SAYA =====================
class Tugas(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)  # pemilik/pengerja tugas
    assigned_by = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)  # atasan pemberi tugas (kosong = tugas pribadi)
    judul = db.Column(db.String(255), nullable=False)
    deskripsi = db.Column(db.Text, nullable=True)
    prioritas = db.Column(db.String(20), nullable=False, default='sedang')  # rendah, sedang, tinggi
    status = db.Column(db.String(20), nullable=False, default='belum')  # belum, proses, selesai
    deadline = db.Column(db.Date, nullable=True)
    dibuat_pada = db.Column(db.DateTime, nullable=False, default=datetime.datetime.utcnow)
# =========================================================================


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

# ===================== TAMBAHAN: DECORATOR ATASAN =====================
def atasan_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            if request.path.startswith('/api/'):
                return jsonify(error="Sesi telah berakhir, silakan login kembali."), 401
            return redirect(url_for('serve_login_page'))
        if session.get('user_role') not in ('atasan', 'admin'):
            if request.path.startswith('/api/'):
                return jsonify(error="Akses atasan diperlukan."), 403
            return "Akses Ditolak", 403
        return f(*args, **kwargs)
    return decorated_function
# =========================================================================

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


# ===================== TAMBAHAN: FUNGSI RINGKASAN DIPAKAI ULANG =====================
# Fungsi ini diambil dari logika /api/summarize agar bisa dipakai ulang
# oleh fitur Notulensi Rapat (meringkas hasil transkrip suara).
def generate_summary_from_text(text_content, rasio_ringkasan=0.35, min_kalimat=2, maks_kalimat=10):
    if not text_content or not text_content.strip():
        return "Teks tidak berisi konten untuk diringkas."

    # CATATAN: NLTK tidak menyediakan model tokenizer kalimat untuk Bahasa
    # Indonesia (punkt_tab hanya mendukung beberapa bahasa seperti Inggris,
    # Jerman, Prancis, dll). Kita pakai tokenizer Inggris untuk memisah
    # kalimat (aturan tanda baca cukup mirip), sementara pemahaman kata
    # Bahasa Indonesia tetap ditangani oleh stemmer & stopword Sastrawi.
    parser = PlaintextParser.from_string(text_content, Tokenizer("english"))

    total_kalimat = len(list(parser.document.sentences))

    # Kalau teks aslinya memang sudah pendek (sedikit kalimat), tidak ada
    # gunanya "meringkas" lebih jauh lagi -- kembalikan apa adanya saja
    # daripada memaksakan hasil yang terlihat identik dengan aslinya.
    if total_kalimat <= min_kalimat:
        return text_content.strip()

    stemmer_factory = StemmerFactory()
    stemmer = stemmer_factory.create_stemmer()

    summarizer = LexRankSummarizer(stemmer.stem)

    stopword_factory = StopWordRemoverFactory()
    summarizer.stop_words = stopword_factory.get_stop_words()

    # Jumlah kalimat ringkasan menyesuaikan panjang teks asli (proporsional),
    # dibatasi antara min_kalimat dan maks_kalimat.
    jumlah_kalimat = max(min_kalimat, min(maks_kalimat, round(total_kalimat * rasio_ringkasan)))

    summary_sentences = summarizer(parser.document, jumlah_kalimat)

    if not summary_sentences:
        return "Gagal membuat ringkasan. Teks mungkin terlalu pendek atau kurang bervariasi."

    return '\n\n'.join([str(sentence) for sentence in summary_sentences])
# ======================================================================================


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

        # 2. Lakukan peringkasan menggunakan fungsi bersama (lihat generate_summary_from_text)
        summary = generate_summary_from_text(text_content)

        return jsonify({"summary": summary})
    except Exception as e:
        return jsonify({"error": f"Gagal memproses dokumen: {str(e)}"}), 500


# ===================== TAMBAHAN: FITUR NOTULENSI VIDEO RAPAT =====================

def _extract_audio_to_wav(input_path, wav_path):
    """Mengekstrak/mengonversi file video/audio apa pun menjadi WAV mono 16kHz
    agar bisa diproses oleh SpeechRecognition. Membutuhkan ffmpeg terpasang
    dan tersedia di PATH sistem."""
    audio = AudioSegment.from_file(input_path)
    audio = audio.set_channels(1).set_frame_rate(16000)
    audio.export(wav_path, format='wav')
    return audio.duration_seconds


def _transcribe_wav(wav_path, durasi_detik, bahasa='id-ID', panjang_chunk_detik=20):
    """Memecah audio panjang menjadi beberapa potongan (chunk) lalu mengirim
    tiap potongan ke Google Web Speech API untuk ditranskripsi. Membutuhkan
    koneksi internet aktif di server."""
    recognizer = sr.Recognizer()
    hasil_teks = []

    jumlah_chunk = max(1, math.ceil(durasi_detik / panjang_chunk_detik))
    audio_full = AudioSegment.from_wav(wav_path)

    for i in range(jumlah_chunk):
        start_ms = int(i * panjang_chunk_detik * 1000)
        end_ms = int(min((i + 1) * panjang_chunk_detik * 1000, durasi_detik * 1000))
        chunk = audio_full[start_ms:end_ms]

        with tempfile.NamedTemporaryFile(suffix='.wav', delete=False, dir=TEMP_AUDIO_FOLDER) as tmp_chunk:
            chunk_path = tmp_chunk.name
        chunk.export(chunk_path, format='wav')

        try:
            with sr.AudioFile(chunk_path) as source:
                audio_data = recognizer.record(source)
            teks_chunk = recognizer.recognize_google(audio_data, language=bahasa)
            # PENTING: Google Speech API mengembalikan teks TANPA tanda baca sama
            # sekali. Kalau tidak diberi tanda titik di sini, seluruh transkrip
            # akan dianggap 1 kalimat raksasa oleh sistem peringkas (sumy),
            # sehingga hasil "ringkasan" jadi sama saja dengan teks aslinya.
            # Kita tandai batas tiap potongan audio sebagai batas kalimat.
            if teks_chunk:
                hasil_teks.append(teks_chunk.strip() + '.')
        except sr.UnknownValueError:
            hasil_teks.append('[bagian ini tidak terdengar jelas].')
        except sr.RequestError as e:
            hasil_teks.append(f'[gagal menghubungi layanan transkripsi: {e}].')
        finally:
            if os.path.exists(chunk_path):
                os.remove(chunk_path)

    return ' '.join(hasil_teks).strip()


@app.route('/api/notulensi', methods=['POST'])
@login_required
def buat_notulensi():
    if not _FFMPEG_READY:
        return jsonify({"error": f"ffmpeg belum siap di server. Path yang dicek: {_FFMPEG_PATH}. Pastikan file ffmpeg.exe benar-benar ada di path tersebut, lalu restart server."}), 500

    if 'file' not in request.files:
        return jsonify({"error": "Tidak ada file video/audio yang diunggah."}), 400

    file = request.files['file']
    judul = request.form.get('judul', '').strip()
    peserta = request.form.get('peserta', '').strip()

    if file.filename == '' or not judul:
        return jsonify({"error": "Judul rapat dan file wajib diisi."}), 400

    if not is_allowed_media(file.filename):
        return jsonify({"error": "Format file tidak didukung. Gunakan mp4, mov, mkv, mp3, wav, m4a, dll."}), 400

    filename = secure_filename(file.filename)
    input_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(input_path)

    wav_path = os.path.join(TEMP_AUDIO_FOLDER, filename.rsplit('.', 1)[0] + '_convert.wav')

    try:
        # 1. Ekstrak audio dari video/audio menjadi WAV
        durasi_detik = _extract_audio_to_wav(input_path, wav_path)

        # 2. Transkripsi audio menjadi teks
        transkrip = _transcribe_wav(wav_path, durasi_detik)

        if not transkrip:
            return jsonify({"error": "Tidak ada suara yang berhasil dikenali dari file ini."}), 422

        # 3. Ringkas transkrip menjadi poin-poin notulensi
        ringkasan = generate_summary_from_text(transkrip)

        # 4. Simpan ke database
        notulensi_baru = Notulensi(
            judul=judul,
            peserta=peserta,
            transkrip=transkrip,
            ringkasan=ringkasan,
            durasi_detik=int(durasi_detik),
            tanggal=datetime.datetime.now().strftime('%d %b %Y %H:%M'),
            dibuat_oleh=session.get('user_name', 'Tidak diketahui')
        )
        db.session.add(notulensi_baru)
        db.session.commit()

        return jsonify({
            "message": "Notulensi berhasil dibuat!",
            "id": notulensi_baru.id,
            "transkrip": transkrip,
            "ringkasan": ringkasan,
            "durasi_detik": int(durasi_detik)
        }), 201

    except FileNotFoundError as e:
        # Biasanya terjadi jika ffmpeg tidak ditemukan atau file rusak
        return jsonify({"error": f"Gagal memproses media (ffmpeg tidak ditemukan): {str(e)}. Cek kembali path ffmpeg.exe di app.py sudah benar, lalu restart server."}), 500
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": f"Gagal membuat notulensi: {str(e)}"}), 500
    finally:
        # Bersihkan file sementara
        if os.path.exists(input_path):
            os.remove(input_path)
        if os.path.exists(wav_path):
            os.remove(wav_path)


@app.route('/api/notulensi', methods=['GET'])
@login_required
def get_notulensi_list():
    daftar = Notulensi.query.order_by(Notulensi.id.desc()).all()
    return jsonify([{
        'id': n.id,
        'judul': n.judul,
        'peserta': n.peserta,
        'ringkasan': n.ringkasan,
        'durasi_detik': n.durasi_detik,
        'tanggal': n.tanggal,
        'dibuat_oleh': n.dibuat_oleh
    } for n in daftar])


@app.route('/api/notulensi/<int:notulensi_id>', methods=['GET'])
@login_required
def get_notulensi_detail(notulensi_id):
    n = Notulensi.query.get(notulensi_id)
    if not n:
        return jsonify({"error": "Notulensi tidak ditemukan."}), 404
    return jsonify({
        'id': n.id,
        'judul': n.judul,
        'peserta': n.peserta,
        'transkrip': n.transkrip,
        'ringkasan': n.ringkasan,
        'durasi_detik': n.durasi_detik,
        'tanggal': n.tanggal,
        'dibuat_oleh': n.dibuat_oleh
    })


@app.route('/api/notulensi/<int:notulensi_id>', methods=['DELETE'])
@login_required
def delete_notulensi(notulensi_id):
    n = Notulensi.query.get(notulensi_id)
    if not n:
        return jsonify({"error": "Notulensi tidak ditemukan."}), 404
    try:
        db.session.delete(n)
        db.session.commit()
        return jsonify({"message": "Notulensi berhasil dihapus."}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": f"Gagal menghapus notulensi: {str(e)}"}), 500


@app.route('/USER/notulensi-rapat.html')
@login_required
def serve_user_notulensi():
    return render_template('USER/notulensi-rapat.html', user_name=session.get('user_name', 'Pengguna'))

# ====================================================================================


# ===================== TAMBAHAN: FITUR TUGAS SAYA =====================

@app.route('/api/tugas', methods=['GET'])
@login_required
def get_tugas_list():
    user_id = session.get('user_id')
    status_filter = request.args.get('status', 'semua')

    query = Tugas.query.filter_by(user_id=user_id)
    if status_filter and status_filter != 'semua':
        query = query.filter_by(status=status_filter)

    daftar = query.order_by(Tugas.status.asc(), Tugas.deadline.asc(), Tugas.id.desc()).all()

    hasil = []
    for t in daftar:
        pemberi = User.query.get(t.assigned_by) if t.assigned_by else None
        hasil.append({
            'id': t.id,
            'judul': t.judul,
            'deskripsi': t.deskripsi,
            'prioritas': t.prioritas,
            'status': t.status,
            'deadline': t.deadline.strftime('%Y-%m-%d') if t.deadline else None,
            'deadline_tampil': t.deadline.strftime('%d %b %Y') if t.deadline else None,
            'dibuat_pada': t.dibuat_pada.strftime('%d %b %Y %H:%M') if t.dibuat_pada else None,
            'diberikan_oleh': pemberi.fullname if pemberi else None,
            # Tugas yang diberikan atasan hanya boleh diubah statusnya oleh staf,
            # tidak boleh diedit detail/dihapus (itu wewenang atasan pemberi tugas).
            'bisa_diedit': t.assigned_by is None,
            'bisa_dihapus': t.assigned_by is None
        })
    return jsonify(hasil)


@app.route('/api/tugas', methods=['POST'])
@login_required
def add_tugas():
    data = request.get_json(silent=True) or request.form

    judul = (data.get('judul') or '').strip()
    deskripsi = (data.get('deskripsi') or '').strip()
    prioritas = data.get('prioritas', 'sedang')
    deadline_str = data.get('deadline')

    if not judul:
        return jsonify({"error": "Judul tugas wajib diisi."}), 400

    if prioritas not in ('rendah', 'sedang', 'tinggi'):
        prioritas = 'sedang'

    deadline = None
    if deadline_str:
        try:
            deadline = datetime.datetime.strptime(deadline_str, '%Y-%m-%d').date()
        except ValueError:
            return jsonify({"error": "Format tanggal deadline tidak valid."}), 400

    try:
        # Dibuat lewat halaman "Tugas Saya" sendiri = tugas pribadi (assigned_by kosong)
        tugas_baru = Tugas(
            user_id=session.get('user_id'),
            assigned_by=None,
            judul=judul,
            deskripsi=deskripsi,
            prioritas=prioritas,
            status='belum',
            deadline=deadline
        )
        db.session.add(tugas_baru)
        db.session.commit()
        return jsonify({"message": "Tugas berhasil ditambahkan.", "id": tugas_baru.id}), 201
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": f"Gagal menambahkan tugas: {str(e)}"}), 500


@app.route('/api/tugas/<int:tugas_id>', methods=['PUT'])
@login_required
def update_tugas(tugas_id):
    tugas = Tugas.query.get(tugas_id)
    if not tugas:
        return jsonify({"error": "Tugas tidak ditemukan."}), 404
    if tugas.user_id != session.get('user_id'):
        return jsonify({"error": "Anda tidak memiliki akses ke tugas ini."}), 403

    data = request.get_json(silent=True) or request.form

    # Kalau tugas ini diberikan oleh atasan, staf yang mengerjakan HANYA boleh
    # mengubah status (progres), tidak boleh mengubah judul/deskripsi/dll.
    hanya_boleh_ubah_status = tugas.assigned_by is not None

    try:
        if not hanya_boleh_ubah_status:
            if 'judul' in data and data.get('judul', '').strip():
                tugas.judul = data.get('judul').strip()
            if 'deskripsi' in data:
                tugas.deskripsi = data.get('deskripsi', '').strip()
            if 'prioritas' in data and data.get('prioritas') in ('rendah', 'sedang', 'tinggi'):
                tugas.prioritas = data.get('prioritas')
            if 'deadline' in data:
                deadline_str = data.get('deadline')
                if deadline_str:
                    tugas.deadline = datetime.datetime.strptime(deadline_str, '%Y-%m-%d').date()
                else:
                    tugas.deadline = None

        if 'status' in data and data.get('status') in ('belum', 'proses', 'selesai'):
            tugas.status = data.get('status')

        db.session.commit()
        return jsonify({"message": "Tugas berhasil diperbarui."}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": f"Gagal memperbarui tugas: {str(e)}"}), 500


@app.route('/api/tugas/<int:tugas_id>', methods=['DELETE'])
@login_required
def delete_tugas(tugas_id):
    tugas = Tugas.query.get(tugas_id)
    if not tugas:
        return jsonify({"error": "Tugas tidak ditemukan."}), 404
    if tugas.user_id != session.get('user_id'):
        return jsonify({"error": "Anda tidak memiliki akses ke tugas ini."}), 403
    if tugas.assigned_by is not None:
        return jsonify({"error": "Tugas ini diberikan oleh atasan, hanya atasan yang bisa menghapusnya."}), 403

    try:
        db.session.delete(tugas)
        db.session.commit()
        return jsonify({"message": "Tugas berhasil dihapus."}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": f"Gagal menghapus tugas: {str(e)}"}), 500


@app.route('/USER/tugas-saya.html')
@login_required
def serve_user_tugas():
    return render_template('USER/tugas-saya.html', user_name=session.get('user_name', 'Pengguna'))

# =========================================================================


# ===================== TAMBAHAN: ATASAN MEMBERI TUGAS KE STAF =====================

@app.route('/api/atasan/staf', methods=['GET'])
@atasan_required
def get_staf_tim():
    """Daftar staf yang berada di bawah atasan yang sedang login (untuk dropdown pemilihan)."""
    atasan_id = session.get('user_id')
    if session.get('user_role') == 'admin':
        # Admin boleh lihat semua user biasa untuk keperluan pengujian/administrasi
        staf = User.query.filter(User.role == 'user').all()
    else:
        staf = User.query.filter_by(superior_id=atasan_id).all()

    return jsonify([{'id': s.id, 'fullname': s.fullname, 'email': s.email} for s in staf])


@app.route('/api/atasan/tugas', methods=['GET'])
@atasan_required
def get_tugas_yang_diberikan():
    """Daftar semua tugas yang sudah diberikan oleh atasan ini ke stafnya, untuk memantau progres."""
    atasan_id = session.get('user_id')
    daftar = Tugas.query.filter_by(assigned_by=atasan_id).order_by(Tugas.status.asc(), Tugas.id.desc()).all()

    hasil = []
    for t in daftar:
        staf = User.query.get(t.user_id)
        hasil.append({
            'id': t.id,
            'judul': t.judul,
            'deskripsi': t.deskripsi,
            'prioritas': t.prioritas,
            'status': t.status,
            'deadline_tampil': t.deadline.strftime('%d %b %Y') if t.deadline else None,
            'diberikan_kepada': staf.fullname if staf else 'Tidak diketahui',
            'diberikan_kepada_id': t.user_id
        })
    return jsonify(hasil)


@app.route('/api/atasan/tugas', methods=['POST'])
@atasan_required
def beri_tugas():
    data = request.get_json(silent=True) or request.form

    target_user_id = data.get('user_id')
    judul = (data.get('judul') or '').strip()
    deskripsi = (data.get('deskripsi') or '').strip()
    prioritas = data.get('prioritas', 'sedang')
    deadline_str = data.get('deadline')

    if not target_user_id or not judul:
        return jsonify({"error": "Staf tujuan dan judul tugas wajib diisi."}), 400

    try:
        target_user_id = int(target_user_id)
    except (ValueError, TypeError):
        return jsonify({"error": "ID staf tidak valid."}), 400

    target_user = User.query.get(target_user_id)
    if not target_user:
        return jsonify({"error": "Staf tujuan tidak ditemukan."}), 404

    # Validasi: atasan (bukan admin) hanya boleh memberi tugas ke stafnya sendiri
    if session.get('user_role') == 'atasan' and target_user.superior_id != session.get('user_id'):
        return jsonify({"error": "Anda hanya bisa memberi tugas ke staf di tim Anda sendiri."}), 403

    if prioritas not in ('rendah', 'sedang', 'tinggi'):
        prioritas = 'sedang'

    deadline = None
    if deadline_str:
        try:
            deadline = datetime.datetime.strptime(deadline_str, '%Y-%m-%d').date()
        except ValueError:
            return jsonify({"error": "Format tanggal deadline tidak valid."}), 400

    try:
        tugas_baru = Tugas(
            user_id=target_user_id,
            assigned_by=session.get('user_id'),
            judul=judul,
            deskripsi=deskripsi,
            prioritas=prioritas,
            status='belum',
            deadline=deadline
        )
        db.session.add(tugas_baru)
        db.session.commit()
        return jsonify({"message": f"Tugas berhasil diberikan ke {target_user.fullname}."}), 201
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": f"Gagal memberi tugas: {str(e)}"}), 500


@app.route('/api/atasan/tugas/<int:tugas_id>', methods=['DELETE'])
@atasan_required
def batalkan_tugas(tugas_id):
    tugas = Tugas.query.get(tugas_id)
    if not tugas:
        return jsonify({"error": "Tugas tidak ditemukan."}), 404
    if session.get('user_role') != 'admin' and tugas.assigned_by != session.get('user_id'):
        return jsonify({"error": "Anda tidak memiliki akses ke tugas ini."}), 403

    try:
        db.session.delete(tugas)
        db.session.commit()
        return jsonify({"message": "Tugas berhasil dibatalkan."}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": f"Gagal membatalkan tugas: {str(e)}"}), 500


@app.route('/atasan/beri-tugas.html')
@login_required
def serve_atasan_beri_tugas():
    if session.get('user_role') != 'atasan':
        return "Akses Ditolak", 403
    return render_template('atasan/beri-tugas.html', user_name=session.get('user_name', 'Atasan'))

# =====================================================================================



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
