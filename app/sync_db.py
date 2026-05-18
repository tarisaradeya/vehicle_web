from flask import Flask
from app.models import db
import urllib.parse

app = Flask(__name__)

# Password Anda mengandung karakter '@', jadi kita perlu meng-encode-nya
password = urllib.parse.quote_plus("fleet@123")
# URI Database dengan password yang sudah di-encode
app.config['SQLALCHEMY_DATABASE_URI'] = f'mysql+pymysql://fleet_app:{password}@localhost/fleet_sinar_group'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db.init_app(app)

with app.app_context():
    print("--- Memulai Sinkronisasi Database ---")
    
    # 1. Pastikan tabel dasar ada
    try:
        db.create_all()
        print("[OK] Tabel dasar sudah diperiksa/dibuat.")
    except Exception as e:
        print(f"[ERROR] Gagal koneksi/create_all: {e}")

    # 2. Perbaiki kolom date_return_actual di loan_transactions
    try:
        db.session.execute(db.text("ALTER TABLE loan_transactions CHANGE COLUMN actual_return_date date_return_actual DATE"))
        print("[OK] Kolom actual_return_date diubah menjadi date_return_actual.")
    except Exception:
        try:
            db.session.execute(db.text("ALTER TABLE loan_transactions ADD COLUMN date_return_actual DATE AFTER date_return_plan"))
            print("[OK] Kolom date_return_actual ditambahkan.")
        except Exception:
            print("[INFO] Kolom date_return_actual sepertinya sudah ada.")

    # 3. Tambahkan kolom borrower_company
    try:
        db.session.execute(db.text("ALTER TABLE loan_transactions ADD COLUMN borrower_company VARCHAR(150) AFTER borrower_name"))
        print("[OK] Kolom borrower_company ditambahkan.")
    except Exception:
        print("[INFO] Kolom borrower_company sepertinya sudah ada.")

    # 4. Tambahkan kolom-kolom di loan_histories
    history_cols = ["action_type VARCHAR(50)", "old_status VARCHAR(50)", "new_status VARCHAR(50)", "old_borrower VARCHAR(150)", "new_borrower VARCHAR(150)"]
    for col in history_cols:
        try:
            db.session.execute(db.text(f"ALTER TABLE loan_histories ADD COLUMN {col}"))
            print(f"[OK] Kolom {col.split()[0]} ditambahkan ke loan_histories.")
        except Exception:
            pass

    db.session.commit()
    print("--- Sinkronisasi Selesai! Silakan restart Flask Anda ---")
