import pandas as pd
import pymysql
from pymysql.cursors import DictCursor

DB_HOST = "127.0.0.1"
DB_PORT = 3306
DB_USER = "root"
DB_PASSWORD = "Fleet@123"
DB_NAME = "fleet_sinar_group"

EXCEL_FILE = r"exel fix.xlsx"

REQUIRED_HEADERS = [
    "PT",
    "PT pemilik aset",
    "NAMA AKTIVA",
    "Nama sesuai asset PT",
    "Nama Aset baru",
    "MERK",
    "TYPE",
    "JENIS",
    "NO POLISI LAMA",
    "NO POLISI BARU",
    "TAHUN PEMAKAIAN",
    "USER LAMA",
    "USER BARU",
    "Status",
    "PT Pemakai 1",
    "PT Pemakai 2",
    "Kondisi terkini",
    "Lokasi",
    "Tambahan Keterangan",
]


def norm(v):
    if pd.isna(v):
        return None
    s = str(v).strip()
    return s if s else None


def to_int_or_none(v):
    if pd.isna(v):
        return None
    s = str(v).strip()
    if not s:
        return None
    try:
        return int(float(s))
    except Exception:
        return None


def connect_db(db_name=None):
    return pymysql.connect(
        host=DB_HOST,
        port=DB_PORT,
        user=DB_USER,
        password=DB_PASSWORD,
        database=db_name,
        cursorclass=DictCursor,
        autocommit=False,
    )


def ensure_tables():
    conn = connect_db(DB_NAME)
    try:
        with conn.cursor() as cur:
            # companies
            cur.execute("""
                CREATE TABLE IF NOT EXISTS companies (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    name VARCHAR(150) NOT NULL UNIQUE,
                    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                )
            """)

            # vehicles
            cur.execute("""
                CREATE TABLE IF NOT EXISTS vehicles (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    no VARCHAR(50),
                    pt VARCHAR(150),
                    asset_owner_company_id INT,
                    pt_pemakai_1_company_id INT,
                    pt_pemakai_2_company_id INT,
                    active_name VARCHAR(255),
                    name_as_asset_pt VARCHAR(255),
                    new_asset_name VARCHAR(255),
                    merk VARCHAR(120),
                    type VARCHAR(120),
                    jenis VARCHAR(120),
                    plate_old VARCHAR(50),
                    plate_new VARCHAR(50),
                    year_of_use INT,
                    user_old VARCHAR(150),
                    user_new VARCHAR(150),
                    status VARCHAR(200),
                    kondisi_terkini VARCHAR(255),
                    lokasi VARCHAR(255),
                    tambahan_keterangan TEXT,
                    is_deleted TINYINT(1) NOT NULL DEFAULT 0,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                )
            """)

            # Kalau tabel sudah ada tapi schema lama, paksa rapihin kolom sistem
            try:
                cur.execute("""
                    ALTER TABLE companies
                    MODIFY created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    MODIFY updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                """)
            except Exception:
                pass

            try:
                cur.execute("""
                    ALTER TABLE vehicles
                    MODIFY is_deleted TINYINT(1) NOT NULL DEFAULT 0,
                    MODIFY created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    MODIFY updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                """)
            except Exception:
                pass

        conn.commit()
    finally:
        conn.close()


def get_company_id(cur, name):
    name = norm(name)
    if not name:
        return None

    cur.execute("SELECT id FROM companies WHERE name=%s", (name,))
    row = cur.fetchone()

    if row:
        return row["id"]

    cur.execute("""
        INSERT INTO companies (name, created_at, updated_at)
        VALUES (%s, NOW(), NOW())
    """, (name,))
    return cur.lastrowid


def normalize_headers(df):
    raw_cols = [str(c).strip() for c in df.columns]
    new_cols = []
    pt_pemakai_count = 0

    for c in raw_cols:
        c_clean = " ".join(c.lower().split())

        if c_clean == "pt pemakai" or c_clean.startswith("pt pemakai."):
            pt_pemakai_count += 1
            if pt_pemakai_count == 1:
                new_cols.append("PT Pemakai 1")
            elif pt_pemakai_count == 2:
                new_cols.append("PT Pemakai 2")
            else:
                new_cols.append(f"PT Pemakai {pt_pemakai_count}")
        else:
            new_cols.append(c)

    df.columns = new_cols
    print("HEADER TERBACA:", df.columns.tolist())
    return df


def validate_headers(df):
    missing = [h for h in REQUIRED_HEADERS if h not in df.columns]
    if missing:
        raise Exception("Header kurang: " + ", ".join(missing))


def insert_data():
    df = pd.read_excel(EXCEL_FILE, header=0)
    df = normalize_headers(df)
    validate_headers(df)

    conn = connect_db(DB_NAME)

    try:
        with conn.cursor() as cur:
            for _, row in df.iterrows():
                cur.execute("""
                    INSERT INTO vehicles (
                        no,
                        pt,
                        asset_owner_company_id,
                        pt_pemakai_1_company_id,
                        pt_pemakai_2_company_id,
                        active_name,
                        name_as_asset_pt,
                        new_asset_name,
                        merk,
                        type,
                        jenis,
                        plate_old,
                        plate_new,
                        year_of_use,
                        user_old,
                        user_new,
                        status,
                        kondisi_terkini,
                        lokasi,
                        tambahan_keterangan,
                        is_deleted,
                        created_at,
                        updated_at
                    ) VALUES (
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                        %s, NOW(), NOW()
                    )
                """, (
                    norm(row["No"]) if "No" in df.columns else None,
                    norm(row["PT"]),
                    get_company_id(cur, row["PT pemilik aset"]),
                    get_company_id(cur, row["PT Pemakai 1"]),
                    get_company_id(cur, row["PT Pemakai 2"]),
                    norm(row["NAMA AKTIVA"]),
                    norm(row["Nama sesuai asset PT"]),
                    norm(row["Nama Aset baru"]),
                    norm(row["MERK"]),
                    norm(row["TYPE"]),
                    norm(row["JENIS"]),
                    norm(row["NO POLISI LAMA"]),
                    norm(row["NO POLISI BARU"]),
                    to_int_or_none(row["TAHUN PEMAKAIAN"]),
                    norm(row["USER LAMA"]),
                    norm(row["USER BARU"]),
                    norm(row["Status"]),
                    norm(row["Kondisi terkini"]),
                    norm(row["Lokasi"]),
                    norm(row["Tambahan Keterangan"]),
                    0
                ))

        conn.commit()
        print("✅ Import berhasil!")

    except Exception as e:
        conn.rollback()
        print("❌ Error:", e)

    finally:
        conn.close()


if __name__ == "__main__":
    ensure_tables()
    insert_data()