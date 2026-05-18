from __future__ import annotations
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash
from . import db


class User(db.Model):
    __tablename__ = "users"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    username = db.Column(db.String(80), unique=True, nullable=False, index=True)
    full_name = db.Column(db.String(150))
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), default="viewer", nullable=False, index=True)
    is_active = db.Column(db.Boolean, default=True, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    pt_accesses = db.relationship(
        "UserPtAccess",
        back_populates="user",
        cascade="all, delete-orphan",
        lazy="joined",
    )

    def set_password(self, password: str) -> None:
        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
        return check_password_hash(self.password_hash, password)

    def is_admin(self) -> bool:
        return self.role in ("admin", "master")

    def is_master(self) -> bool:
        return self.role == "master"

    def display_name(self) -> str:
        return (self.full_name or self.username or self.email or "User").strip()


class UserPtAccess(db.Model):
    __tablename__ = "user_pt_accesses"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    pt_name = db.Column(db.String(150), nullable=False, index=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    user = db.relationship("User", back_populates="pt_accesses")

    __table_args__ = (
        db.UniqueConstraint("user_id", "pt_name", name="uq_user_pt_accesses_user_pt"),
    )

class Company(db.Model):
    __tablename__ = "companies"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    name = db.Column(db.String(150), unique=True, nullable=False, index=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        nullable=False,
    )

    vehicles_asset_owner = db.relationship(
        "Vehicle",
        back_populates="asset_owner_company",
        foreign_keys="Vehicle.asset_owner_company_id",
        lazy="dynamic",
    )
    vehicles_pt_pemakai = db.relationship(
        "Vehicle",
        back_populates="pt_pemakai_company",
        foreign_keys="Vehicle.pt_pemakai_company_id",
        lazy="dynamic",
    )

    def __repr__(self):
        return f"<Company {self.name}>"


class Vehicle(db.Model):
    __tablename__ = "vehicles"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    no = db.Column(db.Integer, index=True)

    pt = db.Column(db.String(150), index=True)
    asset_owner_company_id = db.Column(db.Integer, db.ForeignKey("companies.id"))
    pt_pemakai_company_id = db.Column(db.Integer, db.ForeignKey("companies.id"))

    active_name = db.Column(db.String(255))
    name_as_asset_pt = db.Column(db.String(255))
    new_asset_name = db.Column(db.String(255))
    merk = db.Column(db.String(120))
    type = db.Column(db.String(120))
    jenis = db.Column(db.String(120))
    plate_old = db.Column(db.String(50), index=True)
    plate_new = db.Column(db.String(50), index=True)
    year_of_use = db.Column(db.Integer)
    user_old = db.Column(db.String(150))
    user_new = db.Column(db.String(150))
    status = db.Column(db.String(200), index=True)
    kondisi_terkini = db.Column(db.String(255))
    lokasi = db.Column(db.String(255))
    tambahan_keterangan = db.Column(db.Text)
    is_deleted = db.Column(db.Boolean, default=False, nullable=False, index=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        nullable=False,
    )

    asset_owner_company = db.relationship(
        "Company",
        back_populates="vehicles_asset_owner",
        foreign_keys=[asset_owner_company_id],
    )
    pt_pemakai_company = db.relationship(
        "Company",
        back_populates="vehicles_pt_pemakai",
        foreign_keys=[pt_pemakai_company_id],
    )

    annual_taxes = db.relationship(
        "AnnualTaxPayment",
        back_populates="vehicle",
        cascade="all, delete-orphan",
    )
    five_year_taxes = db.relationship(
        "FiveYearTaxPayment",
        back_populates="vehicle",
        cascade="all, delete-orphan",
    )
    kir_histories = db.relationship(
        "KirHistory",
        back_populates="vehicle",
        cascade="all, delete-orphan",
    )
    service_histories = db.relationship(
        "ServiceHistory",
        back_populates="vehicle",
        cascade="all, delete-orphan",
    )
    user_histories = db.relationship(
        "UserHistory",
        back_populates="vehicle",
        cascade="all, delete-orphan",
    )
    plate_histories = db.relationship(
        "PlateHistory",
        back_populates="vehicle",
        cascade="all, delete-orphan",
    )
    loan_histories_legacy = db.relationship(
        "VehicleLoanHistory",
        back_populates="vehicle",
        cascade="all, delete-orphan",
    )

    vehicle_change_histories = db.relationship(
        "VehicleChangeHistory",
        back_populates="vehicle",
        cascade="all, delete-orphan",
    )
    loan_transactions = db.relationship(
        "LoanTransaction",
        back_populates="vehicle",
        cascade="all, delete-orphan",
    )
    loan_history_entries = db.relationship(
        "LoanHistory",
        back_populates="vehicle",
        cascade="all, delete-orphan",
    )

    def current_plate(self) -> str:
        return (self.plate_new or "").strip() or (self.plate_old or "").strip()

    def display_name(self) -> str:
        return (
            (self.new_asset_name or self.active_name or self.name_as_asset_pt or "").strip()
            or f"{(self.merk or '').strip()} {(self.type or '').strip()}".strip()
            or "Kendaraan"
        )

    @property
    def kondisi(self):
        return self.kondisi_terkini

    @kondisi.setter
    def kondisi(self, v):
        self.kondisi_terkini = v

    @property
    def keterangan(self):
        return self.tambahan_keterangan

    @keterangan.setter
    def keterangan(self, v):
        self.tambahan_keterangan = v

    @property
    def user_company(self):
        return self.pt_pemakai_company

    @user_company.setter
    def user_company(self, v):
        self.pt_pemakai_company = v

    @property
    def user_company_id(self):
        return self.pt_pemakai_company_id

    @user_company_id.setter
    def user_company_id(self, v):
        self.pt_pemakai_company_id = v


class AnnualTaxPayment(db.Model):
    __tablename__ = "annual_tax_payments"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey("vehicles.id"), nullable=False, index=True)
    due_date = db.Column(db.Date, nullable=False, index=True)
    paid_date = db.Column(db.Date)
    amount = db.Column(db.Numeric(18, 2))
    note = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        nullable=False,
    )

    vehicle = db.relationship("Vehicle", back_populates="annual_taxes")


class FiveYearTaxPayment(db.Model):
    __tablename__ = "five_year_tax_payments"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey("vehicles.id"), nullable=False, index=True)
    due_date = db.Column(db.Date, nullable=False, index=True)
    paid_date = db.Column(db.Date)
    amount = db.Column(db.Numeric(18, 2))
    plate_before = db.Column(db.String(50))
    plate_after = db.Column(db.String(50))
    note = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        nullable=False,
    )

    vehicle = db.relationship("Vehicle", back_populates="five_year_taxes")


class KirHistory(db.Model):
    __tablename__ = "kir_histories"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey("vehicles.id"), nullable=False, index=True)
    kir_number = db.Column(db.String(100))
    done_date = db.Column(db.Date)           # Tgl KIR selesai/stempel — basis next KIR
    pay_date  = db.Column(db.Date)           # Tgl bayar administrasi (boleh beda dari done_date)
    valid_date = db.Column(db.Date)
    due_date = db.Column(db.Date, index=True)
    test_location = db.Column(db.String(255))
    result = db.Column(db.String(120))
    status = db.Column(db.String(120))
    note = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        nullable=False,
    )

    vehicle = db.relationship("Vehicle", back_populates="kir_histories")

    @property
    def notes(self):
        return self.note

    @notes.setter
    def notes(self, v):
        self.note = v


class ServiceHistory(db.Model):
    __tablename__ = "service_histories"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey("vehicles.id"), nullable=False, index=True)
    service_date = db.Column(db.Date, nullable=False, index=True)
    vendor = db.Column(db.String(150))
    service_type = db.Column(db.String(150))
    odometer_km = db.Column(db.Integer)
    cost = db.Column(db.Numeric(18, 2))
    note = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        nullable=False,
    )

    vehicle = db.relationship("Vehicle", back_populates="service_histories")


class UserHistory(db.Model):
    __tablename__ = "user_histories"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey("vehicles.id"), nullable=False, index=True)

    user_lama = db.Column(db.String(150))
    user_baru = db.Column(db.String(150))
    changed_at = db.Column(db.DateTime, default=datetime.utcnow)
    note = db.Column(db.Text)
    vehicle = db.relationship("Vehicle", back_populates="user_histories")

class PlateHistory(db.Model):
    __tablename__ = "plate_histories"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey("vehicles.id"), nullable=False, index=True)
    plate_old = db.Column(db.String(50))
    plate_new = db.Column(db.String(50))
    change_date = db.Column(db.DateTime, default=datetime.utcnow)
    note = db.Column(db.Text)

    vehicle = db.relationship("Vehicle", back_populates="plate_histories")


class VehicleLoanHistory(db.Model):
    __tablename__ = "vehicle_loan_histories"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey("vehicles.id"), nullable=False, index=True)
    loan_date = db.Column(db.Date)
    return_date = db.Column(db.Date)
    borrower_name = db.Column(db.String(150))
    purpose = db.Column(db.String(255))
    status = db.Column(db.String(50))
    note = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    vehicle = db.relationship("Vehicle", back_populates="loan_histories_legacy")


class VehicleChangeHistory(db.Model):
    __tablename__ = "vehicle_change_histories"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey("vehicles.id"), nullable=False, index=True)
    field_name = db.Column(db.String(100))
    field_label = db.Column(db.String(150))
    old_value = db.Column(db.Text)
    new_value = db.Column(db.Text)
    change_type = db.Column(db.String(50), default="update", index=True)
    changed_by = db.Column(db.String(100))
    changed_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    note = db.Column(db.Text)

    vehicle = db.relationship("Vehicle", back_populates="vehicle_change_histories")


class LoanTransaction(db.Model):
    __tablename__ = "loan_transactions"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey("vehicles.id"), nullable=False, index=True)
    borrower_name = db.Column(db.String(150), nullable=False)
    borrower_company = db.Column(db.String(150))
    date_out = db.Column(db.Date, nullable=False, index=True)
    date_return_plan = db.Column(db.Date)
    date_return_actual = db.Column(db.Date)
    purpose = db.Column(db.String(255))
    status = db.Column(db.String(50), default="active", index=True)
    note = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

    vehicle = db.relationship("Vehicle", back_populates="loan_transactions")
    history_entries = db.relationship(
        "LoanHistory",
        back_populates="transaction",
        cascade="all, delete-orphan",
    )


class LoanHistory(db.Model):
    __tablename__ = "loan_histories"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    loan_id = db.Column(db.Integer, db.ForeignKey("loan_transactions.id"), nullable=False, index=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey("vehicles.id"), nullable=False, index=True)
    
    action_type = db.Column(db.String(50))
    old_status = db.Column(db.String(50))
    new_status = db.Column(db.String(50))
    old_borrower = db.Column(db.String(150))
    new_borrower = db.Column(db.String(150))
    
    changed_by = db.Column(db.String(100))
    changed_at = db.Column(db.DateTime, default=datetime.utcnow)
    note = db.Column(db.Text)

    transaction = db.relationship("LoanTransaction", back_populates="history_entries")
    vehicle = db.relationship("Vehicle", back_populates="loan_history_entries")


class AuditLog(db.Model):
    __tablename__ = "audit_logs"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), index=True)
    action = db.Column(db.String(100), nullable=False, index=True)
    
    # Kolom baru yang baru saja kita tambahkan di MySQL
    entity_type = db.Column(db.String(100), index=True)
    entity_id = db.Column(db.Integer, index=True)
    vehicle_id = db.Column(db.Integer, index=True)
    note = db.Column(db.Text)
    ip = db.Column(db.String(45))
    
    # Kolom lama (tetap simpan jika perlu)
    old_data = db.Column(db.Text)
    new_data = db.Column(db.Text)
    
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False, index=True)

    # Relasi agar l.actor.username di template bisa jalan
    actor = db.relationship("User", backref="audit_logs")




class VehicleGridTable(db.Model):
    __tablename__ = "vehicle_grid_tables"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    name = db.Column(db.String(150), unique=True, nullable=False)
    description = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        nullable=False,
    )

    columns = db.relationship(
        "VehicleGridColumn",
        back_populates="table",
        cascade="all, delete-orphan",
        order_by="VehicleGridColumn.position_no.asc()",
    )
    cells = db.relationship(
        "VehicleGridCellValue",
        back_populates="table",
        cascade="all, delete-orphan",
    )


class VehicleGridColumn(db.Model):
    __tablename__ = "vehicle_grid_columns"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    table_id = db.Column(
        db.Integer,
        db.ForeignKey("vehicle_grid_tables.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    key_name = db.Column(db.String(120), nullable=False)
    label = db.Column(db.String(150), nullable=False)
    data_type = db.Column(db.String(30), default="text", nullable=False)
    position_no = db.Column(db.Integer, default=0, nullable=False)
    is_active = db.Column(db.Boolean, default=True, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        nullable=False,
    )

    table = db.relationship("VehicleGridTable", back_populates="columns")
    cells = db.relationship(
        "VehicleGridCellValue",
        back_populates="column",
        cascade="all, delete-orphan",
    )

    __table_args__ = (
        db.UniqueConstraint("table_id", "key_name", name="uq_vehicle_grid_columns_table_key"),
    )


class VehicleGridCellValue(db.Model):
    __tablename__ = "vehicle_grid_cell_values"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    table_id = db.Column(
        db.Integer,
        db.ForeignKey("vehicle_grid_tables.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    vehicle_id = db.Column(
        db.Integer,
        db.ForeignKey("vehicles.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    column_id = db.Column(
        db.Integer,
        db.ForeignKey("vehicle_grid_columns.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    value_text = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        nullable=False,
    )

    table = db.relationship("VehicleGridTable", back_populates="cells")
    column = db.relationship("VehicleGridColumn", back_populates="cells")
    vehicle = db.relationship("Vehicle")

    __table_args__ = (
        db.UniqueConstraint("table_id", "vehicle_id", "column_id", name="uq_vehicle_grid_cell_values"),
    )


# Compatibility aliases for routes.py
KirRecord = KirHistory
ServiceRecord = ServiceHistory