from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()


class ImportHistory(db.Model):
    __tablename__ = "imports"

    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    sheet_name = db.Column(db.String(255), nullable=True)
    imported_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    total_rows = db.Column(db.Integer, default=0, nullable=False)
    status = db.Column(db.String(50), nullable=False)


class VehicleExpense(db.Model):
    __tablename__ = "vehicle_expenses"

    id = db.Column(db.Integer, primary_key=True)
    import_id = db.Column(db.Integer, db.ForeignKey("imports.id"), nullable=False)

    numar = db.Column(db.String(100), nullable=True)
    marca = db.Column(db.String(100), nullable=True)
    model = db.Column(db.String(100), nullable=True)
    tip = db.Column(db.String(100), nullable=True)
    departament = db.Column(db.String(100), nullable=True)
    locatie = db.Column(db.String(100), nullable=True)
    centru_cost = db.Column(db.String(100), nullable=True)
    entitate = db.Column(db.String(100), nullable=True)
    status = db.Column(db.String(100), nullable=True)
    sofer = db.Column(db.String(150), nullable=True)

    revizii = db.Column(db.Float, default=0)
    reparatii = db.Column(db.Float, default=0)
    carburant = db.Column(db.Float, default=0)
    anvelope = db.Column(db.Float, default=0)
    acumulatori = db.Column(db.Float, default=0)
    accident = db.Column(db.Float, default=0)
    amenzi = db.Column(db.Float, default=0)
    alte_cheltuieli = db.Column(db.Float, default=0)
    retineri = db.Column(db.Float, default=0)
    rate = db.Column(db.Float, default=0)
    amortizari = db.Column(db.Float, default=0)

    total_reparatii = db.Column(db.Float, default=0)
    total_taxe = db.Column(db.Float, default=0)
    total_general = db.Column(db.Float, default=0)