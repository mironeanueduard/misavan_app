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

    transport_rows = db.relationship("TransportData", backref="import_record", lazy=True)


class TransportData(db.Model):
    __tablename__ = "transport_data"

    id = db.Column(db.Integer, primary_key=True)
    import_id = db.Column(db.Integer, db.ForeignKey("imports.id"), nullable=False)

    data = db.Column(db.String(50), nullable=True)
    filiala = db.Column(db.String(100), nullable=True)
    ruta = db.Column(db.String(100), nullable=True)
    km = db.Column(db.Float, nullable=True)
    nr_documente = db.Column(db.Integer, nullable=True)
    valoare_ron = db.Column(db.Float, nullable=True)
    cost_ron = db.Column(db.Float, nullable=True)