from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
import os

from config import Config
from models import db, ImportHistory, TransportData

app = Flask(__name__)
app.config.from_object(Config)

db.init_app(app)

ALLOWED_EXTENSIONS = {"xlsx"}

REQUIRED_COLUMNS = [
    "Data",
    "Filiala",
    "Ruta",
    "Km",
    "Nr_Documente",
    "Valoare_RON",
    "Cost_RON"
]


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def normalize_cell(cell_value):
    if cell_value is None:
        return ""
    return str(cell_value).strip()


def find_header_row(rows, required_columns, max_scan_rows=10):
    best_index = None
    best_match_count = 0
    best_headers = []

    scan_limit = min(len(rows), max_scan_rows)

    for i in range(scan_limit):
        current_row = [normalize_cell(cell) for cell in rows[i]]
        match_count = sum(1 for col in required_columns if col in current_row)

        if match_count > best_match_count:
            best_match_count = match_count
            best_index = i
            best_headers = current_row

    return best_index, best_headers, best_match_count

def safe_float(value):
    if value is None or str(value).strip() == "":
        return None
    try:
        return float(value)
    except ValueError:
        return None


def safe_int(value):
    if value is None or str(value).strip() == "":
        return None
    try:
        return int(float(value))
    except ValueError:
        return None


@app.route("/")
def home():
    records = TransportData.query.all()

    total_inregistrari = len(records)
    total_km = sum(row.km or 0 for row in records)
    total_valoare = sum(row.valoare_ron or 0 for row in records)
    total_cost = sum(row.cost_ron or 0 for row in records)
    profit = total_valoare - total_cost

    rentabilitate = 0
    if total_valoare > 0:
        rentabilitate = (profit / total_valoare) * 100

    # agregare pentru grafic
    filiale_totals = {}
    for row in records:
        filiala = row.filiala if row.filiala else "Necunoscut"
        filiale_totals[filiala] = filiale_totals.get(filiala, 0) + (row.valoare_ron or 0)

    chart_labels = list(filiale_totals.keys())
    chart_values = [round(value, 2) for value in filiale_totals.values()]

    # sumar pe filiale
    filiale_summary_map = {}

    for row in records:
        filiala = row.filiala if row.filiala else "Necunoscut"

        if filiala not in filiale_summary_map:
            filiale_summary_map[filiala] = {
                "filiala": filiala,
                "total_km": 0,
                "total_valoare": 0,
                "total_cost": 0,
            }

        filiale_summary_map[filiala]["total_km"] += row.km or 0
        filiale_summary_map[filiala]["total_valoare"] += row.valoare_ron or 0
        filiale_summary_map[filiala]["total_cost"] += row.cost_ron or 0

    filiale_summary = []
    for item in filiale_summary_map.values():
        item["profit"] = item["total_valoare"] - item["total_cost"]
        filiale_summary.append(item)

    filiale_summary.sort(key=lambda x: x["total_valoare"], reverse=True)

    return render_template(
        "index.html",
        total_inregistrari=total_inregistrari,
        total_km=round(total_km, 2),
        total_valoare=round(total_valoare, 2),
        total_cost=round(total_cost, 2),
        profit=round(profit, 2),
        rentabilitate=round(rentabilitate, 2),
        chart_labels=chart_labels,
        chart_values=chart_values,  
        filiale_summary=filiale_summary,
    )   

@app.route("/import", methods=["GET", "POST"])
def import_excel():
    message = None
    message_type = None
    preview_headers = []
    preview_rows = []
    uploaded_filename = None
    sheet_name = None
    missing_columns = []
    detected_columns = []
    total_data_rows = 0
    header_row_index = None
    recent_imports = ImportHistory.query.order_by(ImportHistory.id.desc()).limit(10).all()
    if request.method == "POST":
        if "excel_file" not in request.files:
            message = "Nu a fost trimis niciun fișier."
            message_type = "error"
            recent_imports = ImportHistory.query.order_by(ImportHistory.id.desc()).limit(10).all()
            return render_template(
            "import_excel.html",
            message=message,
            message_type=message_type,
            preview_headers=preview_headers,
            preview_rows=preview_rows,
            uploaded_filename=uploaded_filename,
            sheet_name=sheet_name,
            missing_columns=missing_columns,
            detected_columns=detected_columns,
            required_columns=REQUIRED_COLUMNS,
            total_data_rows=total_data_rows,
            header_row_index=header_row_index,
            recent_imports=recent_imports,
    )

        file = request.files["excel_file"]

        if file.filename == "":
            message = "Te rog selectează un fișier."
            message_type = "error"
            return render_template(
                "import_excel.html",
                message=message,
                message_type=message_type,
                preview_headers=preview_headers,
                preview_rows=preview_rows,
                uploaded_filename=uploaded_filename,
                sheet_name=sheet_name,
                missing_columns=missing_columns,
                detected_columns=detected_columns,
                required_columns=REQUIRED_COLUMNS,
                total_data_rows=total_data_rows,
                header_row_index=header_row_index,
            )

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            save_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(save_path)

            try:
                workbook = load_workbook(save_path, data_only=True)
                worksheet = workbook.active

                uploaded_filename = filename
                sheet_name = worksheet.title

                rows = list(worksheet.iter_rows(values_only=True))

                if not rows:
                    message = "Fișierul Excel este gol."
                    message_type = "error"
                else:
                    header_row_index, detected_headers, match_count = find_header_row(rows, REQUIRED_COLUMNS)

                    if header_row_index is None or match_count == 0:
                        message = "Nu am putut identifica automat header-ul din fișier."
                        message_type = "error"
                    else:
                        preview_headers = detected_headers
                        detected_columns = [col for col in detected_headers if col]

                        missing_columns = [
                            col for col in REQUIRED_COLUMNS if col not in detected_columns
                        ]

                        data_rows = rows[header_row_index + 1:]
                        non_empty_data_rows = [
                            row for row in data_rows
                            if any(cell is not None and str(cell).strip() != "" for cell in row)
                        ]
                        total_data_rows = len(non_empty_data_rows)

                        max_preview_rows = 20
                        for row in non_empty_data_rows[:max_preview_rows]:
                            preview_rows.append([
                                "" if cell is None else str(cell) for cell in row
                            ])

                        if missing_columns:
                            status = "INVALID"
                            message = "Header-ul a fost identificat, dar lipsesc unele coloane obligatorii."
                            message_type = "error"
                        else:
                            status = "SUCCESS"
                            message = f'Fișierul "{filename}" a fost încărcat, validat și salvat cu succes.'
                            message_type = "success"

                        import_record = ImportHistory(
                            filename=filename,
                            sheet_name=sheet_name,
                            total_rows=total_data_rows,
                            status=status
                        )
                        db.session.add(import_record)
                        db.session.commit()

                        if not missing_columns:
                            column_indexes = {
                                col_name: detected_headers.index(col_name)
                                for col_name in REQUIRED_COLUMNS
                            }

                            for row in non_empty_data_rows:
                                transport_entry = TransportData(
                                    import_id=import_record.id,
                                    data=str(row[column_indexes["Data"]]) if row[column_indexes["Data"]] is not None else None,
                                    filiala=str(row[column_indexes["Filiala"]]) if row[column_indexes["Filiala"]] is not None else None,
                                    ruta=str(row[column_indexes["Ruta"]]) if row[column_indexes["Ruta"]] is not None else None,
                                    km=safe_float(row[column_indexes["Km"]]),
                                    nr_documente=safe_int(row[column_indexes["Nr_Documente"]]),
                                    valoare_ron=safe_float(row[column_indexes["Valoare_RON"]]),
                                    cost_ron=safe_float(row[column_indexes["Cost_RON"]]),
                                )
                                db.session.add(transport_entry)

                            db.session.commit()

                # dacă vrem, mai târziu salvăm aici și rândurile în transport_data

            except Exception as e:
                message = f"A apărut o eroare la citirea fișierului: {str(e)}"
                message_type = "error"
        else:
            message = "Sunt acceptate doar fișiere .xlsx"
            message_type = "error"

    return render_template(
        "import_excel.html",
        message=message,
        message_type=message_type,
        preview_headers=preview_headers,
        preview_rows=preview_rows,
        uploaded_filename=uploaded_filename,
        sheet_name=sheet_name,
        missing_columns=missing_columns,
        detected_columns=detected_columns,
        required_columns=REQUIRED_COLUMNS,
        total_data_rows=total_data_rows,
        header_row_index=header_row_index,
        recent_imports=recent_imports,  
    )


@app.route("/monitorizare")
def monitorizare():
    filiala_filter = request.args.get("filiala", "").strip()
    ruta_filter = request.args.get("ruta", "").strip()

    query = TransportData.query

    if filiala_filter:
        query = query.filter(TransportData.filiala.ilike(f"%{filiala_filter}%"))

    if ruta_filter:
        query = query.filter(TransportData.ruta.ilike(f"%{ruta_filter}%"))

    records = query.order_by(TransportData.id.desc()).limit(100).all()

    return render_template(
        "monitorizare.html",
        records=records,
        filiala_filter=filiala_filter,
        ruta_filter=ruta_filter,
    )


if __name__ == "__main__":
    with app.app_context():
        db.create_all()
    app.run(debug=True)