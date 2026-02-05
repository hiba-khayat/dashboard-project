import os
from datetime import datetime

import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "dev-secret-key"  # for demo only

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# store last uploaded file path (simple simulation)
LAST_FILE_PATH = None


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def find_date_column(df: pd.DataFrame):
    # try common names first
    candidates = ["date", "Date", "DATE", "full_date", "FULL_DATE", "day", "Day", "timestamp", "Timestamp"]
    for c in candidates:
        if c in df.columns:
            return c

    # fallback: attempt to parse object columns to datetime and choose the best
    for c in df.columns:
        if df[c].dtype == "object":
            parsed = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
            if parsed.notna().mean() >= 0.6:  # 60% parsable
                return c
    return None


def find_numeric_column(df: pd.DataFrame):
    # common names first
    candidates = ["sales", "Sales", "SALES", "amount", "Amount", "AMOUNT", "revenue", "Revenue", "RATED_AMOUNT"]
    for c in candidates:
        if c in df.columns:
            # ensure numeric
            try:
                pd.to_numeric(df[c], errors="coerce")
                return c
            except Exception:
                pass

    nums = df.select_dtypes(include="number").columns.tolist()
    return nums[0] if nums else None


def find_category_column(df: pd.DataFrame):
    candidates = ["Category", "CATEGORY", "Channel", "CHANNEL", "Status", "STATUS", "Type", "TYPE", "city", "City"]
    for c in candidates:
        if c in df.columns:
            return c

    obj_cols = df.select_dtypes(include="object").columns.tolist()
    return obj_cols[0] if obj_cols else None


@app.route("/")
def index():
    return redirect(url_for("upload"))


@app.route("/upload", methods=["GET", "POST"])
def upload():
    global LAST_FILE_PATH

    if request.method == "POST":
        if "file" not in request.files:
            flash("No file part in request.")
            return redirect(request.url)

        f = request.files["file"]
        if f.filename == "":
            flash("No file selected.")
            return redirect(request.url)

        if not allowed_file(f.filename):
            flash("Only Excel files are allowed (.xlsx / .xls).")
            return redirect(request.url)

        filename = secure_filename(f.filename)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        saved_name = f"{stamp}_{filename}"
        save_path = os.path.join(UPLOAD_FOLDER, saved_name)

        f.save(save_path)
        LAST_FILE_PATH = save_path

        flash("File uploaded successfully.")
        return redirect(url_for("dashboard"))

    return render_template("upload.html")


@app.route("/dashboard")
def dashboard():
    if not LAST_FILE_PATH or not os.path.exists(LAST_FILE_PATH):
        flash("No uploaded file found. Please upload an Excel file first.")
        return redirect(url_for("upload"))

    # read excel
    try:
        df = pd.read_excel(LAST_FILE_PATH)
    except Exception as e:
        flash(f"Failed to read Excel: {e}")
        return redirect(url_for("upload"))

    # basic cleaning for preview
    df = df.copy()
    columns = df.columns.tolist()

    # KPIs
    rows, cols = df.shape
    missing_cells = int(df.isna().sum().sum())
    empty_string_cells = int((df.astype(str).apply(lambda s: s.str.strip()).eq("")).sum().sum())
    missing_total = missing_cells + empty_string_cells
    duplicate_rows = int(df.duplicated().sum())

    kpis = {
        "Rows": rows,
        "Columns": cols,
        "Missing cells": missing_total,
        "Duplicate rows": duplicate_rows,
    }

    # Preview (Top 20)
    preview = df.head(20).fillna("").to_dict(orient="records")

    # Top values (Bar chart data)
    cat_col = find_category_column(df)
    bar_labels = []
    bar_values = []
    top_values = {}

    if cat_col:
        vc = df[cat_col].astype(str).str.strip()
        vc = vc[vc != ""]
        top_values = vc.value_counts().head(5).to_dict()
        bar_labels = list(top_values.keys())
        bar_values = list(top_values.values())

    # Line chart data (date + numeric)
    date_col = find_date_column(df)
    num_col = find_numeric_column(df)

    line_labels = []
    line_values = []
    used_date_col = None
    used_num_col = None

    if date_col and num_col:
        try:
            d = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
            v = pd.to_numeric(df[num_col], errors="coerce")
            tmp = pd.DataFrame({"d": d, "v": v}).dropna()
            if not tmp.empty:
                tmp["d"] = tmp["d"].dt.date
                grouped = tmp.groupby("d")["v"].sum().sort_index().head(30)
                line_labels = [str(x) for x in grouped.index.tolist()]
                line_values = grouped.values.tolist()
                used_date_col = date_col
                used_num_col = num_col
        except Exception:
            pass

    charts = {
        "bar": {
            "labels": bar_labels,
            "values": bar_values,
            "title": f"Top values ({cat_col})" if cat_col else "Top values",
        },
        "line": {
            "labels": line_labels,
            "values": line_values,
            "title": f"Trend: {used_num_col} by {used_date_col}" if used_date_col and used_num_col else "Trend",
        },
    }

    return render_template(
        "dashboard.html",
        kpis=kpis,
        columns=columns,
        preview=preview,
        top_values=top_values,
        top_col=cat_col,
        charts=charts,
        file_name=os.path.basename(LAST_FILE_PATH),
    )


if __name__ == "__main__":
    app.run(debug=True)
