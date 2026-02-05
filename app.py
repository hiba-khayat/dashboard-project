import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-change-me")

def allowed_file(filename: str) -> bool:
    return filename.lower().endswith((".xlsx", ".xls"))

def load_latest_df() -> pd.DataFrame:
    latest_csv = os.path.join(UPLOAD_DIR, "latest.csv")
    if not os.path.exists(latest_csv):
        return pd.DataFrame()
    return pd.read_csv(latest_csv)

@app.route("/")
def home():
    return redirect(url_for("upload"))

@app.route("/upload", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        f = request.files.get("file")
        if not f or not f.filename.strip():
            flash("Please choose an Excel file.")
            return redirect(url_for("upload"))

        if not allowed_file(f.filename):
            flash("Only .xlsx/.xls files are allowed.")
            return redirect(url_for("upload"))

        save_path = os.path.join(UPLOAD_DIR, f.filename)
        f.save(save_path)

        df = pd.read_excel(save_path)
        df.to_csv(os.path.join(UPLOAD_DIR, "latest.csv"), index=False)

        flash("File uploaded successfully!")
        return redirect(url_for("dashboard"))

    return render_template("upload.html")

@app.route("/dashboard")
def dashboard():
    df = load_latest_df()
    if df.empty:
        flash("Upload an Excel file first.")
        return redirect(url_for("upload"))

    kpis = {
        "rows": int(df.shape[0]),
        "columns": int(df.shape[1]),
    }

    # مثال KPI لو عندك عمود Sales
    if "Sales" in df.columns:
        kpis["total_sales"] = float(df["Sales"].fillna(0).sum())
        kpis["avg_sales"] = float(df["Sales"].fillna(0).mean())

    preview = df.head(20).to_dict(orient="records")
    columns = list(df.columns)

    return render_template("dashboard.html", kpis=kpis, columns=columns, preview=preview)

if __name__ == "__main__":
    app.run(debug=True)
