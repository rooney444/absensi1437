from flask import Flask, render_template, request, redirect, send_file, session
import csv
from datetime import datetime, time
import os
import pandas as pd

app = Flask(__name__)
app.secret_key = "absensi_secret_key"

EMPLOYEES = {
    "MCW1": "Michael Chandra Wijaya",
    "RRS1": "Radhitya Rooney Syahputra"
}

def get_hari(date_obj):
    hari_dict = {
        0: "Senin", 1: "Selasa", 2: "Rabu", 3: "Kamis",
        4: "Jumat", 5: "Sabtu", 6: "Minggu"
    }
    return hari_dict[date_obj.weekday()]

@app.route("/", methods=["GET", "POST"])
def absen():
    message = ""
    if request.method == "POST":
        kode = request.form.get("kode")
        nama = EMPLOYEES.get(kode)
        now = datetime.now()
        batas_waktu = time(9, 15)
        if nama:
            waktu = now.strftime("%Y-%m-%d %H:%M:%S")
            hari = get_hari(now)
            status = "Tepat Waktu" if now.time() <= batas_waktu else "Terlambat"
            tanggal_format = now.strftime("%d %B %Y")
            jam_format = now.strftime("%H:%M:%S")
            session['success'] = {
                "nama": nama,
                "status": status,
                "hari": hari,
                "tanggal": tanggal_format,
                "jam": jam_format
            }
            with open("absensi.csv", "a", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                writer.writerow([nama, waktu, hari, status])
            return redirect("/")
        else:
            message = "Kode tidak dikenali."
    success_data = session.pop('success', None)
    return render_template("index.html", message=message, success=success_data)

@app.route("/data")
def lihat_data():
    data = []
    stats = {}
    if os.path.exists("absensi.csv"):
        with open("absensi.csv", "r", encoding="utf-8") as file:
            reader = csv.reader(file)
            data = sorted(list(reader), key=lambda x: (x[0], x[1]))  # Sort by name, then date
        for row in data:
            nama = row[0]
            stats[nama] = stats.get(nama, 0) + 1
    return render_template("data.html", data=data, stats=stats)

@app.route("/export")
def export_excel():
    if not os.path.exists("absensi.csv"):
        return "Belum ada data untuk diexport."
    df = pd.read_csv("absensi.csv", header=None, names=["Nama", "Waktu", "Hari", "Status"])
    df.sort_values(by=["Nama", "Waktu"], inplace=True)
    stat_df = df["Nama"].value_counts().rename_axis("Nama").reset_index(name="Jumlah Kehadiran")
    file_path = "absensi_dengan_statistik.xlsx"
    with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Absensi", index=False)
        stat_df.to_excel(writer, sheet_name="Statistik", index=False)
        worksheet2 = writer.sheets["Statistik"]
        total_kehadiran = stat_df["Jumlah Kehadiran"].sum()
        worksheet2.write(len(stat_df)+1, 0, "Total Kehadiran")
        worksheet2.write(len(stat_df)+1, 1, total_kehadiran)
        workbook = writer.book
        for sheet_name in ["Absensi", "Statistik"]:
            worksheet = writer.sheets[sheet_name]
            for i, width in enumerate([30, 25, 15, 20]):
                worksheet.set_column(i, i, width)
    return send_file(file_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
