import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from collections import defaultdict


#=============================
# Tìm các khoảng giờ mưa liên tục
#=============================
def tim_khoang_mua(ds_gio):
    if not ds_gio:
        return ""

    ds_gio = sorted(ds_gio)

    start = ds_gio[0]
    end = ds_gio[0]

    ket_qua = []

    for g in ds_gio[1:]:
        if g == end + 1:
            end = g
        else:
            if start == end:
                ket_qua.append(f"{start}h")
            else:
                ket_qua.append(f"{start}h đến {end}h")
            start = end = g

    if start == end:
        ket_qua.append(f"{start}h")
    else:
        ket_qua.append(f"{start}h đến {end}h")

    return "; ".join(ket_qua)


#=============================
# Đọc 1 file txt
#=============================
def xu_ly_file(file_path):

    data_ngay = defaultdict(lambda: {
        "tram": "",
        "nam": "",
        "thang": "",
        "gio_mua": [],
        "tong_mua": 0
    })

    bat_dau = False

    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:

        for line in f:

            line = line.strip().replace('"', '')

            if "CORR_HOURLY_GRAPH_DATA" in line:
                bat_dau = True
                continue

            if "DAILY_GRAPH_DATA" in line:
                break

            if not bat_dau:
                continue

            cot = line.split(";")

            if len(cot) < 12:
                continue

            try:
                tram = cot[0].strip()
                nam = int(cot[1])
                thang = int(cot[2])
                ngay = int(cot[3])
                gio = int(cot[4])

                # cột 11 = lượng vũ ký
                mua = float(cot[11].replace(",", "."))

            except:
                continue

            key = (tram, nam, thang, ngay)

            data_ngay[key]["tram"] = tram
            data_ngay[key]["nam"] = nam
            data_ngay[key]["thang"] = thang

            if mua > 0:
                data_ngay[key]["gio_mua"].append(gio)
                data_ngay[key]["tong_mua"] += mua

    rows = []

    for key, value in data_ngay.items():

        tram, nam, thang, ngay = key

        rows.append([
            "",
            tram,
            nam,
            thang,
            ngay,
            tim_khoang_mua(value["gio_mua"]),
            round(value["tong_mua"], 1)
        ])

    return rows


#=============================
# Chọn nhiều file txt
#=============================
def chon_file():

    files = filedialog.askopenfilenames(
        title="Chọn các file txt",
        filetypes=[("Text files", "*.txt")]
    )

    if not files:
        return

    tat_ca = []

    stt = 1

    for file in files:

        rows = xu_ly_file(file)

        for r in rows:
            r[0] = stt
            tat_ca.append(r)
            stt += 1

    df = pd.DataFrame(
        tat_ca,
        columns=[
            "Stt",
            "Trạm",
            "Năm",
            "Tháng",
            "Ngày",
            "Mưa từ...đến",
            "Tổng lượng mưa"
        ]
    )

    output = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel file", "*.xlsx")]
    )

    if output:
        df.to_excel(output, index=False)
        messagebox.showinfo(
            "Thông báo",
            "Đã xuất Excel thành công!"
        )


#=============================
# Form giao diện
#=============================
root = tk.Tk()
root.title("Tổng hợp lượng mưa từ nhiều file TXT")
root.geometry("500x200")

lbl = tk.Label(
    root,
    text="Tổng hợp lượng mưa từ nhiều file TXT",
    font=("Arial", 14, "bold")
)
lbl.pack(pady=20)

btn = tk.Button(
    root,
    text="Chọn các file TXT và xuất Excel",
    width=35,
    height=2,
    command=chon_file
)
btn.pack()

root.mainloop()
