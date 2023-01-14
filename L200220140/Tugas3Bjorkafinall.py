import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xlrd

class Kendaraan:
    def __init__(self, jenis, merk, tahun_pembuatan, warna, nomor_polisi, nama_pemilik):
        self.jenis = jenis
        self.merk = merk
        self.tahun_pembuatan = tahun_pembuatan
        self.warna = warna
        self.nomor_polisi = nomor_polisi
        self.nama_pemilik = nama_pemilik

    def set_nomor_polisi(self, nomor_polisi):
        self.nomor_polisi = nomor_polisi

    def set_nama_pemilik(self, nama_pemilik):
        self.nama_pemilik = nama_pemilik

    def info_kendaraan(self):
        print("Jenis kendaraan: ", self.jenis)
        print("Merk kendaraan: ", self.merk)
        print("Tahun pembuatan: ", self.tahun_pembuatan)
        print("Warna kendaraan: ", self.warna)
        print("Nomor Polisi: ", self.nomor_polisi)
        print("Nama Pemilik: ", self.nama_pemilik)


root = tk.Tk()

root.geometry("1200x700") # mengatur dimensi
root.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0) # membuat ukuran jendela root

frame1 = tk.LabelFrame(root, text="Kelompok Bjorka Squad")
frame1.place(height=350, width=1200)

file_frame = tk.LabelFrame(root, text="Buka File")
file_frame.place(height=100, width=400, rely=0.65, relx=0.30)

button1 = tk.Button(file_frame, text="Cari File", command=lambda: File_dialog())
button1.place(rely=0.65, relx=0.50)

button2 = tk.Button(file_frame, text="Tampilkan", command=lambda: Load_excel_data())
button2.place(rely=0.65, relx=0.20)

label_file = ttk.Label(file_frame, text="Tidak ada file yang dipilih")
label_file.place(rely=0, relx=0)


## Treeview Widget
tv1 = ttk.Treeview(frame1)
tv1.place(relheight=1, relwidth=1) #atur tinggi dan lebar widget menjadi 100% dari wadahnya (frame1)
# Adding Columns
tv1["columns"] = ("jenis", "merk", "tahun_pembuatan", "warna", "nomor_polisi", "nama_pemilik")
tv1.column("jenis", width=100, stretch=False)
tv1.column("merk", width=150, stretch=False)
tv1.column("tahun_pembuatan", width=150, stretch=False)
tv1.column("warna", width=150, stretch=False)
tv1.column("nomor_polisi", width=200, stretch=False)
tv1.column("nama_pemilik", width=200, stretch=False)

tv1.heading("jenis", text="Jenis", anchor='c')
tv1.heading("merk", text="Merk", anchor='c')
tv1.heading("tahun_pembuatan", text="Tahun Pembuatan", anchor='c')
tv1.heading("warna", text="Warna", anchor='c')
tv1.heading("nomor_polisi", text="Nomor Polisi", anchor='c')
tv1.heading("nama_pemilik", text="Nama Pemilik", anchor='c')



tv1.heading("jenis", text="Jenis")
tv1.heading("merk", text="Merk")
tv1.heading("tahun_pembuatan", text="Tahun Pembuatan")
tv1.heading("warna", text="Warna")
tv1.heading("nomor_polisi", text="Nomor Polisi")
tv1.heading("nama_pemilik", text="Nama Pemilik")



treescrolly = tk.Scrollbar(frame1, orient="vertical", command=tv1.yview) # memperbarui tampilan widget Yaxis
treescrollx = tk.Scrollbar(frame1, orient="horizontal", command=tv1.xview) # memperbarui tampilan widget Xaxis
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set) # scrollbar ke Treeview Widget
treescrollx.pack(side="bottom", fill="x") # buat scrollbar mengisi sumbu x dari lebar Treeview
treescrolly.pack(side="right", fill="y") # buat scrollbar mengisi sumbu y dari lebar Treeview


def File_dialog():
    """Fungsi ini akan membuka file explorer dan menetapkan jalur file yang dipilih ke label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xls files", "*.xls"),("All Files", "*.*")))
    label_file["text"] = filename
    return None


def Load_excel_data():
    """Jika file yang dipilih sesuai, maka akan memuat file ke Treeview"""
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        workbook = xlrd.open_workbook(excel_filename)
        sheet = workbook.sheet_by_index(0)
        for row_idx in range(1, sheet.nrows):
            row = sheet.row(row_idx)
            kendaraan = Kendaraan(row[0].value, row[1].value, row[2].value, row[3].value, row[4].value, row[5].value)
            tv1.insert("", "end", values=(
            kendaraan.jenis, kendaraan.merk, kendaraan.tahun_pembuatan, kendaraan.warna, kendaraan.nomor_polisi,
            kendaraan.nama_pemilik))
    except xlrd.biffh.XLRDError:
        tk.messagebox.showerror("Information", "File yang anda pilih bukan file excel, only support xls")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None


def clear_data():
    tv1.delete(*tv1.get_children())
    return None


root.mainloop()