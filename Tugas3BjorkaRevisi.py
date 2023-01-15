import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xlrd

class Kendaraan:
    merk_list = ['Yamaha','Suzuki', 'Toyota', 'Honda', 'Mitsubishi', 'Nissan','Kia', 'Mazda', 'Daihatsu', 'Hyundai', 'Ford', 'Chevrolet', 'Mercedes-Benz',
                 'BMW', 'Audi', 'Volkswagen', 'Peugeot', 'Renault', 'Citroen', 'Opel', 'Fiat', 'Skoda', 'Seat', 'Isuzu', 'Merkur', 'Porsche', 'Lamborghini',
                 'Ferarri', 'Bugatti', 'Bentley', 'Rolls-Royce']

    def __init__(self,  merk):
        self.set_merk(merk)

    def set_merk(self, merk):
        if merk in self.merk_list:
            self.merk = merk
        else:
            raise ValueError("Invalid merk. Please choose from: {}".format(", ".join(self.merk_list)))



root = tk.Tk()

root.geometry("1200x800") # mengatur dimensi
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
tv1.place(relheight=0.9, relwidth=1)
 #atur tinggi dan lebar widget menjadi 100% dari wadahnya (frame1)
# Adding Columns
tv1["columns"] = ("jenis", "merk", "tahun_pembuatan", "warna", "nomor_polisi", "nama_pemilik")

tv1.heading("#0", text="No", anchor= 'c')
tv1.heading("jenis", text="Jenis", anchor='c')
tv1.heading("merk", text="Merk", anchor='c')
tv1.heading("tahun_pembuatan", text="Tahun Pembuatan", anchor='c')
tv1.heading("warna", text="Warna", anchor='c')
tv1.heading("nomor_polisi", text="Nomor Polisi", anchor='c')
tv1.heading("nama_pemilik", text="Nama Pemilik", anchor='c')



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
        workbook = xlrd.open_workbook(file_path)
        sheet = workbook.sheet_by_index(0)
        tv1.delete(*tv1.get_children())
        for i in range(sheet.nrows): #perulangan for dijalankan untuk setiap baris pada sheet yang telah dibaca
            if sheet.cell_value(i, 1) in Kendaraan.merk_list: #akan dicek apakah nilai dari kolom 1 (index 0-based) pada baris tersebut ada dalam list merk_list
                tv1.insert("", "end", values=(                #Jika ada, maka baris tersebut akan ditambahkan ke dalam Treeview widget dengan menggunakan method insert
                sheet.cell_value(i, 0), sheet.cell_value(i, 1), sheet.cell_value(i, 2), sheet.cell_value(i, 3),
                sheet.cell_value(i, 4), sheet.cell_value(i, 5)))
        label_file.configure(text=file_path)

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