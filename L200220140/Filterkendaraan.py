import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xlrd

class ExcelHandler:
    def __init__(self, file_path):
        self.file_path = file_path
        self.workbook = xlrd.open_workbook(file_path)
        self.worksheet = self.workbook.sheet_by_index(0)
        
    def get_worksheet(self):
        return self.worksheet
    
    def show_size(self):
        """Menampilkan Jumlah baris dan jumlah kolom pada file excel"""
        num_rows = self.worksheet.nrows
        num_cols = self.worksheet.ncols
        return f"Jumlah baris: {num_rows}, Jumlah Kolom: {num_cols}"
    

#tkinter GUI
root = tk.Tk()

root.geometry("1200x700") # mengatur dimensi
root.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0) # makes the root window fixed in size.

# Frame for TreeView
frame1 = tk.LabelFrame(root, text="Excel Data")
frame1.place(height=250, width=1200)

# Frame for open file dialog
file_frame = tk.LabelFrame(root, text="Open File")
file_frame.place(height=100, width=400, rely=0.65, relx=0.30)

# Buttons
button1 = tk.Button(file_frame, text="Browse A File", command=lambda: File_dialog())
button1.place(rely=0.65, relx=0.25)

button2 = tk.Button(file_frame, text="Cari Kendaraan", command=lambda: Load_excel_data())
button2.place(rely=0.65, relx=0.00)

# Button to display the number of rows and columns
button3 = tk.Button(file_frame, text="Tampilkan Jumlah Kolom dan baris", command=lambda: show_size())
button3.place(rely=0.65, relx=0.50)

# The file/file path text
label_file = ttk.Label(file_frame, text="No File Selected")
label_file.place(rely=0, relx=0)

# Label to display the number of rows and columns
label_size = ttk.Label(file_frame, text="")
label_size.place(rely=0.45, relx=0)

## Treeview Widget
tv1 = ttk.Treeview(frame1)
tv1.place(relheight=1, relwidth=1) # set the height and width of the widget to 100% of its container (frame1).

treescrolly = tk.Scrollbar(frame1, orient="vertical", command=tv1.yview) # command means update the yaxis view of the widget
treescrollx = tk.Scrollbar(frame1, orient="horizontal", command=tv1.xview) # command means update the xaxis view of the widget
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set) # assign
 # assign the scrollbars to the Treeview Widget
treescrollx.pack(side="bottom", fill="x") # make the scrollbar fill the x axis of the Treeview widget
treescrolly.pack(side="right", fill="y") # make the scrollbar fill the y axis of the Treeview widget

def show_size():
    """Displays the number of rows and columns in the Excel file"""
    file_path = label_file["text"]
    excel_handler = ExcelHandler(file_path)
    label_size["text"] = excel_handler.show_size()
    return None

def File_dialog():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xls files", "*.xls"),("All Files", "*.*")))
    label_file["text"] = filename
    return None

desired_list = ['Yamaha','Suzuki', 'Toyota', 'Honda', 'Mitsubishi', 'Nissan', 'Mazda', 'Daihatsu', 'Hyundai', 'Ford', 'Chevrolet', 'Mercedes-Benz', 'BMW', 'Audi', 'Volkswagen', 'Peugeot', 'Renault', 'Citroen', 'Opel', 'Fiat', 'Skoda', 'Seat', 'Isuzu', 'Merkur', 'Porsche', 'Lamborghini', 'Ferarri', 'Bugatti', 'Bentley', 'Rolls-Royce']

def Load_excel_data():
    """If the file selected is valid this will load the file into the Treeview"""
    file_path = label_file["text"]
    try:
        excel = ExcelHandler(file_path)
        worksheet = excel.get_worksheet()
    except ValueError:
        tk.messagebox.showerror
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None
    except Exception as e:
        tk.messagebox.showerror("Information", "File not compatible, please choose the correct file")
        return None
    clear_data()
    tv1["column"] = [worksheet.cell_value(0,col) for col in range(worksheet.ncols)]
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column) # let the column heading = column nameS
    for rowx in range(1, worksheet.nrows):
        row_data = [worksheet.cell_value(rowx,col) for col in range(worksheet.ncols)]
        if any(val in desired_list for val in row_data):
            tv1.insert("", "end", values=row_data)


def clear_data():
    tv1.delete(*tv1.get_children())
    return None


root.mainloop()
