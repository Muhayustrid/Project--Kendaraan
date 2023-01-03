import xlrd
open = xlrd.open_workbook("dataMagetan.xls") #membuka file
opensheet = open.sheet_by_index(0) #membuka sheet pertama
cell = opensheet.cell_value(0, 0) #mengambil nilai dari cell
semuabaris = opensheet.row_values(0)#mengambil semua nilai dalam satu baris sekaligus
semuakolom = opensheet.col_values(0)#mengambil semua nilai dalam satu kolom sekaligus
kolom = opensheet.ncols #jumlahkolom
baris = opensheet.nrows #jumlahbaris

print(opensheet.cell_value(0, 0))
print(opensheet.row_values(7))
print(opensheet.col_values(0))

for i in range(baris):
    if i == 0:
        continue
    print(tuple(opensheet.row_values(i)))