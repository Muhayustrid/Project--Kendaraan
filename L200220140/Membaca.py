import xlrd
buka = xlrd.open_workbook("CAR.xls") #membuka file
bukasheet = buka.sheet_by_index(0) #membuka sheet pertama
cell = bukasheet.cell_value(0, 0) #mengambil nilai dari cell
semuabaris = bukasheet.row_values(0)#mengambil semua nilai dalam satu baris sekaligus
semuakolom = bukasheet.col_values(0)#mengambil semua nilai dalam satu kolom sekaligus
kolom = bukasheet.ncols #jumlahkolom
baris = bukasheet.nrows #jumlahbaris

print(bukasheet.cell_value(0, 0))
print(bukasheet.row_values(12))
print(bukasheet.col_values(0))

#for b in range(baris):
#    if b == 0:
#        continue
#    print(tuple(bukasheet.row_values(b)))
