import os
from Combine import ConcatTable
from CreareTabelMediu import ConvToMediu
from ConvertToSmallTable import ConvToSmall
dir = "C:/FloeTare2"
excel_file = os.path.join(dir, "Book1.xlsx")
excel_file2 = os.path.join(dir, "Book2.xlsx")
#CONCATENARE TABELELOR
#ConcatTable(excel_file, excel_file2)

#TRANSFORMARE DIN TABEL MARE IN TABEL MEDIU
ConvToMediu(excel_file)

#TRANSFRMARE DIN TABEL MEDIU IN TABEL MIC
#ConvToSmall(excel_file)


