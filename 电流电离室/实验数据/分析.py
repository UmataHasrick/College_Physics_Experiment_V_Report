from encodings import utf_8
import xlrd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path

#-------------------------------------definition of working path and open the excel document-------------------------------------#

working_path = str(Path(__file__).parent)

excel = xlrd.open_workbook(working_path+"\\data.xlsx",encoding_override = "utf-8")
sheet = (excel.sheets())[0]

output_txtfile = open(working_path+"output.txt", mode='w',encoding='utf-8')
output_figfile = working_path+"\\Figure\\"

color_list = ["purple", "red", "blue", "black"]
symbol_list = [".", "s", "^", "*"]

#-------------------------------------drawing the figure for the Saturation characteristic curve-------------------------------------#


plt.figure

for j in range(4):
    column = sheet.col(2*j)[5:30]
    for i in range(len(column)):
        column[i] = column[i].value

    V = np.array(column, dtype=np.float64)

    column = sheet.col(2*j+1)[5:30]
    for i in range(len(column)):
        column[i] = column[i].value

    n = np.array(column, dtype=np.float64)

    plt.plot(V,n,color = color_list[j], marker = symbol_list[j], label = str(2.5-0.5*j) + "MPa")

plt.title("Saturation characteristic curve")
plt.xlabel("V")
plt.ylabel("n")
plt.legend()
plt.savefig(output_figfile + "Saturation_characteristic_curve.jpg")
plt.cla()


#-------------------------------------drawing the figure for the Sensitivity characteristic curve-------------------------------------#


for j in range(4):
    column = sheet.col(2*j+10)[5:30]
    for i in range(len(column)):
        column[i] = column[i].value

    V = np.array(column, dtype=np.float64)

    column = sheet.col(2*j+11)[5:30]
    for i in range(len(column)):
        column[i] = column[i].value

    n = np.array(column, dtype=np.float64)

    plt.plot(V,n,color = color_list[j], marker = symbol_list[j], label = str(j) + "pcs")

plt.title("Sensitivity characteristic curve")
plt.xlabel("V")
plt.ylabel("n")
plt.legend()
plt.savefig(output_figfile + "Sensitivity_characteristic_curve.jpg")
plt.cla()


#-------------------------------------drawing the figure for the Axial uniformity curve-------------------------------------#


column = sheet.col(20)[5:]
for i in range(len(column)):
    column[i] = column[i].value

D = np.array(column, dtype=np.float64)

column = sheet.col(21)[5:]
for i in range(len(column)):
    column[i] = column[i].value

n = np.array(column, dtype=np.float64)

plt.plot(D,n,color = "blue")

plt.title("Axial uniformity curve")
plt.xlabel("D")
plt.ylabel("n")
plt.savefig(output_figfile + "Axial_uniformity_curve.jpg")
plt.cla()