from encodings import utf_8
import xlrd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path
from scipy import optimize

#-------------------------------------definition of working path and open the excel document-------------------------------------#

working_path = str(Path(__file__).parent)

excel = xlrd.open_workbook(working_path+"\\data.xlsx",encoding_override = "utf-8")
sheet = (excel.sheets())[0]

output_txtfile = open(working_path+"output.txt", mode='w',encoding='utf-8')
output_figfile = working_path+"\\Figure\\"

color_list = ["red", "blue"]
symbol_list = [".", "^"]
label_list = ["20Pa 320V","35Pa 320V"]

def Paschen(x,a,b,d):
    return a*x*d/np.log(x*d) + b

param_bounds = ([0,0,0], [np.inf,np.inf,np.inf])

plt.figure

#-------------------------------------drawing the figure for the Paschen's Law-------------------------------------#


column = sheet.col(0)[1:9]
for i in range(len(column)):
    column[i] = column[i].value

Pressure = np.array(column, dtype=np.float64)

column = sheet.col(1)[1:9]
for i in range(len(column)):
    column[i] = column[i].value

Ignition_voltage = np.array(column, dtype=np.float64)

result = optimize.curve_fit(Paschen, Pressure, Ignition_voltage, maxfev = 500000, bounds = param_bounds)
A,B,D = result[0]
perr = np.sqrt(np.diag(result[1]))

mean = np.mean(Ignition_voltage)
ss_tot = np.sum((Ignition_voltage - mean) ** 2)
ss_res = np.sum((Ignition_voltage - Paschen(Pressure, A, B, D)) ** 2)
r_squared = 1 - (ss_res / ss_tot)

print("拟合得到Paschen定律的的参数\n\n", file=output_txtfile)
print("a\t b \t d\t", file=output_txtfile)
print("{0}\t {1}\t {2}\t".format(A,B,D),file=output_txtfile)
print("\n\n\n", file = output_txtfile)

print("拟合得到瞬时方法的不确定度\n\n", file=output_txtfile)
print("a\t b \t d\t", file=output_txtfile)
print("{0}\t {1}\t {2}\t".format(perr[0],perr[1],perr[2]),file=output_txtfile)
print("\n\n\n", file = output_txtfile)

print("拟合优度\n\n", file=output_txtfile)
print("R^2", file=output_txtfile)
print(r_squared,file=output_txtfile)
print("\n\n\n", file = output_txtfile)

x = np.arange(6.5,100)

y = Paschen(x, A, B, D)


plt.scatter(Pressure,Ignition_voltage)
plt.plot(x,y,c = "purple")
plt.title("Paschen's Law")
plt.xlabel("Pressure /Pa")
plt.ylabel("Ignition voltage /V")
plt.savefig(output_figfile + "Paschen's_Law.jpg")
plt.cla()


#-------------------------------------drawing the figure for the Diagnosis-------------------------------------#


for j in range(2):
    column = sheet.col(4+2*j)[1:]

    for i in range(len(column)):
        column[i] = column[i].value

    V = np.array(column, dtype=np.float64)



    column = sheet.col(5+2*j)[1:]

    for i in range(len(column)):
        column[i] = column[i].value

    I = np.array(column, dtype=np.float64)


    plt.plot(V,I,c = color_list[j], marker = symbol_list[j], label = label_list[j])

plt.title("Diagnosis")
plt.xlabel("V /V")
plt.ylabel("I /uA")
plt.legend()
plt.savefig(output_figfile + "Diagnosis.jpg")
plt.cla()
