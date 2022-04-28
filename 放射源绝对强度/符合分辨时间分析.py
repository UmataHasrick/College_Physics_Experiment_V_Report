from encodings import utf_8
import xlrd
import numpy as np
import matplotlib.pyplot as plt
from scipy import optimize

#-------------------------------------definition of working path and open the excel document-------------------------------------#

working_path = "D:\DeskTop\学习用\\2022春季学期\大物实验五级报告\放射源绝对强度\\"

excel = xlrd.open_workbook(working_path+"data.xlsx",encoding_override = "utf-8")
sheet = (excel.sheets())[0]

output = open(working_path+"output.txt", mode='w',encoding='utf-8')

def Gauss(t,a,t0,b):
    return a*np.exp(-((t-t0)/b)**2)

def Linear(x,a,b):
    return a*x+b

#-------------------------------------fitting for curve of instantaneous method-------------------------------------#

inst_t = []
inst_N = []

column = sheet.col(0)
column.pop(0)
column.pop(0)
for i in range(len(column)):
    column[i] = column[i].value

inst_t = np.array(column, dtype=np.float64)

column = sheet.col(1)
column.pop(0)
column.pop(0)
for i in range(len(column)):
    column[i] = column[i].value

inst_N = np.array(column, dtype=np.float64)

#print(inst_N)

result = optimize.curve_fit(Gauss,inst_t,inst_N)
A,T0,B = result[0]
perr = np.sqrt(np.diag(result[1]))

mean = np.mean(inst_N)
ss_tot = np.sum((inst_N - mean) ** 2)
ss_res = np.sum((inst_N - Gauss(inst_t, A,T0,B)) ** 2)
r_squared = 1 - (ss_res / ss_tot)


print("拟合得到瞬时方法的参数\n\n", file=output)
print("a\t t0 \t b\t", file=output)
print("{0}\t {1}\t {2}\t".format(A,T0,B),file=output)
print("\n\n\n", file = output)

print("拟合得到瞬时方法的不确定度\n\n", file=output)
print("a\t t0 \t b\t", file=output)
print("{0}\t {1}\t {2}\t".format(perr[0],perr[1],perr[2]),file=output)
print("\n\n\n", file = output)

print("拟合优度\n\n", file=output)
print("R^2", file=output)
print(r_squared,file=output)
print("\n\n\n", file = output)

y = Gauss(inst_t,A,T0,B)

plt.figure

plt.scatter(inst_t, inst_N, s = 2, c = "red")
plt.plot(inst_t,y,c = "purple")
plt.savefig(working_path+"Figure\\" + "Fitting_graph_for_instantaneous_method.jpg")
plt.cla()


#-------------------------------------fitting for curve of occasional method-------------------------------------#

column = sheet.col(3)
column = column[2:8]

for i in range(len(column)):
    column[i] = column[i].value

occa_n_beta = np.array(column, dtype=np.float64)

column = sheet.col(4)
column = column[2:8]
for i in range(len(column)):
    column[i] = column[i].value

occa_n_gamma = np.array(column, dtype=np.float64)

column = sheet.col(5)
column = column[2:8]
for i in range(len(column)):
    column[i] = column[i].value

occa_n_rc = np.array(column, dtype=np.float64)

occa_n_12 = occa_n_beta*occa_n_gamma

result = optimize.curve_fit(Linear,occa_n_12,occa_n_rc)
A,B = result[0]
perr = np.sqrt(np.diag(result[1]))

mean = np.mean(occa_n_rc)
ss_tot = np.sum((occa_n_rc - mean) ** 2)
ss_res = np.sum((occa_n_rc - Linear(occa_n_12, A, B)) ** 2)
r_squared = 1 - (ss_res / ss_tot)

print("拟合得到偶然符合方法的参数\n\n", file=output)
print("a\tb\t", file=output)
print("{0}\t{1}\t".format(A,B),file=output)
print("\n\n\n", file = output)

print("拟合得到偶然方法的不确定度\n\n", file=output)
print("a\tb\t", file=output)
print("{0}\t{1}\t".format(perr[0],perr[1]),file=output)
print("\n\n\n", file = output)

print("拟合优度\n\n", file=output)
print("R^2", file=output)
print(r_squared,file=output)
print("\n\n\n", file = output)

y = Linear(occa_n_12,A,B)

plt.scatter(occa_n_12, occa_n_rc, s = 2, c = "red")
plt.plot(occa_n_12,y,c = "purple")
plt.savefig(working_path+"Figure\\" + "Fitting_graph_for_occasional_method.jpg")
plt.cla()


#-------------------------------------modifying the curve of occasional method-------------------------------------#

occa_n_12 = np.delete(occa_n_12,[1,3])
occa_n_rc = np.delete(occa_n_rc,[1,3])


result = optimize.curve_fit(Linear,occa_n_12,occa_n_rc)
A,B = result[0]
perr = np.sqrt(np.diag(result[1]))

mean = np.mean(occa_n_rc)
ss_tot = np.sum((occa_n_rc - mean) ** 2)
ss_res = np.sum((occa_n_rc - Linear(occa_n_12, A, B)) ** 2)
r_squared = 1 - (ss_res / ss_tot)

print("修正后拟合得到偶然符合方法的参数\n\n", file=output)
print("a\tb\t", file=output)
print("{0}\t{1}\t".format(A,B),file=output)
print("\n\n\n", file = output)

print("修正后拟合得到偶然方法的不确定度\n\n", file=output)
print("a\tb\t", file=output)
print("{0}\t{1}\t".format(perr[0],perr[1]),file=output)
print("\n\n\n", file = output)

print("修正后拟合优度\n\n", file=output)
print("R^2", file=output)
print(r_squared,file=output)
print("\n\n\n", file = output)

y = Linear(occa_n_12,A,B)

plt.scatter(occa_n_12, occa_n_rc, s = 2, c = "red")
plt.plot(occa_n_12,y,c = "purple")
plt.savefig(working_path+"Figure\\" + "Fitting_graph_for_occasional_method_modified.jpg")
plt.cla()