from encodings import utf_8
import xlrd
import numpy as np
import matplotlib.pyplot as plt
from scipy import optimize

#-------------------------------------definition of functions and working path-------------------------------------#

working_path = "D:\DeskTop\学习用\\2022春季学期\大物实验五级报告\闪烁体荧光\\20220413\\"
file_list = ["NaI","CeF3","BGO","BaF2","塑料","有机玻璃"]
data_index = range(1,11)

def funcform_1(x,A,B,C):
    return A*(np.exp(-x/B,dtype = np.float64)-np.exp(-x/C,dtype = np.float64))

def funcform_2(x,A,B):
    return A*np.exp(-x/B,dtype = np.float64)

output = open(working_path+"output.txt", mode='w',encoding='utf-8')

plt.figure

param_bounds = ([-np.inf, 0],[0, 1000])

#-------------------------------------fitting for NaI and find RC-------------------------------------#

print("利用NaI拟合RC值的大小",file = output)
print("组数\tRC",file = output)

excel = xlrd.open_workbook(working_path+file_list[0]+".xlsx",encoding_override = "utf-8")
sheet = (excel.sheets())[0]

def func_RC(x,A,C):
    return funcform_1(x,A,250,C)

RC = []


for index in data_index:
    column = sheet.col(2*index-1)
    column.pop(0)
    for i in range(len(column)):
        column[i] = column[i].value
    min_idx = column.index(min(column))
    voltage = column[min_idx:]
    time = np.arange(1,len(voltage)+1)
    voltage_arange = np.array(voltage, dtype=np.float64)

    A,RC_value = optimize.curve_fit(func_RC,time,voltage_arange, bounds = param_bounds)[0]

    RC.append(RC_value)
    
    y = func_RC(time,A,RC[index-1])

    print(str(index)+"\t" + str(RC[index-1]),file = output)

    plt.scatter(time, voltage_arange, s = 2, c = "red")
    plt.plot(time,y,c = "purple")
    plt.savefig(working_path+"Figure\\" + file_list[0] + " " + str(index) + ".jpg")
    plt.cla()

RC_mean = np.mean(RC)
print("得到RC的均值为"+str(RC_mean)+"ns",file = output)

#-------------------------------------using RC to fit for different fluorophor-------------------------------------#

print("利用已知RC拟合其他荧光体本征时间的大小",file = output)


def func_tau(x,A,B):
    return funcform_1(x,A,B,RC_mean)

tau_mean = []

for Index in range(1,len(file_list)):
    excel = xlrd.open_workbook(working_path+file_list[Index]+".xlsx",encoding_override = "utf-8")
    sheet = (excel.sheets())[0]
    tau = []
    print(file_list[Index]+"的荧光体特征时间",file = output)
    print("组数\ttau",file = output)
    for index in data_index:
        column = sheet.col(2*index-1)
        column.pop(0)
        for i in range(len(column)):
            column[i] = column[i].value
        min_idx = column.index(min(column))
        voltage = column[min_idx:]
        time = np.arange(1,len(voltage)+1)
        voltage_arange = np.array(voltage)

        A,tau_value = optimize.curve_fit(func_tau,time,voltage_arange, bounds = param_bounds)[0]

        tau.append(tau_value)
    
        y = func_tau(time,A,tau_value)

        print(str(index)+"\t" + str(tau_value),file = output)

        plt.scatter(time, voltage_arange, s = 2, c = "red")
        plt.plot(time,y,c = "purple")
        plt.savefig(working_path+"Figure\\" + file_list[Index] + " " + str(index) + " model A.jpg")
        plt.cla()

    tau_mean_value = np.mean(tau)
    tau_mean.append(tau_mean_value)
    print("得到该元素特征时间的均值为"+str(tau_mean_value)+"ns",file = output)


#-------------------------------------using model B to fit for different fluorophor-------------------------------------#




print("利用近似公式拟合其他荧光体本征时间的大小",file = output)


tau_mean_2 = []

for Index in range(len(file_list)):
    excel = xlrd.open_workbook(working_path+file_list[Index]+".xlsx",encoding_override = "utf-8")
    sheet = (excel.sheets())[0]
    tau = []
    print(file_list[Index]+"的荧光体特征时间",file = output)
    print("组数\ttau",file = output)
    for index in data_index:
        column = sheet.col(2*index-1)
        column.pop(0)
        for i in range(len(column)):
            column[i] = column[i].value
        min_idx = column.index(min(column))
        voltage = column[min_idx:]
        time = np.arange(1,len(voltage)+1)
        voltage_arange = np.array(voltage, dtype=np.float64)

        A,tau_value = optimize.curve_fit(funcform_2,time,voltage_arange, bounds = param_bounds)[0]

        tau.append(tau_value)
    
        y = funcform_2(time,A,tau_value)

        print(str(index)+"\t" + str(tau_value),file = output)

        plt.scatter(time, voltage_arange, s = 2, c = "red")
        plt.plot(time,y,c = "purple")
        plt.savefig(working_path+"Figure\\" + file_list[Index] + " " + str(index) + " model B.jpg")
        plt.cla()

    tau_mean_value = np.mean(tau)
    tau_mean_2.append(tau_mean_value)
    print("得到该元素特征时间的均值为"+str(tau_mean_value)+"ns",file = output)
