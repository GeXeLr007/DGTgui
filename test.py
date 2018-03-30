import xlrd
from numpy import array
import numpy as np
import matplotlib.pyplot as plt
import Calculator
import matplotlib.pyplot as plt


xlsfile = r"data.xls"  # 打开指定路径中的xls文件

book = xlrd.open_workbook(xlsfile, formatting_info=True)  # 得到Excel文件的book对象，实例化对象
sheet0 = book.sheet_by_index(0)  # 通过sheet索引获得sheet对象

delt_g = []
Mab_plus = []
delt_g0 = sheet0.cell_value(1, 0)
for i in range(1, sheet0.nrows):
    values = []
    for j in range(2):
        values.append(sheet0.cell_value(i, j))
    delt_g.append(values[0])
    Mab_plus.append(values[1])

dmDict = {
    'Ag': 14.11,
    'Al': 4.75,
    'Cd': 6.09,
    'Co': 5.94,
    'Cr': 5.05,
    'Cu': 6.23,
    'Fe': 6.11,
    'Mn': 5.85,
    'Ni': 5.77,
    'Pb': 8.03,
    'Zn': 6.08
}
dm = dmDict['Cd'] * 10 ** (-6)
dml = 0.2 * dm
A = 4.64
t = 86400

cdgt = Mab_plus[0] * delt_g[0] / (dm * A * t)
Mab = [Mab_plus[0]]
for i in range(1, sheet0.nrows - 1):
    Mab.append(cdgt * dm * A * t / delt_g[i])
y = []
x = []
for i in range(sheet0.nrows - 1):
    x.append(delt_g[i] - delt_g0)
    y.append((Mab_plus[i] - Mab[i]) * x[i] / (dml * A * t))

xdata = array(x)
ydata = array(y)

# plt.plot(x, Mab)
# plt.plot(x, Mab_plus)
# plt.show()

n = 2
x_l = np.array([0, 0])
x_u = np.array([100, 1])

m_size = 50
# 原来是f = 0.5，实数编码变异的变化范围参数
f = 0.5
# 原来是cr = 0.3，交叉概率
cr = 0.3
iterate_times = int(10000)
checkN = 100
leastsqN = checkN / 10
ratio = 1 / 10
eps = 1e-10
calc = Calculator.Calculator(xdata, ydata, dml)

best_x_i, best_result_record, average_result_record = calc.fitting(n=n, m_size=m_size, f=f, cr=cr,
                                                                   iterate_times=iterate_times, x_l=x_l, x_u=x_u,
                                                                   leastsqN=leastsqN,
                                                                   ratio=ratio, eps=eps, checkN=checkN)

plt.plot(best_result_record)
plt.plot(average_result_record)
plt.show()
