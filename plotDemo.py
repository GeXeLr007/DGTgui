import xlrd
from numpy import array
import numpy as np
import matplotlib.pyplot as plt
import Calculator


def myplot(yplot1, yplot2, title_str, x_label, y_label, legend1, legend2, xplot=None):
    if xplot is None:
        xplot = [i + 1 for i in range(len(yplot1))]
    plt.figure(1)

    plt.plot(xplot, yplot1, 'ro-', label=legend1, linewidth=1)
    plt.plot(xplot, yplot2, 'bo-', label=legend2, linewidth=1)

    plt.title(title_str, size=20)
    plt.legend()

    ax = plt.gca()
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.xaxis.set_ticks_position('bottom')
    ax.spines['bottom'].set_position(('data', 0))
    ax.yaxis.set_ticks_position('left')
    ax.spines['left'].set_position(('data', 0))

    ax.set_xlim(0, x[len(x) - 1] * 1.05)
    ax.set_ylim(0, Mab_plus[0] * 1.05)

    plt.xlabel(x_label, size=20)
    plt.ylabel(y_label, size=20)

    plt.grid(True)

    plt.show()


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

# plt.plot([], [2, 3, 4, 7, 9])
# plt.show()
myplot(yplot1=Mab_plus, yplot2=Mab, xplot=x, title_str=u"$change\quadof\quadM\quadwith\quaddifferent\quad\Delta g$", x_label=u"$\Delta g/cm$",
       y_label=u"$M/ng$", legend1=u"$Ma^{b+}$", legend2=u"$Ma^{b}$")
