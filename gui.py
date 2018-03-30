# -*- coding: utf-8 -*-

from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QLabel, QInputDialog, QGridLayout, QMessageBox,
                             QLineEdit, QDesktopWidget)
from PyQt5.QtGui import QIcon
import sys
from numpy import array
import numpy as np
import xlrd
from xlutils.copy import copy
import Calculator
import time
from sklearn.metrics import r2_score
import math
import matplotlib.pyplot as plt


class Example(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):

        self.center()
        self.setWindowTitle('DGT参数计算软件')
        self.setWindowIcon(QIcon('logo.ico'))

        self.lb1 = QLabel('元素：')
        self.lb2 = QLabel('窗口面积A：')
        self.lb3 = QLabel('时间t：')
        self.lb4 = QLabel('Cd')
        self.lb5 = QLabel('4.64')
        self.lb6 = QLabel('86400')

        self.lb7 = QLabel('收敛指标：')
        self.lb8 = QLabel('1e-10')
        self.lb9 = QLabel('最大迭代次数：')
        self.lb10 = QLabel('10000')

        self.bt1 = QPushButton('修改元素')
        self.bt2 = QPushButton('修改窗口面积')
        self.bt3 = QPushButton('修改时间')
        self.bt4 = QPushButton('修改收敛指标')
        self.bt5 = QPushButton('修改最大迭代次数')
        self.btCalc = QPushButton('开始计算')

        self.bt1.clicked.connect(self.showDialog)
        self.bt2.clicked.connect(self.showDialog)
        self.bt3.clicked.connect(self.showDialog)
        self.bt4.clicked.connect(self.showDialog)
        self.bt5.clicked.connect(self.showDialog)

        self.btCalc.clicked.connect(self.calc)

        grid = QGridLayout()
        grid.setSpacing(10)
        grid.addWidget(self.lb1, 0, 0)
        grid.addWidget(self.lb2, 1, 0)
        grid.addWidget(self.lb3, 2, 0)
        grid.addWidget(self.lb4, 0, 1)
        grid.addWidget(self.lb5, 1, 1)
        grid.addWidget(self.lb6, 2, 1)
        grid.addWidget(self.lb7, 3, 0)
        grid.addWidget(self.lb8, 3, 1)
        grid.addWidget(self.lb9, 4, 0)
        grid.addWidget(self.lb10, 4, 1)
        grid.addWidget(self.bt1, 0, 2)
        grid.addWidget(self.bt2, 1, 2)
        grid.addWidget(self.bt3, 2, 2)
        grid.addWidget(self.bt4, 3, 2)
        grid.addWidget(self.bt5, 4, 2)
        grid.addWidget(self.btCalc, 5, 1)

        self.setLayout(grid)
        # self.setFixedSize(self.width(), self.height())

    def center(self):
        screen = QDesktopWidget().screenGeometry()  # QDesktopWidget为一个类，调用screenGeometry函数获得屏幕的尺寸
        self.resize(screen.width() / 4, screen.height() / 3)
        size = self.geometry()  # 同上
        self.move((screen.width() - size.width()) / 2, (screen.height() - size.height()) / 2)  # 调用move移动到指定位置

    def showDialog(self):
        sender = self.sender()
        element = ['Ag', 'Al', 'Cd', 'Co', 'Cr', 'Cu', 'Fe', 'Mn', 'Ni', 'Pb', 'Zn']
        if sender == self.bt1:
            text, ok = QInputDialog.getItem(self, '修改元素', '请选择元素：', element, current=element.index(self.lb4.text()))
            if ok:
                self.lb4.setText(text)
        elif sender == self.bt2:
            text, ok = QInputDialog.getDouble(self, '修改面积', '请输入窗口面积：', min=1.0, decimals=2,
                                              value=float(self.lb5.text()))
            if ok:
                self.lb5.setText(str(text))
        elif sender == self.bt3:
            text, ok = QInputDialog.getInt(self, '修改时间', '请输入时间：', min=1, value=int(self.lb6.text()))
            if ok:
                self.lb6.setText(str(text))
        elif sender == self.bt4:
            while (True):
                try:
                    text, ok = QInputDialog.getText(self, '修改收敛指标', '请输入收敛指标：', QLineEdit.Normal, self.lb8.text())
                    if ok and text.strip() != '':
                        d = float(text)
                        self.lb8.setText(str(text))
                        print(d)
                        print(type(d))
                        break
                    else:
                        if not ok:
                            break
                        elif text.strip() == '':
                            QMessageBox.about(self, '错误', '请输入有效收敛指标')
                except Exception as e:
                    QMessageBox.about(self, '错误', '请输入有效收敛指标')

        elif sender == self.bt5:
            text, ok = QInputDialog.getInt(self, '修改最大迭代次数', '最大迭代次数：', min=1, value=int(self.lb10.text()))
            if ok:
                self.lb10.setText(str(text))

    def calc(self):
        xlsfile = r"data.xls"  # 打开指定路径中的xls文件
        file = None
        try:
            file = open(xlsfile, 'a')
        except PermissionError as e:
            QMessageBox.about(self, '错误', '请先关闭data.xls再开始计算')
            return
        finally:
            if file:
                file.close()

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
        dm = dmDict[self.lb4.text()] * 10 ** (-6)
        dml = 0.2 * dm
        A = float(self.lb5.text())
        t = int(self.lb6.text())

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

        n = 2
        x_l = np.array([0, 0])
        x_u = np.array([100, 1])

        m_size = 50
        # 原来是f = 0.5，实数编码变异的变化范围参数
        f = 0.5
        # 原来是cr = 0.3，交叉概率
        cr = 0.3
        iterate_times = int(self.lb10.text())
        eps = float(self.lb8.text())
        checkN = 100
        leastsqN = checkN / 10
        ratio = 1 / 10
        eps = 1e-10

        calc = Calculator.Calculator(xdata, ydata, dml)

        best_x_i, best_result_record, average_result_record = calc.fitting(n=n, m_size=m_size, f=f, cr=cr,
                                                                           iterate_times=iterate_times,
                                                                           x_l=x_l, x_u=x_u,
                                                                           leastsqN=leastsqN,
                                                                           ratio=ratio, eps=eps, checkN=checkN)

        timeStr = time.strftime('%Y-%m-%d %H-%M-%S', time.localtime(time.time()))
        timeStrTotal = time.strftime('%Y/%m/%d %H:%M:%S', time.localtime(time.time()))
        newWb = copy(book)
        sheet1 = newWb.add_sheet(timeStr)
        sheet1.write(0, 0, '计算结果')
        sheet1.write(0, 1, '当前日期 ' + timeStrTotal)
        i = 1
        sheet1.write(i, 1, '方程系数')
        sheet1.write(i + 1, 0, 'c1')
        sheet1.write(i + 1, 1, best_x_i[0])
        sheet1.write(i + 2, 0, 'c2')
        sheet1.write(i + 2, 1, best_x_i[1])
        i += 3
        sheet1.write(i, 0, '观察值')
        sheet1.write(i, 1, '拟合值')
        sheet1.write(i, 2, '残差')

        ypred = []
        for j in range(len(calc.ydata)):
            sheet1.write(i + 1 + j, 0, calc.ydata[j])
        err = calc.residuals(best_x_i, calc.ydata, calc.xdata)
        for j in range(len(calc.ydata)):
            sheet1.write(i + 1 + j, 2, err[j])
        for j in range(len(calc.ydata)):
            ypred.append(calc.ydata[j] - err[j])
            sheet1.write(i + 1 + j, 1, ypred[j])

        i += (1 + len(calc.ydata))
        rr = r2_score(ydata, ypred)
        r = math.sqrt(rr)
        sheet1.write(i, 0, '相关R=')
        sheet1.write(i, 1, r)
        sheet1.write(i, 2, '相关RR=')
        sheet1.write(i, 3, rr)

        i += 1
        sheet1.write(i, 0, '厚度差')
        sheet1.write(i, 1, 'Mab+')
        sheet1.write(i, 2, 'Mab')
        i += 1
        for j in range(len(x)):
            sheet1.write(i + j, 0, x[j])
            sheet1.write(i + j, 1, Mab_plus[j])
            sheet1.write(i + j, 2, Mab[j])

        newWb.save(xlsfile)
        QMessageBox.about(self, '完成', '已完成计算，请打开data.xls')

        plt.figure(1)
        self.myplot(yplot1=average_result_record, yplot2=best_result_record,
                    title_str=u"$change\quadof\quadaverage\quadand\quadbest\quadresult$",
                    x_label=u"$iteration\quadtime$",
                    y_label=u"$loss$", legend1=u"$average\quadresult$", legend2=u"$best\quadresult$")

        plt.figure(2)
        self.myplot(yplot1=Mab_plus, yplot2=Mab, xplot=x,
                    title_str=u"$change\quadof\quadM\quadwith\quaddifferent\quad\Delta g$",
                    x_label=u"$\Delta g/cm$",
                    y_label=u"$M/ng$", legend1=u"$Ma^{b+}$", legend2=u"$Ma^{b}$", isO='o')

    def myplot(self, yplot1, yplot2, title_str, x_label, y_label, legend1, legend2, isO='', xplot=None):
        if xplot is None:
            xplot = [i + 1 for i in range(len(yplot1))]

        plt.plot(xplot, yplot1, 'r' + isO + '-', label=legend1, linewidth=1)
        plt.plot(xplot, yplot2, 'b' + isO + '-', label=legend2, linewidth=1)

        plt.title(title_str, size=20)
        plt.legend()

        ax = plt.gca()
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        ax.xaxis.set_ticks_position('bottom')
        ax.spines['bottom'].set_position(('data', 0))
        ax.yaxis.set_ticks_position('left')
        ax.spines['left'].set_position(('data', 0))

        ax.set_xlim(0, xplot[len(xplot) - 1] * 1.05)
        ax.set_ylim(0, yplot1[0] * 1.05)

        plt.xlabel(x_label, size=20)
        plt.ylabel(y_label, size=20)

        plt.grid(True)

        plt.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    ex.show()
    sys.exit(app.exec_())
