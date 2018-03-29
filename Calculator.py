from numpy import exp
import numpy as np
from scipy.optimize import leastsq
import matplotlib.pyplot as plt


class Calculator(object):
    def __init__(self, xdata, ydata, dml):
        super().__init__()
        self.xdata = xdata
        self.ydata = ydata
        self.dml = dml

    def residuals(self, p, y, x):
        c1, c2 = p
        err = y - c1 * (1 - exp(-c2 * x ** 2 / (2 * self.dml)))
        return err

    def evaluate_func(self, p):
        # c1, c2 = p
        # return sum((ydata - c1 * (1 - exp(-c2 * xdata ** 2 / (2 * dml)))) ** 2)
        return sum((self.residuals(p, self.ydata, self.xdata)) ** 2)

    def fitting(self, n, m_size, f, cr, iterate_times, x_l, x_u, leastsqN, ratio, eps,checkN):
        # 建立一个全是0的三维数组
        x_all = np.zeros((iterate_times, m_size, n))
        # 初始化第1代种群,第1代为随机生成
        for i in range(m_size):
            x_all[0][i] = x_l + np.random.random() * (x_u - x_l)
        # print('寻优参数个数为', n, '优化区间分别为', x_l, x_u)
        # g的值为0到iterate_times-2，x_all的索引范围是0--99一共100代，下标为iterate_times - 1的就是最后一代
        best_result_record = []
        end_iterate_time = 0
        for g in range(iterate_times - 1):
            evaluate_result_temp = [self.evaluate_func(x_all[g][i]) for i in range(m_size)]
            best_result_temp = evaluate_result_temp[np.argmin(evaluate_result_temp)]
            best_result_record.append(best_result_temp)
            # print(best_result_temp)
            # 连续checkN代的最优结果的差异小于eps即终止
            if g > checkN and ((best_result_record[g - checkN] - best_result_record[g]) <= eps):
                end_iterate_time = g
                break
            # 每leastsqN代执行一次LM算法寻找使残差平方和最小的参数
            if (g + 1) % leastsqN == 0:
                result = [self.evaluate_func(x_all[g][i]) for i in range(m_size)]
                p0 = x_all[g][np.argmin(result)]
                least = leastsq(self.residuals, p0, args=(self.ydata, self.xdata), maxfev=0)
                np.random.shuffle(x_all[g])
                x_all[g + 1] = x_all[g]
                # 用多少比例的LM算法较优解插入到下一代种群，这个参数的设置比较玄学。
                for i in range(int(m_size * ratio)):
                    x_all[g + 1][i] = least[0]
            else:
                for i in range(m_size):
                    x_g_without_i = np.delete(x_all[g], i, 0)
                    # 将种群个体的顺序打乱
                    np.random.shuffle(x_g_without_i)
                    # 不清楚f参数如何设置
                    h_i = x_g_without_i[1] + f * (x_g_without_i[2] - x_g_without_i[3])
                    h_i = [h_i[item] if h_i[item] < x_u[item] else x_u[item] for item in range(n)]
                    h_i = [h_i[item] if h_i[item] > x_l[item] else x_l[item] for item in range(n)]
                    # print(h_i)
                    v_i = np.array([x_all[g][i][j] if (np.random.random() > cr) else h_i[j] for j in range(n)])
                    # 默认求解的是  最小化  问题
                    if self.evaluate_func(x_all[g][i]) > self.evaluate_func(v_i):
                        x_all[g + 1][i] = v_i
                    else:
                        x_all[g + 1][i] = x_all[g][i]

            # 此处输出的代数从1开始计算，即第1代、第2代、第3代...
            # print('第', g + 2, '代')
            # print(x_all[g + 1])
        evaluate_result = [self.evaluate_func(x_all[end_iterate_time][i]) for i in range(m_size)]
        best_x_i = x_all[end_iterate_time][np.argmin(evaluate_result)]
        # plt.plot(best_result_record)
        # plt.show()
        # best_result = evaluate_result[np.argmin(evaluate_result)]

        # print(evaluate_result)
        print('最优个体为', best_x_i)
        # print('最优解为', best_result)
        return best_x_i
