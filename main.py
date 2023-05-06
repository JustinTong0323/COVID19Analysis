from PySide2.QtWidgets import QMessageBox
from PySide2.QtCore import QFile  #这段c要补上，笔记文件里没有
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import *
from PySide2.QtWidgets import *
from PySide2.QtWebEngineWidgets import *
import xlrd
import pandas as pd
import numpy as np
from scipy.optimize import curve_fit
from matplotlib.pylab import mpl
from matplotlib import pyplot as plt
from PySide2.QtGui import QIcon


class Stats:

    def __init__(self):
        # 从文件中加载UI定义
        qfile_stats = QFile("ui/stats.ui")  # 加载的ui文件
        qfile_stats.open(QFile.ReadOnly)  # 这句和下句是固定写法
        qfile_stats.close()  # 关闭

        # 从 UI 定义中动态 创建一个相应的窗口对象
        # 注意：里面的控件对象也成为窗口对象的属性了
        # 比如 self.ui.button , self.ui.textEdit
        self.ui = QUiLoader().load(qfile_stats)  # 返回值是窗体对象（最外面的那个窗体）
        # tab1
        self.ui.pushButton_report.clicked.connect(self.OpenReport)

        # tab2
        self.ui.pushButton_loadexc.clicked.connect(self.Loadexc)
        self.ui.pushButton_savecsv.clicked.connect(self.Savecsv)

        # tab3
        self.ui.pushButton_loadcsv.clicked.connect(self.Loadcsv)
        self.ui.pushButton_visible.clicked.connect(self.Plot)

        # tab4
        self.ui.pushButton_loadcsv_2.clicked.connect(self.Loadcsv_2)
        self.ui.pushButton_forecast.clicked.connect(self.Forecast)
    # tab1
    def OpenReport(self):
        web = TabWidget()
        web.show()

    # tab2

    def Loadexc(self):
        self.ui.textBrowser_pathexc.clear()
        self.filePath, _ = QFileDialog.getOpenFileName(
            self.ui,  # 父窗口对象
            "选择你要上传的表格",  # 标题
            r"c:\\",  # 起始目录
            "表格类型 (*.xlsx *.xls)"  # 选择类型过滤项，过滤内容在括号中
        )
        if self.filePath:
            QMessageBox.information(
                self.ui,
                '操作成功',
                '请继续下一步操作')
            self.ui.textBrowser_pathexc.append(self.filePath)
        else:
            QMessageBox.critical(
                self.ui,
                '错误',
                '没有检测到对应路径！')

    def Savecsv(self):
        if self.ui.textBrowser_pathexc.toPlainText():
            if self.ui.lineEdit_NewFileName.text():
                try:
                    NewFileName = self.ui.lineEdit_NewFileName.text()
                    Number = self.ui.spinBox_formnub.value() - 1
                    NewFilePath = QFileDialog.getExistingDirectory(self.ui, "选择存储路径")
                    if NewFilePath:

                        # 读取Excel数据
                        book = xlrd.open_workbook(self.filePath)
                        sheet = book.sheet_by_index(Number)

                        sheet_col = sheet.col_values(colx=0)
                        sheet_row = sheet.row_values(rowx=0)  # 表头
                        # 过滤后的每行存入二维数组中，【行坐标，0开始，到最后一个数据结束】【列坐标，0开始到最后一个标题结束】
                        TwoArray = [[' ' for i in range(len(sheet_row))] for i in range(len(sheet_col))]  # 二维数组初始化
                        for r in range(len(sheet_col)):
                            for c in range(len(sheet_row)):
                                TwoArray[r][c] = sheet.row_values(rowx=r)[c]

                        FinalPath = f'{NewFilePath}/{NewFileName}.csv'
                        #print(TwoArray)
                        if TwoArray[0][0] == 'date' and TwoArray[0][1] == 'confirm' and TwoArray[0][2] == 'heal':

                            with open(FinalPath, mode='a') as f:
                                for r in range(len(sheet_col)):  # 0 1
                                    for c in range(len(sheet_row)):  # 0 1 2
                                        if c == len(sheet_row) - 1:
                                            f.write(f'{TwoArray[r][c]}\n')
                                        else:
                                            f.write(f'{TwoArray[r][c]},')
                            QMessageBox.information(
                                self.ui,
                                '导出成功',
                                '导出csv格式文件成功，请到指定路径的文件夹下查看')
                        else:
                            QMessageBox.critical(
                                self.ui,
                                '错误',
                                '文件内容错误！')
                    else:
                        QMessageBox.critical(
                            self.ui,
                            '错误',
                            '未检测到路径！')
                except IndexError as i:
                        QMessageBox.critical(
                            self.ui,
                            '错误',
                            f'格式转换失败,错误信息为表单号错误:\'{i}\'')

                except Exception as e:
                    QMessageBox.critical(
                        self.ui,
                        '错误',
                        f'格式转换失败,错误信息为{e}')

            else:
                QMessageBox.critical(
                    self.ui,
                    '错误',
                    '缺失文件名！')
        else:
            QMessageBox.critical(
                self.ui,
                '错误',
                '请先导入EXCEL文件！')

    # tab3
    def Loadcsv(self):
        self.ui.textBrowser_pathcsv.clear()
        self.filePath1, _ = QFileDialog.getOpenFileName(
            self.ui,  # 父窗口对象
            "选择你要上传的csv文件",  # 标题
            r"c:\\",  # 起始目录
            "csv类型 (*.csv)"  # 选择类型过滤项，过滤内容在括号中
        )
        if self.filePath1:
            with open(f'{self.filePath1}', 'r') as f:
                line = f.readline()
            li = line.split(',')

            if li[0] == 'date' and li[1] == 'confirm' and li[2] == 'heal\n':

                QMessageBox.information(
                    self.ui,
                    '操作成功',
                    '请继续下一步操作')
                self.ui.textBrowser_pathcsv.append(self.filePath1)
            else:
                QMessageBox.critical(
                    self.ui,
                    '错误',
                    '文件内容错误！')
                self.filePath1 = None
        else:
            QMessageBox.critical(
                self.ui,
                '错误',
                '没有检测到对应路径！')

    def Plot(self):
        self.filePath1 = self.ui.textBrowser_pathcsv.toPlainText()
        if self.ui.textBrowser_pathcsv.toPlainText():
            try:
                Type = self.ui.comboBox_PlotType.currentText()
                mpl.rcParams['font.sans-serif'] = ['SimHei']  # 解决中文乱码问题
                data = pd.read_csv(self.filePath1)
                x = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24,
                     25]  # 依次对应2020/7/10-2020/8/3
                y = data['confirm']
                z = data['heal']
                if Type == '散点图':
                    plt1 = plt.plot(x, y, 's', label='疫情确诊人数')
                    plt2 = plt.plot(x, z, 's', label='治愈人数')
                    plt.legend(loc=0)
                    plt.title('疫情趋势图（散点）')  # 加标题
                    plt.ylabel('人数')
                    plt.xlabel("日期")
                    plt.show()
                elif Type == '折线图':
                    plt.plot(x, y, 'r', label='疫情确诊人数')
                    plt.title('疫情趋势图（折线）')  # 加标题
                    plt.ylabel('人数')
                    plt.xlabel("日期")
                    plt.legend(loc=0)
                    plt.show()

            except Exception as e:
                QMessageBox.critical(
                    self.ui,
                    '错误',
                    f'可视化失败,错误信息为{e}')
        else:
            QMessageBox.critical(
                self.ui,
                '错误',
                '请先导入csv文件！')

    # tab4
    def Loadcsv_2(self):
        self.ui.textBrowser_pathcsv_2.clear()
        self.filePath2, _ = QFileDialog.getOpenFileName(
            self.ui,  # 父窗口对象
            "选择你要上传的csv文件",  # 标题
            r"c:\\",  # 起始目录
            "csv类型 (*.csv)"  # 选择类型过滤项，过滤内容在括号中
        )
        if self.filePath2:
            with open(f'{self.filePath2}', 'r') as f:
                line = f.readline()
            li = line.split(',')

            if li[0] == 'date' and li[1] == 'confirm' and li[2] == 'heal\n':

                QMessageBox.information(
                    self.ui,
                    '操作成功',
                    '请继续下一步操作')
                self.ui.textBrowser_pathcsv_2.append(self.filePath2)
            else:
                QMessageBox.critical(
                    self.ui,
                    '错误',
                    '文件内容错误！')
                self.filePath2 = None
        else:
            QMessageBox.critical(
                self.ui,
                '错误',
                '没有检测到对应路径！')

    def Forecast(self):
        if self.ui.textBrowser_pathcsv_2.toPlainText():
            try:
                Type = self.ui.comboBox_ForecastType.currentText()
                mpl.rcParams['font.sans-serif'] = ['SimHei']  # 解决中文乱码问题
                data = pd.read_csv(self.filePath2)
                x = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24,
                     25]  # 依次对应2020/7/10-2020/8/3
                y = data['confirm']
                z = data['heal']
                plt.plot(x, y, 'r', label='疫情确诊人数')
                plt.title('疫情趋势图（折线）')  # 加标题
                plt.ylabel('人数')
                plt.xlabel("日期")
                plt.legend(loc=0)
                plt.show()

                def logistic_increase_function(t, K, P0, r):
                    r = 0.035  # 外国防控消极
                    t0 = 1
                    exp_value = np.exp(r * (t - t0))
                    return (K * exp_value * P0) / (K + (exp_value - 1) * P0)
                    # 日期与感染人数
                t = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25]
                t = np.array(t)
                P = data['confirm']
                # 最小二乘拟合
                P = np.array(P)  # 转化为数组
                popt, pocv = curve_fit(logistic_increase_function, t, P)

                # 所获取的opt皆为拟合系数
                print("K:capacity P0:intitial_value r:increase_rate t:time")
                print(popt)
                # 拟合后对未来情况进行预测
                P_predict = logistic_increase_function(t, popt[0], popt[1], popt[2])
                future = [26, 35, 50, 70, 90, 130, 180, 250]
                future = np.array(future)
                future_predict = logistic_increase_function(future, popt[0], popt[1], popt[2])

                if Type == '短期预测':
                    # 近期情况
                    tomorrow = [26, 27, 28, 29, 30, 31, 32]
                    tomorrow = np.array(tomorrow)
                    tomorrow_predict = logistic_increase_function(tomorrow, popt[0], popt[1], popt[2])
                    # 图像绘制
                    plot1 = plt.plot(t, P, 's', label="疫情确诊感染人数")
                    plot2 = plt.plot(t, P_predict, 'r', label='感染人数拟合曲线')
                    plot3 = plt.plot(tomorrow, tomorrow_predict, 's', label='近期感染人数预测')

                    plt.title('近期感染人数预测')
                    plt.xlabel('日期')
                    plt.ylabel('确诊人数')
                    plt.legend(loc=0)
                    plt.grid()
                    plt.show()
                    self.ui.textBrowser_forecast.clear()
                    for i in range(0, 7):
                        Sicknumber = logistic_increase_function(tomorrow[i], popt[0], popt[1], popt[2])
                        self.ui.textBrowser_forecast.append(f'预测的第{i + 1}天：{Sicknumber}人')

                else:
                    plot1 = plt.plot(t, P, 's', label="疫情确诊感染人数")
                    plot2 = plt.plot(t, P_predict, 'r', label='感染人数拟合曲线')
                    plot4 = plt.plot(future, future_predict, 's', label='未来感染人数预测')
                    plt.title('未来感染人数预测(长期/拐点预测)')  # 加标题
                    plt.show()  # 显示网格

            except Exception as e:
                QMessageBox.critical(
                    self.ui,
                    '错误',
                    f'可视化失败,错误信息为{e}')

        else:
            QMessageBox.critical(
                self.ui,
                '错误',
                '请先导入csv文件！')

class TabWidget(QTabWidget):#分页
    def __init__(self, *args, **kwargs):
        QTabWidget.__init__(self, *args, **kwargs)
        url = QUrl("https://voice.baidu.com/act/newpneumonia/newpneumonia/?from=osari_aladin_banner#tab4")
        view = HtmlView(self)
        view.load(url)
        ix = self.addTab(view, "实时疫情大数据报告")
        self.resize(800, 800)

#内嵌浏览器
class HtmlView(QWebEngineView):
    def __init__(self, *args, **kwargs):
        QWebEngineView.__init__(self, *args, **kwargs)
        self.tab = self.parent()

    def createWindow(self, windowType):
        if windowType == QWebEnginePage.WebBrowserTab:
            webView = HtmlView(self.tab)
            ix = self.tab.addTab(webView, "实时疫情大报告")
            self.tab.setCurrentIndex(ix)
            return webView
        return QWebEngineView.createWindow(self, windowType)

if __name__ == "__main__":
    app = QApplication([])
    app.setWindowIcon(QIcon('ui/Logo.png'))
    stats = Stats()
    stats.ui.show()
    app.exec_()
    # pyinstaller main.py --noconsole  --hidden-import PySide2.QtXml --icon="ui/Logo.ico"
    # pyinstaller main.py  --hidden-import PySide2.QtXml --icon="ui/Logo.ico"
