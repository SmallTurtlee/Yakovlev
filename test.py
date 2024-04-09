from PyQt5.QtWidgets import     QApplication, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem, QPushButton, QHBoxLayout, QCheckBox, QGroupBox, QScrollArea, QScrollBar
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot
import sys
import openpyxl
from PyQt5 import QtCore, QtGui, QtWidgets
import matplotlib
matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg, NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure

class MplCanvas(FigureCanvasQTAgg):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = fig.add_subplot(111)
        self.axes.set_xlabel("time")
        super(MplCanvas, self).__init__(fig)

class Main(QWidget):
    def __init__(self):
        super(Main, self).__init__()
        self.setWindowTitle("Loading Excel")
        self.checkboxes=[]
        self.number_of_grafics = 1
        self.canvases = []
        layout = QVBoxLayout() 
        self.manygrafics = QVBoxLayout()
        grafic = QHBoxLayout()
        checkboxes1 = QGroupBox()
        checkboxes1.setTitle("График №1")
        checkboxes1.setFixedWidth(250)
        self.boxes = []
        self.boxes.append(QVBoxLayout())
        self.checkboxes = []
        self.checkboxes.append([])
        self.setLayout(layout)
        self.canvases.append(MplCanvas(self, width=5, height=4, dpi=100))
        toolbar = NavigationToolbar(self.canvases[0], self)
        self.table_widget = QTableWidget()
        layout.addLayout(self.button_block())   #
        layout.addWidget(toolbar)               #
        grafic.addWidget(self.canvases[0])
        grafic.addWidget(checkboxes1)
        checkboxes1.setLayout(self.boxes[0])
        layout.addLayout(self.manygrafics)      #
        self.manygrafics.addLayout(grafic)
        layout.addWidget(self.table_widget)     #

    def button_block(self):        #основные кнопки верхней панели
        block = QHBoxLayout()
        self.button1 = QPushButton("Открыть новый файл")
        self.button2 = QPushButton("Создать график")
        self.button3 = QPushButton("Добавить координаты")
        #self.button4 = QPushButton("Удалить координаты")
        self.button1.clicked.connect(self.on_click)
        self.button2.clicked.connect(self.on_click2)
        self.button3.clicked.connect(self.on_click3)
        #self.button4.clicked.connect(self.delete_graph)
        block.addWidget(self.button1)
        block.addWidget(self.button2)
        block.addWidget(self.button3)
        #block.addWidget(self.button4)
        return block
        
    @pyqtSlot()
    def on_click(self):
        name, done1 = QtWidgets.QInputDialog.getText(
			self, 'Input Dialog', 'Введите название или путь к нужному файлу:') 
        if done1: 
            self.load_data(name)
            self.make_checkboxes()
            self.clear_canvases()
            
    def on_click2(self):
        for i in self.canvases: i.axes.clear()
        list_values = list(self.sheet.values)
        legend = list(list_values[0][1:])
        Main_sheet = list(reversed(list(zip(*list_values[1:]))))
        for graph in range(self.number_of_grafics):
            for i in range(len(legend)):
                if self.checkboxes[graph][i].isChecked():
                    self.canvases[graph].axes.plot(Main_sheet[-1], Main_sheet[-1*(2+i)])
            self.canvases[graph].draw()  

    def on_click3(self):
        self.number_of_grafics += 1
        grafic = QHBoxLayout()
        canv = QVBoxLayout()
        checkboxes1 = QGroupBox()
        checkboxes1.setTitle("График №"+str(self.number_of_grafics))
        checkboxes1.setFixedWidth(250)
        self.boxes.append(QVBoxLayout())
        self.checkboxes.append([])
        for i in range(len(self.legend) - 1):
                self.checkboxes[-1].append(QCheckBox(self.legend[i+1], self))
                self.boxes[-1].addWidget(self.checkboxes[-1][i])
        checkboxes1.setLayout(self.boxes[-1])
        self.canvases.append(MplCanvas(self, width=5, height=4, dpi=100))
        toolbar = NavigationToolbar(self.canvases[self.number_of_grafics-1], self)
        canv.addWidget(toolbar)
        canv.addWidget(self.canvases[self.number_of_grafics-1])
        grafic.addLayout(canv)
        grafic.addWidget(checkboxes1)
        self.manygrafics.addLayout(grafic)

    def clear_canvases(self):
        for i in range(self.number_of_grafics):
            self.canvases[i].axes.clear()

    def load_data(self, path):
        workbook = None
        workbook = openpyxl.load_workbook(path)
        self.sheet = workbook.active

        self.table_widget.setRowCount(self.sheet.max_row - 1)
        self.table_widget.setColumnCount(self.sheet.max_column)

        list_values = list(self.sheet.values)
        self.legend = list(list_values[0])                                   #self.legeng - матрица со всеми названиями, включая ось х
        self.table_widget.setHorizontalHeaderLabels(self.legend)

        row_index = 0
        for value_tuple in list_values[1:]:
            column_index = 0
            for value in value_tuple:
                self.table_widget.setItem(row_index, column_index,QTableWidgetItem(str(value)))
                column_index += 1
            row_index += 1

    def make_checkboxes(self):
        for graph in range(self.number_of_grafics):
            for i in self.checkboxes[graph]: self.boxes[graph].removeWidget(i)                
        self.checkboxes =[]
        for graph in range(self.number_of_grafics):
            self.checkboxes.append([])
            for i in range(len(self.legend) - 1):
                self.checkboxes[graph].append(QCheckBox(self.legend[i+1], self))
                self.boxes[graph].addWidget(self.checkboxes[graph][i])


app = QApplication(sys.argv)
window = Main()
window.showMaximized()
app.exec_()