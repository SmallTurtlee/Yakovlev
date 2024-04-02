from PyQt5.QtWidgets import     QApplication, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem, QPushButton, QHBoxLayout, QCheckBox, QGroupBox, QScrollArea, QScrollBar
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot
import sys
import openpyxl
from PyQt5 import QtCore, QtGui, QtWidgets
import sys
import matplotlib
matplotlib.use('Qt5Agg')

from PyQt5 import QtCore, QtWidgets

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
        self.w1=[]
        layout = QVBoxLayout() 
        block = QHBoxLayout()
        self.manygrafics = QVBoxLayout()
        grafic = QHBoxLayout()
        checkboxes1 = QGroupBox()
        checkboxes1.setTitle("График №1")
        checkboxes1.setFixedSize(250, 500)
        self.boxes1 = QVBoxLayout()
        self.setLayout(layout)
        self.button1 = QPushButton("Открыть новый файл")
        self.button2 = QPushButton("Создать график")
        self.button3 = QPushButton("Добавить координаты")
        self.sc1 = MplCanvas(self, width=5, height=4, dpi=100)
        toolbar = NavigationToolbar(self.sc1, self)
        self.button1.clicked.connect(self.on_click)
        self.button2.clicked.connect(self.on_click2)
        self.button3.clicked.connect(self.on_click3)
        self.table_widget = QTableWidget()
        block.addWidget(self.button1)
        block.addWidget(self.button2)
        block.addWidget(self.button3)
        layout.addLayout(block)
        layout.addWidget(toolbar)
        grafic.addWidget(self.sc1)
        grafic.addWidget(checkboxes1)
        checkboxes1.setLayout(self.boxes1)
        layout.addLayout(self.manygrafics)
        self.manygrafics.addLayout(grafic)
        layout.addWidget(self.table_widget)
        


        
    @pyqtSlot()
    def on_click(self):
        name, done1 = QtWidgets.QInputDialog.getText(
			self, 'Input Dialog', 'Введите название или путь к нужному файлу:') 
        if done1: 
            self.load_data(name)
            self.make_checkboxes()
            
    def on_click2(self):
        self.sc1.axes.clear()
        list_values = list(self.sheet.values)
        legend = list(list_values[0][1:])
        Main_sheet = list(reversed(list(zip(*list_values[1:]))))
        for i in range(len(legend)):
            if self.w1[i].isChecked():
                self.sc1.axes.plot(Main_sheet[-1], Main_sheet[-1*(2+i)])
        self.sc1.draw()  

    def on_click3(self):
        grafic = QHBoxLayout()
        checkboxes1 = QGroupBox()
        checkboxes1.setTitle("График №1")
        checkboxes1.setFixedWidth(250)
        self.manygrafics.addLayout(grafic)
        self.sc1 = MplCanvas(self, width=5, height=4, dpi=100)
        grafic.addWidget(self.sc1)
        grafic.addWidget(checkboxes1)

        

    def load_data(self, path):
        workbook = None
        workbook = openpyxl.load_workbook(path)
        self.sheet = workbook.active

        self.table_widget.setRowCount(self.sheet.max_row - 1)
        self.table_widget.setColumnCount(self.sheet.max_column)

        list_values = list(self.sheet.values)
        self.table_widget.setHorizontalHeaderLabels(list_values[0])

        row_index = 0
        for value_tuple in list_values[1:]:
            column_index = 0
            for value in value_tuple:
                self.table_widget.setItem(row_index, column_index,QTableWidgetItem(str(value)))
                column_index += 1
            row_index += 1

    def make_checkboxes(self):
        list_values = list(self.sheet.values)
        legend = list(list_values[0][1:])
        for i in self.w1: self.boxes1.removeWidget(i)
        self.w1 =[]
        for i in range(len(legend)):
            self.w1.append(QCheckBox(legend[i], self))
            self.boxes1.addWidget(self.w1[i])


app = QApplication(sys.argv)
window = Main()
window.showMaximized()
app.exec_()