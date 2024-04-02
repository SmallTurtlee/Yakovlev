from PyQt5.QtWidgets import     QApplication, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem, QPushButton, QHBoxLayout
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot
import sys
import openpyxl
from PyQt5 import QtCore, QtGui, QtWidgets
import sys
import matplotlib
matplotlib.use('Qt5Agg')

from PyQt5 import QtCore, QtWidgets

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
from matplotlib.figure import Figure

class MplCanvas(FigureCanvasQTAgg):

    def __init__(self, parent=None, width=5, height=4, dpi=100):
        fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = fig.add_subplot(111)
        super(MplCanvas, self).__init__(fig)

class Main(QWidget):
    def __init__(self):
        super(Main, self).__init__()
        self.setWindowTitle("Loading Excel")
        layout = QVBoxLayout()
        block = QHBoxLayout()
        self.setLayout(layout)
        self.button1 = QPushButton("Open new file")
        self.button2 = QPushButton("Create plotting")
        self.sc = MplCanvas(self, width=5, height=4, dpi=100)
        self.button1.clicked.connect(self.on_click)
        self.button2.clicked.connect(self.on_click2)
        self.table_widget = QTableWidget()
        block.addWidget(self.button1)
        block.addWidget(self.button2)
        layout.addLayout(block)
        layout.addWidget(self.sc)
        layout.addWidget(self.table_widget)   


        
    @pyqtSlot()
    def on_click(self):
        name, done1 = QtWidgets.QInputDialog.getText(
			self, 'Input Dialog', 'Введите название или путь к нужному файлу:') 
        if done1: 
            self.load_data(name)
            


    def on_click2(self):
        self.sc.axes.clear()
        list_values = list(self.sheet.values)
        legend = list(list_values[0][1:])
        Main_sheet = list(reversed(list(zip(*list_values[1:]))))
        for i in range(len(legend)):
            self.sc.axes.plot(Main_sheet[-1], Main_sheet[-1*(2+i)])
        self.sc.draw()
        

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

app = QApplication(sys.argv)
window = Main()
window.showMaximized()
app.exec_()