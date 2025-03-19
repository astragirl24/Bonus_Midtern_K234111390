from PyQt6.QtWidgets import QMainWindow, QApplication

from bonus_midterm_K234111390.MainWindowEx import MainWindowEx

app=QApplication([])
myWindow=MainWindowEx()
myWindow.setupUi(QMainWindow())
myWindow.show()
app.exec()
