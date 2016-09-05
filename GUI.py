import sys, webbrowser
from PyQt4 import QtGui, QtCore

url = 'https://www.google.com'

class Window(QtGui.QMainWindow):

    def __init__(self):
        super(Window, self).__init__()
        self.setGeometry(50, 50, 500, 300)
        self.setWindowTitle("Network/System Admin Tools")
        #self.setWindowIcon(QtGui.QIcon('pythonlogo.png'))
        self.home()

    def home(self):
		#Quit Button
        Quit_btn = QtGui.QPushButton("Quit", self)
        Quit_btn.clicked.connect(QtCore.QCoreApplication.instance().quit)
        Quit_btn.resize(100,50)
        Quit_btn.move(0,250)
		#PRTG Button
        prtg_btn = QtGui.QPushButton("PRTG Web Console", self)
        prtg_btn.clicked.connect(self.prtg)
        prtg_btn.resize(100,50)
        prtg_btn.move(0,0)
		#Show
        self.show()
	
	def prtg(self):
		webbrowser(url, 2)
        
def run():
    app = QtGui.QApplication(sys.argv)
    GUI = Window()
    sys.exit(app.exec_())

run()
