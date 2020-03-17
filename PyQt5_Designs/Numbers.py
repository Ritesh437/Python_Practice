from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import QInputDialog
from PyQt5.QtWidgets import QFileDialog
import pyperclip


# Functions Generating Different Numbers

# check if given num is prime or composite
def isPrime(a):
    if(a==0 or a==1):
        return 'notany'
    for i in range (2,int(a/2)+1):
        if(a%i==0):
            return 'c'
    return 'p'

# 1 prime
def prime(a,b):
    p=''
    for i in range (a,b+1):
        x=isPrime(i)
        if(x=='p'):
            p += '  {prime}'.format(prime = i)
    return p

# 2 Composite
def composite(a,b):
    p=''
    for i in range (a,b+1):
        x=isPrime(i)
        if(x=='c'):
            p += '  {composite}'.format(composite = i)
    return p

# 3 odd
def odd(a,b):
    p=''
    for i in range (a,b+1):
        if(i%2==1):
            p += '  {odd}'.format(odd = i)
    return p

# 4 even
def even(a,b):
    p=''
    for i in range (a,b+1):
        if(i==0):
            continue
        if(i%2==0):
            p += '  {even}'.format(even = i)
    return p

# 5 Natural
def natural(a,b):
    p=''
    for i in range (a,b+1):
        if(i==0):
            continue
        p += '  {natural}'.format(natural = i)
    return p


class Ui_MainWindow(object):

    state = {
        'typ': 'Prime',
        'num': ''
    }

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(744, 449)
        MainWindow.setMouseTracking(False)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.spinBox1 = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox1.setGeometry(QtCore.QRect(70, 50, 42, 22))
        self.spinBox1.setMouseTracking(True)
        self.spinBox1.setMaximum(10000)
        self.spinBox1.setObjectName("spinBox1")
        self.label2 = QtWidgets.QLabel(self.centralwidget)
        self.label2.setGeometry(QtCore.QRect(10, 50, 50, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.label2.setFont(font)
        self.label2.setObjectName("label2")
        self.label3 = QtWidgets.QLabel(self.centralwidget)
        self.label3.setGeometry(QtCore.QRect(130, 50, 30, 19))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.label3.setFont(font)
        self.label3.setObjectName("label3")
        self.spinBox2 = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox2.setGeometry(QtCore.QRect(170, 50, 42, 22))
        self.spinBox2.setMouseTracking(True)
        self.spinBox2.setMaximum(100000)
        self.spinBox2.setProperty("value", 100)
        self.spinBox2.setObjectName("spinBox2")
        self.label1 = QtWidgets.QLabel(self.centralwidget)
        self.label1.setGeometry(QtCore.QRect(20, 10, 171, 31))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.label1.setFont(font)
        self.label1.setObjectName("label1")
        self.generate = QtWidgets.QPushButton(self.centralwidget)
        self.generate.setGeometry(QtCore.QRect(230, 50, 75, 23))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.generate.setFont(font)
        self.generate.setObjectName("generate")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(500, 0, 241, 16))
        self.label.setObjectName("label")
        self.label4 = QtWidgets.QTextBrowser(self.centralwidget)
        self.label4.setGeometry(QtCore.QRect(10, 90, 721, 301))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.label4.setFont(font)
        self.label4.setMouseTracking(True)
        self.label4.setObjectName("label4")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 744, 21))
        self.menubar.setObjectName("menubar")
        self.menufile = QtWidgets.QMenu(self.menubar)
        self.menufile.setObjectName("menufile")
        self.menuedit = QtWidgets.QMenu(self.menubar)
        self.menuedit.setObjectName("menuedit")
        self.menuoptions = QtWidgets.QMenu(self.menubar)
        self.menuoptions.setObjectName("menuoptions")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionopen = QtWidgets.QAction(MainWindow)
        self.actionopen.setObjectName("actionopen")
        self.actionsave = QtWidgets.QAction(MainWindow)
        self.actionsave.setObjectName("actionsave")
        self.actioncopy = QtWidgets.QAction(MainWindow)
        self.actioncopy.setObjectName("actioncopy")
        self.actionPrime = QtWidgets.QAction(MainWindow)
        self.actionPrime.setObjectName("actionPrime")
        self.actionComposite = QtWidgets.QAction(MainWindow)
        self.actionComposite.setObjectName("actionComposite")
        self.actionnatural = QtWidgets.QAction(MainWindow)
        self.actionnatural.setObjectName("actionnatural")
        self.actionodd = QtWidgets.QAction(MainWindow)
        self.actionodd.setObjectName("actionodd")
        self.actioneven = QtWidgets.QAction(MainWindow)
        self.actioneven.setObjectName("actioneven")
        self.menufile.addAction(self.actionopen)
        self.menufile.addAction(self.actionsave)
        self.menuedit.addAction(self.actioncopy)
        self.menuoptions.addAction(self.actionPrime)
        self.menuoptions.addAction(self.actionComposite)
        self.menuoptions.addAction(self.actionnatural)
        self.menuoptions.addAction(self.actionodd)
        self.menuoptions.addAction(self.actioneven)
        self.menubar.addAction(self.menufile.menuAction())
        self.menubar.addAction(self.menuedit.menuAction())
        self.menubar.addAction(self.menuoptions.menuAction())

        self.actionComposite.triggered.connect(lambda : self.changeHead('Composite'))
        self.actionPrime.triggered.connect(lambda : self.changeHead('Prime'))
        self.actionodd.triggered.connect(lambda : self.changeHead('Odd'))
        self.actioneven.triggered.connect(lambda : self.changeHead('Even'))
        self.actionnatural.triggered.connect(lambda : self.changeHead('Natural'))

        self.actioncopy.triggered.connect(self.copyToClipboard)
        self.actionsave.triggered.connect(self.save)
        self.actionopen.triggered.connect(self.open)

        self.generate.clicked.connect(self.generateNum)


        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)


    def save(self):
        filename = QFileDialog.getSaveFileName(self.centralwidget,'Single File','d:\'','*.txt')
        file = open(filename[0], 'w')
        file.write(self.state['num'])
        file.close()

    def open(self):
        fileName = QFileDialog.getOpenFileName(self.centralwidget,'Single File','d:\'','*.txt')
        file = open(fileName[0], 'r')
        st = file.read()
        file.close()
        self.state['num'] = st
        self.label4.setText(st)
        

    def copyToClipboard(self):
        pyperclip.copy(self.state['num'])
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("list Copied To Clipboard")
        msg.setWindowTitle("Message")

        x = msg.exec_()


    def changeHead(self,t):
        self.state['typ'] = t
        self.label1.setText('{typ} Numbers'.format(typ = self.state['typ']))
        self.label1.adjustSize()

    def generateNum(self):
        t = self.state['typ']
        a = self.spinBox1.value()
        b = self.spinBox2.value()
        if(t == 'Composite'):
            p= composite(a,b)
        if(t == 'Prime'):
            p= prime(a,b)
        if(t == 'Odd'):
            p= odd(a,b)
        if(t == 'Even'):
            p= even(a,b)
        if(t == 'Natural'):
            p= natural(a,b)
        self.label4.setText(p)
        self.state['num'] = p

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label2.setText(_translate("MainWindow", "From :"))
        self.label3.setText(_translate("MainWindow", "To :"))
        self.label1.setText(_translate("MainWindow", "Prime Numbers List"))
        self.generate.setText(_translate("MainWindow", "Generate"))
        self.label.setText(_translate("MainWindow", "Select Types of numbers from options above"))
        self.label4.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:18pt; font-weight:400; font-style:normal;\">\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:8.25pt;\">Numbers Will Appear Here</span></p></body></html>"))
        self.menufile.setTitle(_translate("MainWindow", "file"))
        self.menuedit.setTitle(_translate("MainWindow", "edit"))
        self.menuoptions.setTitle(_translate("MainWindow", "options"))
        self.actionopen.setText(_translate("MainWindow", "open"))
        self.actionsave.setText(_translate("MainWindow", "save"))
        self.actioncopy.setText(_translate("MainWindow", "copy"))
        self.actionPrime.setText(_translate("MainWindow", "Prime"))
        self.actionComposite.setText(_translate("MainWindow", "Composite"))
        self.actionnatural.setText(_translate("MainWindow", "natural"))
        self.actionodd.setText(_translate("MainWindow", "odd"))
        self.actioneven.setText(_translate("MainWindow", "even"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
