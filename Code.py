# -*- coding: utf-8 -*-
# Form implementation generated from reading ui file 'Get.ui'
# Created by: PyQt5 UI code generator 5.13.2

import os
import xlsxwriter as excel
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QMessageBox


def caution_window(title='Erro', message='Algum erro aconteceu, revise seu procedimento'):
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Warning)
    msg.setWindowTitle(title)
    msg.setText(message)
    msg.exec_()


class UiMainWindow(object):
    def setup_ui(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(640, 480)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setGeometry(QtCore.QRect(80, 140, 101, 21))
        self.textEdit.setObjectName("textEdit")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(80, 90, 161, 16))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(80, 120, 81, 21))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(80, 170, 71, 16))
        self.label_3.setObjectName("label_3")
        self.textEdit_2 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_2.setGeometry(QtCore.QRect(80, 190, 101, 21))
        self.textEdit_2.setObjectName("textEdit_2")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(250, 380, 101, 41))
        self.pushButton.setObjectName("pushButton")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(360, 60, 201, 16))
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(360, 180, 211, 16))
        self.label_5.setObjectName("label_5")
        self.listWidget = QtWidgets.QListWidget(self.centralwidget)
        self.listWidget.setGeometry(QtCore.QRect(360, 240, 191, 111))
        self.listWidget.setObjectName("listWidget")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(480, 200, 81, 31))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(360, 150, 81, 31))
        self.pushButton_3.setObjectName("pushButton_3")
        self.textEdit_3 = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.textEdit_3.setGeometry(QtCore.QRect(360, 90, 201, 50))
        self.textEdit_3.setObjectName("textEdit_3")
        self.textEdit_4 = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.textEdit_4.setGeometry(QtCore.QRect(360, 200, 111, 31))
        self.textEdit_4.setObjectName("textEdit_4")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 640, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.translating_ui(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.pushButton.clicked.connect(self.execute)
        self.pushButton_3.clicked.connect(self.open_file_dialog)
        self.pushButton_2.clicked.connect(self.add_identifier)

    def translating_ui(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label.setText(_translate("MainWindow", "Comprimento da onda procurado"))
        self.label_2.setText(_translate("MainWindow", "Limite Superior"))
        self.label_3.setText(_translate("MainWindow", "Limite Inferior"))
        self.pushButton.setText(_translate("MainWindow", "Executar"))
        self.label_4.setText(_translate("MainWindow", "Escolha a pasta que contém os arquivos"))
        self.label_5.setText(_translate("MainWindow", "Quais são os identificadores dos arquivos ?"))
        self.pushButton_2.setText(_translate("MainWindow", "Adicionar"))
        self.pushButton_3.setText(_translate("MainWindow", "Adicionar"))

    def add_identifier(self):
        new_identifier = str(self.textEdit_4.toPlainText())
        print(new_identifier)
        self.listWidget.addItem(new_identifier)
        self.textEdit_4.setPlainText("")

    def open_file_dialog(self):
        path = QFileDialog.getExistingDirectory()
        print(path)
        self.textEdit_3.setPlainText(path)

    def execute(self):
        path = str(self.textEdit_3.toPlainText())
        data = []
        upper_limit = float(self.textEdit.toPlainText())
        lower_limit = float(self.textEdit_2.toPlainText())
        print(upper_limit, lower_limit)
        print(path)

        if upper_limit < lower_limit:
            caution_window()

        def analysis_time(name):
            return int(name[name.index('_') + 1:name.index('.')])

        list_file_names = list(os.listdir(path))

        try:
            list_file_names = sorted(list_file_names, key=analysis_time)
        except:
            caution_window('Erro', 'Pelo menos um dos seus dados não estão '
                           'no modelo necessário, cheque o resultado')
        print(list_file_names)

        for file_name in list_file_names:
            if 'csv' in file_name:
                file = open(path + "/" + file_name, 'r')
                for lines in file:
                    identifier = file_name
                    try:
                        if upper_limit > float(lines[:5]) > lower_limit:
                            wave_length = float(lines[:6])
                            transmittance = float(lines[lines.index(',') + 1:lines.index(',') + 6])
                            line = [identifier, wave_length, transmittance]
                            data.append(line)
                    except:
                        continue
            else:
                continue
        output = 'saida_' + path[-5:] + '.xlsx'

        user_identifiers = []

        for i in range(0, self.listWidget.count()):
            user_identifiers.append(self.listWidget.item(i).text())
            print(user_identifiers)

        workbook = excel.Workbook(output)
        worksheet = workbook.add_worksheet()
        worksheet.write('A1', 'Tempo')
        worksheet.write('B1', 'Comprimento')

        row = 0
        column = 2

        for identifier in user_identifiers:
            worksheet.write_string(row, column, identifier)
            column += 1

        i = -1
        print(len(user_identifiers))
        for row in range(0, int(len(data) / len(user_identifiers))):
            for column in range(0, len(user_identifiers)):
                i += 1
                print(i, end='')
                if column == 0:
                    worksheet.write_number(row + 1, column, int(data[i][0][data[i][0].index('_')
                                                                           + 1:data[i][0].index('.')]))
                if column == 1:
                    worksheet.write_number(row + 1, column, data[i][1])

        i = -1
        for row in range(0, int(len(data) / len(user_identifiers))):
            for column in range(2, len(user_identifiers) + 2):
                i += 1
                print(i, end='')
                worksheet.write_number(row + 1, column, data[i][2])

        workbook.close()
        caution_window('Sucesso', 'Não se esqueça de checar seu resultado')


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = UiMainWindow()
    ui.setup_ui(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
