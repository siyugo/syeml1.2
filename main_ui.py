# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'main.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(848, 779)
        self.listWidget = QtWidgets.QListWidget(Form)
        self.listWidget.setGeometry(QtCore.QRect(30, 120, 791, 611))
        self.listWidget.setObjectName("listWidget")
        self.layoutWidget = QtWidgets.QWidget(Form)
        self.layoutWidget.setGeometry(QtCore.QRect(30, 80, 791, 32))
        self.layoutWidget.setObjectName("layoutWidget")
        self.horizontalLayout_6 = QtWidgets.QHBoxLayout(self.layoutWidget)
        self.horizontalLayout_6.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_6.setObjectName("horizontalLayout_6")
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_5 = QtWidgets.QLabel(self.layoutWidget)
        self.label_5.setObjectName("label_5")
        self.horizontalLayout_4.addWidget(self.label_5)
        self.radioButton = QtWidgets.QRadioButton(self.layoutWidget)
        self.radioButton.setChecked(True)
        self.radioButton.setObjectName("radioButton")
        self.horizontalLayout_4.addWidget(self.radioButton)
        self.radioButton_2 = QtWidgets.QRadioButton(self.layoutWidget)
        self.radioButton_2.setObjectName("radioButton_2")
        self.horizontalLayout_4.addWidget(self.radioButton_2)
        self.radioButton_3 = QtWidgets.QRadioButton(self.layoutWidget)
        self.radioButton_3.setObjectName("radioButton_3")
        self.horizontalLayout_4.addWidget(self.radioButton_3)
        self.radioButton_4 = QtWidgets.QRadioButton(self.layoutWidget)
        self.radioButton_4.setObjectName("radioButton_4")
        self.horizontalLayout_4.addWidget(self.radioButton_4)
        self.radioButton_5 = QtWidgets.QRadioButton(self.layoutWidget)
        self.radioButton_5.setObjectName("radioButton_5")
        self.horizontalLayout_4.addWidget(self.radioButton_5)
        self.horizontalLayout_6.addLayout(self.horizontalLayout_4)
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.lineEdit = QtWidgets.QLineEdit(self.layoutWidget)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout_5.addWidget(self.lineEdit)
        self.pushButton_2 = QtWidgets.QPushButton(self.layoutWidget)
        self.pushButton_2.setObjectName("pushButton_2")
        self.horizontalLayout_5.addWidget(self.pushButton_2)
        self.horizontalLayout_6.addLayout(self.horizontalLayout_5)
        self.layoutWidget1 = QtWidgets.QWidget(Form)
        self.layoutWidget1.setGeometry(QtCore.QRect(31, 20, 441, 30))
        self.layoutWidget1.setObjectName("layoutWidget1")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.layoutWidget1)
        self.horizontalLayout_3.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.pushButton = QtWidgets.QPushButton(self.layoutWidget1)
        self.pushButton.setObjectName("pushButton")
        self.horizontalLayout_3.addWidget(self.pushButton)
        self.lineEdit_2 = QtWidgets.QLineEdit(self.layoutWidget1)
        self.lineEdit_2.setEnabled(True)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.horizontalLayout_3.addWidget(self.lineEdit_2)
        self.layoutWidget2 = QtWidgets.QWidget(Form)
        self.layoutWidget2.setGeometry(QtCore.QRect(480, 20, 341, 31))
        self.layoutWidget2.setObjectName("layoutWidget2")
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout(self.layoutWidget2)
        self.horizontalLayout_7.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label = QtWidgets.QLabel(self.layoutWidget2)
        self.label.setObjectName("label")
        self.horizontalLayout_2.addWidget(self.label)
        self.label_3 = QtWidgets.QLabel(self.layoutWidget2)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_2.addWidget(self.label_3)
        self.horizontalLayout_7.addLayout(self.horizontalLayout_2)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label_2 = QtWidgets.QLabel(self.layoutWidget2)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout.addWidget(self.label_2)
        self.label_4 = QtWidgets.QLabel(self.layoutWidget2)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout.addWidget(self.label_4)
        self.horizontalLayout_7.addLayout(self.horizontalLayout)
        self.layoutWidget3 = QtWidgets.QWidget(Form)
        self.layoutWidget3.setGeometry(QtCore.QRect(560, 740, 261, 30))
        self.layoutWidget3.setObjectName("layoutWidget3")
        self.horizontalLayout_8 = QtWidgets.QHBoxLayout(self.layoutWidget3)
        self.horizontalLayout_8.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_8.setObjectName("horizontalLayout_8")
        self.pushButton_3 = QtWidgets.QPushButton(self.layoutWidget3)
        self.pushButton_3.setObjectName("pushButton_3")
        self.horizontalLayout_8.addWidget(self.pushButton_3)
        self.label_6 = QtWidgets.QLabel(self.layoutWidget3)
        self.label_6.setText("")
        self.label_6.setObjectName("label_6")
        self.horizontalLayout_8.addWidget(self.label_6)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "sy邮件附件查看器（支持eml格式）"))
        self.label_5.setText(_translate("Form", "选择查看的附件："))
        self.radioButton.setText(_translate("Form", "excel"))
        self.radioButton_2.setText(_translate("Form", "pdf"))
        self.radioButton_3.setText(_translate("Form", "word"))
        self.radioButton_4.setText(_translate("Form", "egm"))
        self.radioButton_5.setText(_translate("Form", "其他"))
        self.pushButton_2.setText(_translate("Form", "确认"))
        self.pushButton.setText(_translate("Form", "邮件导入"))
        self.label.setText(_translate("Form", "邮件总数："))
        self.label_3.setText(_translate("Form", "TextLabel"))
        self.label_2.setText(_translate("Form", "附件数量："))
        self.label_4.setText(_translate("Form", "TextLabel"))
        self.pushButton_3.setText(_translate("Form", "附件名翻译英->中"))