#!/uer/bin/python
# -*- coding: utf-8 -*-

import email
from PyQt5 import QtWidgets
from PyQt5.QtCore import QDir
from PyQt5.QtWidgets import QFileDialog
from main_ui import Ui_Form
import sys
# import win32com.client
import os
from argostranslate import package, translate

a=0
def CounterA():     #计数器，每次调用自增1
    def f():
        global a
        a = a + 1
        return a
    return f
b=0
def CounterB():     #计数器，每次调用自增1
    def g():
        global b
        b = b + 1
        return b
    return g

class My_Form(Ui_Form,QtWidgets.QWidget):  # 继承自UI_Diglog类，注意我把UI_Dialog放在了untitled.py中，如果你把这个类放在了XXX.py文件中，就应该是XXX.UI_Dialog
    def __init__(self, parent =None):
        super(My_Form, self).__init__()
        super().setupUi(self)  # 调用父类的setupUI函数
        self.label_3.setText("0")
        self.label_4.setText("0")
        self.pushButton.clicked.connect(self.on_btnImportFolder_clicked)
        self.pushButton_2.clicked.connect(self.on_btnImportFolder2_clicked)
        self.pushButton_3.clicked.connect(self.on_btnImportFolder3_clicked)

    def on_btnImportFolder_clicked(self):
        cur_dir = QDir.currentPath()  # 获取当前文件夹路径
        # 选择文件夹
        root_path = QFileDialog.getExistingDirectory(self, '打开文件夹', cur_dir)
        file = open(root_path+"/emlatt.txt", 'w').close()
        self.lineEdit_2.setText(root_path)
        for root, dirs, files in os.walk(root_path):  # root, dirs不能删掉，否则程序报错
            for file_name in files:
                absfile_name = os.path.join(root, file_name)
                # if file_name[-4:] == ".msg":
                #     counter_ = CounterA()  # 邮件数量统计
                #     i = counter_()
                #     self.label_3.setText(str(i))
                #     outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")    #需安装outlook
                #     msg = outlook.OpenSharedItem(absfile_name)  # 解析邮件
                #     if hasattr(msg, "Subject"):
                #         subject = msg.Subject
                #     # get_attachments(msg,file_name, root_path)   #获取附件
                #     count_attachments = msg.Attachments.Count  # 附件数目
                #     attachments = msg.Attachments  # 附件
                #     if count_attachments > 0:
                #         # print(count_attachments)
                #         for att in attachments:
                #             counter_ = CounterB()
                #             j = counter_()
                #             self.label_4.setText(str(j))
                #             # print(att.FileName)
                #             aa=str(absfile_name)+"\t"+str(att.FileName)+"\n"
                #             f = open(root_path + "/emlatt.txt", "a+", encoding="utf-8")
                #             f.writelines(aa)
                #             f.close()

                if file_name[-4:]==".eml":
                    counter_ = CounterA()  # 邮件数量统计
                    i = counter_()
                    self.label_3.setText(str(i))
                    def decode_str(s):  # 字符编码转换
                        value, charset = email.header.decode_header(s)[0]
                        if charset:
                            value = value.decode(charset)
                        return value

                    def get_annex_filename(name):
                        h = email.header.Header(name)
                        dh = email.header.decode_header(h)  # 对附件名称进行解码
                        filename = dh[0][0]
                        if dh[0][1]:
                            filename = decode_str(str(filename, dh[0][1]))  # 将附件名称可读化
                        return filename

                    def get_annex(emlfile):
                        try:
                            with open(emlfile, 'rb') as eml:
                                msg = email.message_from_binary_file(eml)
                                for part in msg.walk():
                                    if not part.is_multipart():
                                        name = part.get_filename()
                                        if name:
                                            # print(get_annex_filename(name))
                                            counter_ = CounterB()
                                            j = counter_()
                                            self.label_4.setText(str(j))
                                            bb = str(absfile_name) + "\t"+str(get_annex_filename(name))+"\n"
                                            f3 = open(root_path + "/emlatt.txt", "a+", encoding="utf-8")
                                            f3.writelines(bb)
                                            f3.close()
                        except Exception as e:
                            print(e)
                    get_annex(absfile_name)


    def on_btnImportFolder2_clicked(self):
        Extname=""
        l2=[]
        if self.radioButton.isChecked():
            Extname="xls"
        if self.radioButton_2.isChecked():
            Extname="pdf"
        if self.radioButton_3.isChecked():
            Extname="doc"
        if self.radioButton_4.isChecked():
            Extname="egm"
        if self.radioButton_5.isChecked():
            Extname=self.lineEdit.text()
        self.listWidget.clear()
        f1 = open(self.lineEdit_2.text() + "/emlatt.txt", "r", encoding="utf-8")
        a1 = list(f1)
        for i1 in a1:
            if i1.find(Extname) != -1:
                l2.append(i1[:-1])
        f1.close()
        self.listWidget.addItems(l2)


    def on_btnImportFolder3_clicked(self):
        self.label_6.setText("正在翻译")
        package.install_from_path('argostranslate/translate-en_zh-1_1.argosmodel')  # 这里的模型地址，根据自己放置模型地址去写即可
        installed_languages = translate.get_installed_languages()
        translation_en_es = installed_languages[0].get_translation(installed_languages[1])
        # print(translation_en_es)
        f = open(self.lineEdit_2.text() + "\emlatt.txt", "r", encoding="utf-8")
        lines=list(f)
        lines1=lines
        f.close()
        j=0
        for i in lines:
            b1 = i.find("\t")
            aa = translation_en_es.translate(i[b1:])
            #print(aa)
            lines1[j]=i[:-1]+"\t"+aa+"\n"
            j=j+1
        f = open(self.lineEdit_2.text() + "\emlatt.txt", "w", encoding="utf-8")
        f.writelines(lines1)
        f.close()
        self.label_6.setText("翻译完成")

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = My_Form()
      # 注意把类名修改为myDialog
    # ui.setupUi(MainWindow)  myDialog类的构造函数已经调用了这个函数，这行代码可以删去
    MainWindow.show()
    sys.exit(app.exec_())
