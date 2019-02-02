import os
import xlsxwriter
from PyQt5 import QtCore, QtWidgets
import sys
import threading
import cv2


class Ui_Form(object):

    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(683, 210)
        self.verticalLayout = QtWidgets.QVBoxLayout(Form)
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(Form)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.lineEdit = QtWidgets.QLineEdit(Form)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout.addWidget(self.lineEdit)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_2.addWidget(self.label_2)
        self.lineEdit_2 = QtWidgets.QLineEdit(Form)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.horizontalLayout_2.addWidget(self.lineEdit_2)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.label_3 = QtWidgets.QLabel(Form)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_3.addWidget(self.label_3)
        self.lineEdit_3 = QtWidgets.QLineEdit(Form)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.horizontalLayout_3.addWidget(self.lineEdit_3)
        self.label_4 = QtWidgets.QLabel(Form)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_3.addWidget(self.label_4)
        self.lineEdit_4 = QtWidgets.QLineEdit(Form)
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.horizontalLayout_3.addWidget(self.lineEdit_4)
        self.verticalLayout.addLayout(self.horizontalLayout_3)
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_5 = QtWidgets.QLabel(Form)
        self.label_5.setObjectName("label_5")
        self.horizontalLayout_4.addWidget(self.label_5)
        self.lineEdit_5 = QtWidgets.QLineEdit(Form)
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.horizontalLayout_4.addWidget(self.lineEdit_5)
        self.verticalLayout.addLayout(self.horizontalLayout_4)
        self.label_6 = QtWidgets.QLabel(Form)
        self.label_6.setText("")
        self.label_6.setObjectName("label_6")
        self.verticalLayout.addWidget(self.label_6)
        self.label_7 = QtWidgets.QLabel(Form)
        self.label_7.setText("")
        self.label_7.setObjectName("label_7")
        self.verticalLayout.addWidget(self.label_7)
        self.label_8 = QtWidgets.QLabel(Form)
        self.label_8.setText("")
        self.label_8.setObjectName("label_8")
        self.verticalLayout.addWidget(self.label_8)
        self.label_9 = QtWidgets.QLabel(Form)
        self.label_9.setText("")
        self.label_9.setObjectName("label_9")
        self.verticalLayout.addWidget(self.label_9)
        self.label_10 = QtWidgets.QLabel(Form)
        self.label_10.setText("")
        self.label_10.setObjectName("label_10")
        self.verticalLayout.addWidget(self.label_10)
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setEnabled(True)
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.main)
        self.verticalLayout.addWidget(self.pushButton)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "视频素材归档"))
        self.label.setText(_translate("Form", "素材文件夹路径："))
        self.label_2.setText(_translate("Form", "视频文件类型："))
        self.lineEdit_2.setText(_translate("Form", "mov mp4"))
        self.lineEdit_2.setPlaceholderText(_translate("Form", "要搜索的视频后缀名，多个后缀用空格隔开，不区分大小写"))
        self.label_3.setText(_translate("Form", "截取图片高度："))
        self.lineEdit_3.setText(_translate("Form", "135"))
        self.label_4.setText(_translate("Form", "截取图片宽度："))
        self.lineEdit_4.setText(_translate("Form", "240"))
        self.label_5.setText(_translate("Form", "Excel文件保存路径："))
        self.lineEdit_5.setPlaceholderText(_translate("Form", "不需要写.xlsx后缀名"))
        self.pushButton.setText(_translate("Form", "开始"))



    def noEnabled(self):
        self.pushButton.setEnabled(False)
        self.lineEdit.setEnabled(False)
        self.lineEdit_2.setEnabled(False)
        self.lineEdit_3.setEnabled(False)
        self.lineEdit_4.setEnabled(False)
        self.lineEdit_5.setEnabled(False)

    def enabled(self):
        self.pushButton.setEnabled(True)
        self.lineEdit.setEnabled(True)
        self.lineEdit_2.setEnabled(True)
        self.lineEdit_3.setEnabled(True)
        self.lineEdit_4.setEnabled(True)
        self.lineEdit_5.setEnabled(True)



    def main(self):

        t = threading.Thread(target=finishing)
        t.setDaemon(True)
        t.start()







def path_name(p):
    suf = ui.lineEdit_2.text().lower().split()
    for i in os.listdir(p):
        new_path = os.path.join(p,i)
        if os.path.isdir(new_path):
            path_name(new_path)
        else:
            if new_path[-3:].lower() in suf:
                f_name.append(new_path)





suffix = ["mov"]
f_name = []





def finishing():
    try:
        ui.noEnabled()

        PATH = os.path.normpath(ui.lineEdit.text()) #获取素材路径
        suffix = ui.lineEdit_2.text().lower().split() #获取视频后缀名
        xlsx_path = os.path.normpath(ui.lineEdit_5.text())  #获取Excel文件保存路径



        ui.label_6.setText("正在查找视频文件")
        path_name(PATH)#递归查找视频文件
        ui.label_6.setText("找到 "+ str(len(f_name))+" 个视频文件")



        test_jpg = os.path.normpath(os.path.expanduser("~") + "\\test_jpg")

        if not os.path.exists(test_jpg):
            os.makedirs(test_jpg)#创建临时文件夹

        if not os.path.exists(os.path.dirname(xlsx_path)):
            os.makedirs(os.path.dirname(xlsx_path))#创建Excel文件夹



        xlsx = xlsxwriter.Workbook(xlsx_path + ".xlsx")
        sheet = xlsx.add_worksheet()#创建Excel文档


        style = xlsx.add_format()
        style.set_align('center')
        style.set_align('vcenter')#设置Excel格式

        p_height = ui.lineEdit_3.text()
        p_width = ui.lineEdit_4.text()
        xlsx_height = (int(p_height)+20)/1.3
        xlsx_width = (int(p_width)+30)/8.4*1.15+1



        head_style = xlsx.add_format({
            'bold': True,
            'border': 6,
            'align': 'center',
            'valign': 'vcenter',
        })

        style = xlsx.add_format()
        style.set_align('center')
        style.set_align('vcenter')

        sheet.set_column(0, 14, xlsx_width, head_style)

        head = ["文件名", "文件路径", "文件格式", "文件总帧数", "视频高度宽度"]

        for i in range(0, 5):
            sheet.write_string(0, i, head[i], head_style)

        sheet.merge_range("F1:O1", '视频截图', head_style)

        Number_of_files = str(len(f_name))

        y = 1  # 写入数据高度，从第二行开始
        mov_number = 1

        for mov in f_name:


            ui.label_7.setText("正在读取文件 " + os.path.split(mov)[1] + " [当前第 "+str(mov_number) +  " 个，共计 "+Number_of_files+" 个]")

            vc = cv2.VideoCapture(mov)
            if vc.isOpened():
                is_open,frame = vc.read()
            else:
                is_open = False

            if is_open:
                all_frame = int(vc.get(7))
            else:
                all_frame = "视频无法读取"

            if all_frame<11:
                continue

            if is_open:
                v_height = str(int(vc.get(4)))
            else:
                v_height = "视频无法读取"

            if is_open:
                v_width = str(int(vc.get(3)))
            else:
                v_width = "视频无法读取"

            if is_open:
                Intercept_interval = int(all_frame/11)

            if is_open:
                Screenshot_frames = []
                for i in range(1,11):
                    Screenshot_frames.append(Intercept_interval*i)
            is_excel = False

            while is_open:
                is_excel = True

                if vc.get(1)-1 in Screenshot_frames:
                    inage = cv2.resize(frame, (int(p_width), int(p_height)), 0, 0, cv2.INTER_LINEAR)
                    cv2.imencode('.jpg', inage)[1].tofile(test_jpg +"\\"+ str(int(vc.get(1))-1) + "_.jpg")

                is_open, frame = vc.read()


            ui.label_8.setText("正在写入Excel文件 "+ " [当前第 "+str(mov_number) +  " 个，共计 "+Number_of_files+" 个]")

            sheet.set_row(y, xlsx_height,style)

            for i in range(0,5):
                if i == 0:
                    sheet.write_url(y,i,url=mov,string=os.path.split(mov)[1])
                elif i == 1:
                    sheet.write_url(y, i, url=os.path.split(mov)[0], string=os.path.split(mov)[0])
                elif i == 2:
                    sheet.write_string(y,i,string=os.path.splitext(mov)[1][1:])
                elif i == 3:
                    sheet.write_string(y,i,string=str(all_frame))
                elif i == 4:
                    sheet.write_string(y,i,string=v_width + "x" + v_height)

            if is_excel:
                for i in range(5,15):
                    sheet.insert_image(y, i, test_jpg +"\\"+ str(Screenshot_frames[i-5]) + "_.jpg",{'x_offset': 15, 'y_offset': 10})


            mov_number +=1
            y +=1
            vc.release()




        xlsx.close()

        ui.label_9.setText("正在删除临时截图文件")
        for i in os.listdir(test_jpg):
            os.remove(test_jpg+"\\"+i)
        os.rmdir(test_jpg)

        ui.label_7.setText("文件截图完成")
        ui.label_8.setText("Excel文件写入完成")
        ui.label_9.setText("临时截图文件已删除")
        ui.enabled()

    except Exception as ev:
        e_value = "出现错误，任务已终止，错误内容：" + str(ev)
        ui.label_6.setText(e_value)
        ui.label_7.setText(e_value)
        ui.label_8.setText(e_value)
        ui.label_9.setText(e_value)
        ui.enabled()





























if __name__ == '__main__':

    app = QtWidgets.QApplication(sys.argv)

    widget = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(widget)

    widget.show()


    sys.exit(app.exec_())

