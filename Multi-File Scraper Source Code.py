import PyPDF2
import glob
import time
import pandas as pd
import docx
from PyQt5.QtWidgets import QApplication, QDialog, QVBoxLayout, QLineEdit, QPushButton, QHBoxLayout, QMessageBox, \
    QLCDNumber, QLabel, QWidget, QFileDialog, QListWidget, QListWidgetItem, QCheckBox
import sys
from PyQt5.QtGui import QFont, QIcon, QPixmap
from PyQt5.QtCore import QSize, QTime, QTimer, Qt

word = False
excel = False
pdf = False

a = chr(34)
file_link = ""
files_links = []
current_folder = False
current_folder_path = ""

excel_links = []
pdf_links = []

num = 0


class Window(QWidget):
    def __init__(self):
        super().__init__()

        self.setGeometry(272, 48, 800, 600)
        # self.setFixedHeight(600)
        # self.setFixedWidth(800)
        self.setWindowTitle("\tWord Scraper")
        self.setWindowIcon(QIcon("burger.ico"))

        self.lcd_number()

    def lcd_number(self):

        vbox = QVBoxLayout()
        self.label121212 = QLabel("Word Scraper")
        self.label121212.setAlignment(Qt.AlignCenter)
        self.label121212.setStyleSheet("text-shadow: 1px 2px 2px #1C6EA4;background-color:#19fc78; border: 3px solid rgba(0,0,0,0.1); box-shadow:inset 0 1px 0 rgba(255,255,255,0.5),0 2px 2px rgba(0,0,0,0.3),0 0 4px 1px rgba(0,0,0,0.2),inset 0 3px 2px rgba(255,255,255,.22),inset 0 -3px 2px rgba(0,0,0,.15),inset 0 20px 10px rgba(255,255,255,.12),0 0 4px 1px rgba(0,0,0,.1),0 3px 2px rgba(0,0,0,.2);")
        self.label121212.setFont(QFont("times new roman", 48))
        self.label121212.setFixedHeight(72)
        vbox.addWidget(self.label121212)
        fhand = open('image link.txt')
        fhand = fhand.readlines()[0]
        self.pixmap = QPixmap(fhand)
        self.pixmap2 = self.pixmap.scaled(800, 600, Qt.KeepAspectRatio)
        self.label2 = QLabel()
        self.label2.setPixmap(self.pixmap2)
        self.label2.setStyleSheet("border-radius:5px;")
        self.label2.setAlignment(Qt.AlignCenter)
        vbox.addWidget(self.label2)

        self.label2 = QLabel(
            "\n      Copy the Folder Path and paste in respective box\n"
            "      Folder Path should be without single/double quotes\n"
            )

        self.label2.setStyleSheet("color:red")
        self.label2.setFont(QFont("times new roman", 12))
        self.label2.setFixedHeight(77)

        hbox = QHBoxLayout()

        self.label3 = QLabel(" NOTE  ")
        self.label3.setStyleSheet("color:Red")
        self.label3.setFont(QFont("castellar", 27))
        self.label3.setFixedHeight(72)
        self.label3.setFixedWidth(144)

        hbox.addWidget(self.label3)
        hbox.addWidget(self.label2)

        vbox.addLayout(hbox)

        self.input1 = QLineEdit()
        self.input1.setPlaceholderText("\tEnter the Folder link here")
        self.input1.setFont(QFont("times new roman", 12))
        self.input1.setFixedHeight(60)
        self.input1.setStyleSheet("background-color:white")
        hbox12 = QHBoxLayout()
        hbox12.addWidget(self.input1)

        btn2 = QPushButton(" OPEN ")
        btn2.setFont(QFont("times new roman", 29))
        btn2.setStyleSheet("background-color:yellow")
        btn2.setFixedWidth(120)
        btn2.clicked.connect(self.open)
        # hbox12.addWidget(btn2)

        vbox.addLayout(hbox12)

        hbox1 = QHBoxLayout()
        hbox2 = QHBoxLayout()

        self.input2 = QLineEdit()
        self.input2.setPlaceholderText("\tEnter the Keyword here")
        self.input2.setFont(QFont("times new roman", 12))
        self.input2.setFixedHeight(60)
        self.input2.setStyleSheet("background-color:white")
        hbox1.addWidget(self.input2)

        hbox22 = QHBoxLayout()

        labl = QLabel()
        labl.setMaximumWidth(120)
        hbox22.addWidget(labl)

        self.check1 = QCheckBox("docx")
        self.check1.setFont(QFont("Sanserif", 22))
        self.check1.toggled.connect(self.item_selected)
        hbox22.addWidget(self.check1)

        self.check2 = QCheckBox("XLS")
        self.check2.setFont(QFont("Sanserif", 22))
        self.check2.toggled.connect(self.item_selected)
        hbox22.addWidget(self.check2)

        self.check3 = QCheckBox("PDF")
        self.check3.setFont(QFont("Sanserif", 22))
        self.check3.toggled.connect(self.item_selected)
        hbox22.addWidget(self.check3)

        btn1 = QPushButton(" SCAN 1 ")
        btn1.setFont(QFont("times new roman", 36))
        btn1.setStyleSheet("background-color:pink")
        btn1.clicked.connect(self.read_file)
        # hbox2.addWidget(btn1)

        # hbox22.setAlignment(Qt.AlignCenter)

        btn3 = QPushButton(" SCAN ")
        btn3.setFont(QFont("times new roman", 36))
        # btn3.setMaximumWidth(600)
        btn3.setStyleSheet("text-shadow: 1px 2px 2px #1C6EA4;background-color:violet; border: 3px solid rgba(0,0,0,0.1); box-shadow:inset 0 1px 0 rgba(255,255,255,0.5),0 2px 2px rgba(0,0,0,0.3),0 0 4px 1px rgba(0,0,0,0.2),inset 0 3px 2px rgba(255,255,255,.22),inset 0 -3px 2px rgba(0,0,0,.15),inset 0 20px 10px rgba(255,255,255,.12),0 0 4px 1px rgba(0,0,0,.1),0 3px 2px rgba(0,0,0,.2);")
        btn3.clicked.connect(self.all_files)
        hbox2.addWidget(btn3)

        vbox.addLayout(hbox1)
        vbox.addLayout(hbox22)
        vbox.addLayout(hbox2)

        self.setLayout(vbox)

    def read_file_controller(self):
        global file_link

        a = chr(34)
        b = chr(92)
        c = chr(47)
        d = chr(39)
        splitted = file_link.split(c)
        file_link = splitted[0] + c + splitted[1] + b + splitted[2]
        self.read_file(self, file_link)

    def open(self):
        global file_link
        path = QFileDialog.getOpenFileName(self, 'Open a file', '',
                                           'All Files (*.*)')
        if path != ('', ''):
            file_link = path[0]
            self.input1.setPlaceholderText(file_link)

    def all_files(self):
        global file_link
        global files_links
        global current_folder
        global current_folder_path
        global excel_links
        global pdf_links

        print(file_link)
        b = chr(92)
        c = chr(47)
        d = chr(39)
        folder_link = self.input1.text().lstrip().rstrip()
        self.keyword = self.input2.text().lstrip().rstrip().lower()

        if current_folder == False:
            if len(file_link) < 1:
                if len(folder_link) < 1:
                    return
                else:
                    folder_link = self.input1.text().lstrip().rstrip()
            else:
                if '.' in file_link:
                    if file_link[2] == b:
                        temp = file_link.split(b)
                        temp2 = temp.pop(len(temp) - 1)
                        folder_link = b.join(temp)
                    elif file_link[2] == c:
                        temp = file_link.split(c)
                        temp2 = temp.pop(len(temp) - 1)
                        folder_link = c.join(temp)

            if '.' in file_link:
                path = folder_link
                self.input1.setPlaceholderText(path)
            elif len(file_link) < 1:
                path = self.input1.text().lstrip().rstrip()
                if path[0] == a:
                    stripper = path.split(a)
                    path = stripper[1]
                elif path[0] == d:
                    stripper = path.split(d)
                    path = stripper[1]
                else:
                    path = self.input1.text().lstrip().rstrip()
            print(path)

            if len(path) < 1:
                return

            path3 = path + '/*.docx'
            path4 = path + '/*.xlsx'
            path5 = path + '/*.pdf'

        else:
            path3 = current_folder_path + '/*.docx'
            path4 = current_folder_path + '/*.xlsx'
            path5 = current_folder_path + '/*.pdf'

        files = glob.glob(path3)
        print(files)
        for i in files:
            file_link = i
            files_links.append(i)

        excel_files = glob.glob(path4)
        print(excel_files)
        for j in excel_files:
            excel_links.append(j)

        pdf_files = glob.glob(path5)
        print(pdf_files)
        for k in pdf_files:
            pdf_links.append(k)

        self.read_file()


    def item_selected(self):
        global word
        global excel
        global pdf

        word = False
        excel = False
        pdf = False

        value = ""

        if self.check1.isChecked():
            value = value + "\t" +  self.check1.text()
            word = True

        if self.check2.isChecked():
            value = value + "\t" + self.check2.text()
            excel = True

        if self.check3.isChecked():
            value = value + "\t" +  self.check3.text()
            pdf = True

        if value == "":
            value = "Nothing Selected"
        else:
            value = "You have selected:\t" + value
            print(f"word: {word}, \t excel: {excel}, \t pdf: {pdf}")

        print(value)

    def scan_excel(self):
        global num
        global excel_links
        print(excel_links)
        print("Excel Function working fine")

        a = chr(34)
        b = chr(92)
        c = chr(47)
        d = chr(39)
        text = ' '


        for file in excel_links:

            name_printed = False
            self.label.insertItem(num, text)
            self.setFont(QFont("times new roman", 12))
            self.setStyleSheet("background-color:white")

            num += 1

            text = QListWidgetItem()
            name = file.split(b)[-1]
            text.setText("\n" + name + "\n")
            text.setFont(QFont("times new roman", 24))
            num1212 = num

            num += 1

            word_list = self.keyword.split()

            print("Wordlist is :  ", word_list)

            df = pd.read_excel(file)
            keys = df.keys()
            records = df.count(0).values.max()

            i = 1
            for record in range(records):
                for column in keys:
                    printed = False
                    for word in word_list:
                        # print(record, column, word)
                        if word in str(df.iloc[record][column]).lower():
                            print(i)
                            print(record, column)
                            print(df.iloc[record][column])
                            i += 1

                            if printed == False:
                                if name_printed == False:
                                    self.label.insertItem(num1212, text)
                                    name_printed = True

                                text = "\n\n\nRow \t" + str(record + 1) + f"\n Column : \t{column}\n\t{str(df.iloc[record][column])}"
                                self.label.insertItem(num, text)
                                num += 1
                                printed = True

                            text = "\n" + "***** " + word + " *****\tfound at this location / Cell "
                            self.label.insertItem(num, text)
                            num += 1
            text = "\n\n\n\n\n"
            self.label.insertItem(num, text)
            num += 1



    def scan_pdf(self):
        global pdf_links
        global num

        a = chr(34)
        b = chr(92)
        c = chr(47)
        d = chr(39)
        text = ' '

        print("PDF Function working fine")
        i = 1

        for link in pdf_links:

            name_printed = False
            self.label.insertItem(num, text)
            self.setFont(QFont("times new roman", 12))
            self.setStyleSheet("background-color:white")

            num += 1

            text = QListWidgetItem()
            name = link.split(b)[-1]
            text.setText("\n" + name + "\n")
            text.setFont(QFont("times new roman", 24))
            num1212 = num

            num += 1

            word_list = self.keyword.split()

            print("Wordlist is :  ", word_list)


            pdfFileObj = open(link, 'rb')
            print("pdf 1")
            try:
                pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
            except:
                continue
            print("pdf 2")

            pages = pdfReader.numPages
            print("pdf 3")


            for page_num in range(pages):
                pageObj = pdfReader.getPage(page_num)
                print("pdf 4")

                page_content = pageObj.extractText()
                print("pdf 5")

                lines = page_content.split('\n')
                print("pdf 6")

                for count, line in enumerate(lines):
                    printed = False
                    print("pdf 7")
                    for word in word_list:
                        # print(record, column, word)
                        print("pdf 8")

                        if word in str(line.lower()):
                            print(i)
                            print(word)
                            print(line)
                            i += 1

                            if printed == False:
                                if name_printed == False:
                                    self.label.insertItem(num1212, text)
                                    name_printed = True

                                text = "\n\n\nPage # \t" + str(
                                    page_num + 1) + f"\n Line # \t{count + 1}\n\t{line}"
                                self.label.insertItem(num, text)
                                num += 1
                                printed = True

                            text = "\n" + "***** " + word + " *****\tfound at this location / Cell "
                            self.label.insertItem(num, text)
                            num += 1
                text = "\n\n\n\n\n"
                self.label.insertItem(num, text)
                num += 1




            pdfFileObj.close()

    def read_file(self):
        global num
        global file_link
        global files_links
        global excel_links
        global pdf_links
        a = chr(34)
        b = chr(92)
        c = chr(47)
        d = chr(39)
        print('1', file_link)
        sub_link = self.input1.text().lstrip().rstrip()
        self.keyword = self.input2.text().lstrip().rstrip().lower()
        if len(file_link) < 1:
            if len(sub_link) < 1:
                return
            else:
                stripper = sub_link.split(a)

                sub_link = stripper[0]

        else:
            sub_link = file_link
            print(sub_link)

        if sub_link.split('.')[-1] != 'docx':
            self.label2.setText(
                "\tFile Type not supported\n\tOnly MS Word (.docx) files are supported\n\tKindly select a MS Word file to Proceed")
            self.label2.setFont(QFont("times new roman", 16))
            return
        print(2, sub_link)
        if len(files_links) < 1:
            files_links.append(sub_link)

        settings_dialog = QDialog()
        settings_dialog.setModal(True)
        settings_dialog.setStyleSheet("background-color:white")
        settings_dialog.setWindowTitle("\ttext file")
        settings_dialog.setGeometry(35, 50, 1300, 660)
        # settings_dialog.showFullScreen()
        vbox_layout = QVBoxLayout()

        self.label = QListWidget()
        num = 0
        text = ""

        for link in files_links:
            name_printed = False
            self.label.insertItem(num, text)
            self.setFont(QFont("times new roman", 12))
            self.setStyleSheet("background-color:white")

            num += 1

            text = QListWidgetItem()
            name = link.split(b)[-1]
            text.setText("\n" + name + "\n")
            text.setFont(QFont("times new roman", 24))
            num1212 = num

            num += 1
            doc = docx.Document(link)

            for z, i in enumerate(doc.paragraphs):
                printed = False
                count = i.text.lower().count(self.keyword)
                data = i.text.lower().split()
                word_list = self.keyword.split()
                for word in word_list:
                    address = []
                    for j, index in enumerate(data):
                        if word in index:
                            address.append(j + 1)

                    # print(count, "\t#\t algorithm")
                    if len(address) > 0:
                        if printed == False:
                            if name_printed == False:
                                self.label.insertItem(num1212, text)
                                name_printed = True
                            text = "\n\n\nParagraph " + str(z + 1) + "\n" + i.text.lstrip().rstrip()
                            self.label.insertItem(num, text)
                            num += 1
                            printed = True
                        g = ", "
                        address = [str(locat) for locat in address]
                        text = "\n" + "***** " + word + " *****\tfound at these locations\t" + g.join(address)
                        self.label.insertItem(num, text)
                        num += 1
            text = "\n\n\n\n\n"
            self.label.insertItem(num, text)
            num += 1
        if excel:
            self.scan_excel()

        if pdf:
            self.scan_pdf()
        vbox_layout.addWidget(self.label)

        settings_dialog.setLayout(vbox_layout)
        settings_dialog.exec_()
        files_links = []
        self.label = 121212
        excel_links = []
        pdf_links = []


app = QApplication(sys.argv)
window = Window()
window.show()
sys.exit(app.exec_())
'''1 E:/copywriting project\kp20220309_revised.docx
E:/copywriting project\kp20220309_revised.docx
2 E:/copywriting project\kp20220309_revised.docx'''
