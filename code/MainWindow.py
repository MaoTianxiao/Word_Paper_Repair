from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from Title import Title
from MainBody import MainBody
from docx import Document

class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setWindowTitle("Word Format V0.1")
        self.setWindowIcon(QIcon("./image/icon.png"))
        
        # 加入菜单栏（运行）
        bar = self.menuBar()
        file = bar.addMenu("&File(F)")
        openAction = QAction(QIcon("./image/open.png"), "Choose File", self)#打开操作
        openAction.setShortcut("Ctrl+O")
        quitAction = QAction(QIcon("./image/quit.png"), "Quit", self)#退出操作
        file.addAction(openAction)
        file.addAction(quitAction)

        run = bar.addMenu("&Run(R)")
        runAction = QAction(QIcon('./image/run.png'), "Run", self) # 运行操作
        runAction.setShortcut("F5")
        run.addAction(runAction)

        # 全局变量
        self.filename = ""

        # 界面设置
        tips = QLabel(self)
        tips.setText("<p><font size=50 face=arial color=red>选择需要调整的格式</font></p>")
        ok_bt = QPushButton("勾选完成")
        mainFirstLayout = QHBoxLayout()
        mainFirstLayout.addStretch(2)
        mainFirstLayout.addWidget(tips,1)
        mainFirstLayout.addWidget(ok_bt,1)
        mainFirstLayout.addStretch(2)

        # 标题页
        title = QLabel(self)
        title.setText("<p><font size=20 face=arial color=black>标题</font></p>")
        self.titleBoxButton = []
        titleGroupBox = QGroupBox("titles")
        firstTitleLineLayout = QHBoxLayout()
        secondTitleLineLayout = QHBoxLayout()
        for i in range(1,9):
            bt = QCheckBox(str(i) + "级标题")
            self.titleBoxButton.append(bt)
        for i in range(0,4):
            firstTitleLineLayout.addWidget(self.titleBoxButton[i])
        for i in range(4,8):
            secondTitleLineLayout.addWidget(self.titleBoxButton[i])
        titleGroupLayout = QVBoxLayout()
        titleGroupLayout.addLayout(firstTitleLineLayout)
        titleGroupLayout.addLayout(secondTitleLineLayout)
        titleGroupBox.setLayout(titleGroupLayout)
        mainSecondLayout = QHBoxLayout()
        mainSecondLayout.addStretch(5)
        mainSecondLayout.addWidget(title,1)
        mainSecondLayout.addWidget(titleGroupBox,4)
        mainSecondLayout.addStretch(5)

        # 正文页
        self.mainBodyFlag = 1 # 是否设置正文 
        mainBody = QLabel(self)
        mainBody.setText("<p><font size=20 face=arial color=black>正文</font></p>")
        self.mainBodyBt = QCheckBox("正文设置")
        self.mainBodyBt.setChecked(True)
        mainThreeLayout = QHBoxLayout()
        mainThreeLayout.addStretch(5)
        mainThreeLayout.addWidget(mainBody,1)
        mainThreeLayout.addWidget(self.mainBodyBt,4)
        mainThreeLayout.addStretch(5)

        # 图名设置
        self.picFlag = 0 # 是否设置
        self.picStart = "" # 图片开始的标志
        pic = QLabel(self)
        pic.setText("<p><font size=20 face=arial color=black>图片</font></p>")
        self.picBt = QCheckBox("图名设置")
        picLabel = QLabel(self)
        picLabel.setText('图名开始的字符:')
        self.picLine = QLineEdit(self)
        self.picLine.setText('图 ')
        mainForthLayout = QHBoxLayout()
        mainForthLayout.addStretch(5)
        mainForthLayout.addWidget(pic,1)
        mainForthLayout.addWidget(self.picBt,1)
        mainForthLayout.addWidget(picLabel,1)
        mainForthLayout.addWidget(self.picLine,2)
        mainForthLayout.addStretch(5)

        # 表名设置
        self.tbFlag = 0 # 是否设置
        self.tbStart = "" # 表格开始的标志
        tb = QLabel(self)
        tb.setText("<p><font size=20 face=arial color=black>表格</font></p>")
        self.tbBt = QCheckBox("表名设置")
        tbLabel = QLabel(self)
        tbLabel.setText('表名开始的字符:')
        self.tbLine = QLineEdit(self)
        self.tbLine.setText('表 ')
        mainFifthLayout = QHBoxLayout()
        mainFifthLayout.addStretch(5)
        mainFifthLayout.addWidget(tb,1)
        mainFifthLayout.addWidget(self.tbBt,1)
        mainFifthLayout.addWidget(tbLabel,1)
        mainFifthLayout.addWidget(self.tbLine,2)
        mainFifthLayout.addStretch(5)

        #主页面layout
        mainLayout = QVBoxLayout()
        mainLayout.addStretch(2)
        mainLayout.addLayout(mainFirstLayout,1)
        mainLayout.addLayout(mainThreeLayout,1)
        mainLayout.addLayout(mainForthLayout,1)
        mainLayout.addLayout(mainFifthLayout,1)
        mainLayout.addLayout(mainSecondLayout,2)
        mainLayout.addStretch(15)

        start = QWidget()
        start.setLayout(mainLayout)

        self.tab = QTabWidget() # 主标签页
        self.tab.addTab(start,"start")

        #状态栏提示语
        self.status = self.statusBar()
        #设置主页面
        self.setCentralWidget(self.tab)
        self.showMaximized() #初始最大化显示


        #设置信号槽
        quitAction.triggered.connect(self.QuitProgram) # 退出
        openAction.triggered.connect(self.ChooseFile) # 选择文件
        runAction.triggered.connect(self.Run) # 运行
        ok_bt.clicked.connect(self.ChooseOver) # 勾选完成

    def QuitProgram(self):
        self.close()

    def ChooseFile(self):
        self.filename = QFileDialog.getOpenFileName(self, 'Choose File', './', '*.docx')[0]

    def Run(self):
        # 是否选择文件
        if self.filename == "":
            QMessageBox.information(self, "Attention", "Please Choose File")#啥也没选
            return
        
        # 获取配置
        ## 正文
        if self.mainBodyFlag == 1:
            ms = self.mainSet.getSetter()
        if self.picFlag == 1:
            ps = self.picSet.getSetter()
        if self.tbFlag == 1:
            tbs = self.tbSet.getSetter()
        ## Title
        self.ts = []
        for title in self.titleSet:
            ts = title.getSetter()
            self.ts.append(ts)
        
        # 循环播放原Document
        self.file = Document(self.filename)
        for para in self.file.paragraphs:
            # 如果有图片直接跳过
            for run in para.runs:
                xmlstr = str(run.element.xml)
                if 'pic:pic' in xmlstr:
                    print("adadsda")
                    continue
            if para.style.name.split()[0] == 'Heading':
                level = int(para.style.name.split()[1])
                index = self.title_levels.index(level)
                self.ts[index].run(para)
            elif self.mainBodyFlag == 1:
                ms.run(para)
            if self.picFlag == 1 and para.text[0:len(self.picStart)] == self.picStart:
                ps.run(para)
            elif self.tbFlag == 1 and para.text[0:len(self.tbStart)] == self.tbStart:
                tbs.run(para)
        self.file.save(self.filename.replace('.docx','_new.docx'))
        #状态栏提示保存成功
        self.status.showMessage("File Saves Successfully", 3000)

    def ChooseOver(self):
        titleTab = QWidget()
        secondTab = QScrollArea()
        layout = QVBoxLayout()
        # 检查正文是否需要设置
        if self.mainBodyBt.isChecked():
            self.mainSet = MainBody("正文格式设置")
            layout.addWidget(self.mainSet,1)
        else:
            self.mainBodyFlag = 0
        # 检查图片是否需要设置
        if self.picBt.isChecked():
            self.picFlag = 1
            self.picStart = self.picLine.text()
            self.picSet = MainBody("图名格式设置")
            layout.addWidget(self.picSet,1)
        # 检查表格是否需要设置
        if self.tbBt.isChecked():
            self.tbFlag = 1
            self.tbStart = self.tbLine.text()
            self.tbSet = MainBody("表名格式设置")
            layout.addWidget(self.tbSet,1)

        # 检查标题的设置
        self.title_levels = []
        for bt in self.titleBoxButton:
            if bt.isChecked():
                self.title_levels.append(int(bt.text()[0]))
        if len(self.title_levels) == 0 and self.mainBodyFlag == 0 and self.picFlag == 0 and self.tbFlag == 0:
            QMessageBox.information(self, "Attention", "Nothing Choosed,Nothing Changed")#啥也没选
            return
        self.titleSet = []
        for index in self.title_levels:
            ts = Title(index)
            layout.addWidget(ts,1)
            self.titleSet.append(ts)
        layout.addStretch(10)
        titleTab.setLayout(layout)
        secondTab.setWidget(titleTab)
        while self.tab.count() != 1:
            self.tab.removeTab(1)
        self.tab.addTab(secondTab,"Setting")
        self.tab.setCurrentIndex(1)
        





if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create("Fusion"))
    demo = MainWindow()
    demo.show()
    sys.exit(app.exec_())