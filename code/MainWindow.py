from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from Title import Title
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
        title.setText("<p><font size=20 face=arial color=red>标题</font></p>")
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

        #主页面layout
        mainLayout = QVBoxLayout()
        mainLayout.addStretch(2)
        mainLayout.addLayout(mainFirstLayout,1)
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
        ## Title
        self.ts = []
        for title in self.titleSet:
            ts = title.getSetter()
            self.ts.append(ts)
        
        # 循环播放原Document
        self.file = Document(self.filename)
        for para in self.file.paragraphs:
            if para.style.name.split()[0] == 'Heading':
                level = int(para.style.name.split()[1])
                index = self.title_levels.index(level)
                self.ts[index].run(para)
        self.file.save(self.filename.replace('.docx','_new.docx'))
        #状态栏提示保存成功
        self.status.showMessage("File Saves Successfully", 3000)

    def ChooseOver(self):
        self.title_levels = []
        for bt in self.titleBoxButton:
            if bt.isChecked():
                self.title_levels.append(int(bt.text()[0]))
        if len(self.title_levels) == 0:
            QMessageBox.information(self, "Attention", "Nothing Choosed,Nothing Changed")#啥也没选
            return
        titleTab = QWidget()
        titleLayout = QVBoxLayout()
        self.titleSet = []
        for index in self.title_levels:
            ts = Title(index)
            titleLayout.addWidget(ts,1)
            self.titleSet.append(ts)
        titleLayout.addStretch(8-len(self.title_levels))
        titleTab.setLayout(titleLayout)
        self.tab.addTab(titleTab,"Title")
        self.tab.setCurrentIndex(1)





if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create("Fusion"))
    demo = MainWindow()
    demo.show()
    sys.exit(app.exec_())