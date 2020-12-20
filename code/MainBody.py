from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from Setter import TextSetter
from docx.enum.text import WD_LINE_SPACING

class MainBody(QWidget):
    def __init__(self, title, parent=None):
        super(MainBody, self).__init__(parent)
        #界面设计
        titleGroupBox = QGroupBox(title)
        firstTitleLineLayout = QHBoxLayout()
        secondTitleLineLayout = QHBoxLayout()
        label1 = QLabel("中文字体:")
        label2 = QLabel("英文字体:")
        label3 = QLabel("字号:")
        label4 = QLabel("居中方式:")
        label5 = QLabel("段前:")
        label6 = QLabel("段后:")
        label7 = QLabel("行距:")
        label8 = QLabel("首行缩进:")
        self.is_bold = QCheckBox("加粗")
        self.is_italic = QCheckBox("斜体")

        self.font_name_ch = QLineEdit()
        self.font_name_ch.setText("宋体")
        self.font_name_en = QLineEdit()
        self.font_name_en.setText("Times New Roman")
        self.font_size = QLineEdit()
        self.font_size.setText("小四")

        self.space_before_value = QLineEdit()
        self.space_before_value.setText("0")
        self.space_after_value = QLineEdit()
        self.space_after_value.setText("0")
        self.line_space_value = QComboBox(self)
        self.line_space_value.addItems(['单倍行距','1.5倍行距','2倍行距','最小值','固定值','多倍行距'])
        self.line_space_value_space = QLineEdit()
        self.line_space_value_space.setText("")
        self.alignment = QComboBox(self)
        self.alignment.addItems(['左对齐','居中','右对齐'])
        self.first_line_indent_value = QLineEdit()
        self.first_line_indent_value.setText("无")

        # 设计布局
        firstTitleLineLayout.addWidget(label1, 2)
        firstTitleLineLayout.addWidget(self.font_name_ch, 2)
        firstTitleLineLayout.addStretch(1)
        firstTitleLineLayout.addWidget(label2, 2)
        firstTitleLineLayout.addWidget(self.font_name_en, 2)
        firstTitleLineLayout.addStretch(1)
        firstTitleLineLayout.addWidget(label3, 2)
        firstTitleLineLayout.addWidget(self.font_size, 2)
        firstTitleLineLayout.addWidget(self.is_bold)
        firstTitleLineLayout.addWidget(self.is_italic)
        secondTitleLineLayout.addWidget(label5,2)
        secondTitleLineLayout.addWidget(self.space_before_value,2)
        secondTitleLineLayout.addStretch(1)
        secondTitleLineLayout.addWidget(label6,2)
        secondTitleLineLayout.addWidget(self.space_after_value,2)
        secondTitleLineLayout.addStretch(1)
        secondTitleLineLayout.addWidget(label7,1)
        secondTitleLineLayout.addWidget(self.line_space_value,2)
        secondTitleLineLayout.addWidget(self.line_space_value_space,1)
        secondTitleLineLayout.addStretch(1)
        secondTitleLineLayout.addWidget(label8,2)
        secondTitleLineLayout.addWidget(self.first_line_indent_value,2)
        secondTitleLineLayout.addStretch(1)
        secondTitleLineLayout.addWidget(label4,2)
        secondTitleLineLayout.addWidget(self.alignment,2)

        titleGroupLayout = QVBoxLayout()
        titleGroupLayout.addLayout(firstTitleLineLayout)
        titleGroupLayout.addLayout(secondTitleLineLayout)
        titleGroupBox.setLayout(titleGroupLayout)

        layout = QHBoxLayout()
        layout.addStretch(1)
        layout.addWidget(titleGroupBox,1)
        layout.addStretch(1)
        self.setLayout(layout)


        # 信号槽
        self.line_space_value.currentIndexChanged.connect(self.LineSpace)
        

    def LineSpace(self):
        index = self.line_space_value.currentIndex()
        if index == 0 or index == 1 or index == 2:
            self.line_space_value_space.setText("")
        elif index == 3:
            self.line_space_value_space.setText("12")
        elif index == 4:
            self.line_space_value_space.setText("12")
        elif index == 5:
            self.line_space_value_space.setText("3")

    def getSetter(self):
        font_name_ch = self.font_name_ch.text()
        font_name_en = self.font_name_en.text()
        is_bold = self.is_bold.isChecked()
        is_italic = self.is_italic.isChecked()
        alignment = self.alignment.currentIndex()

        # 字体大小
        font_size = 12.0
        font_size_str = self.font_size.text().replace(' ','')
        size_str = ['初号', '小初', '一号', '小一', '二号', '小二', '三号', '小三', '四号','小四', '五号', '小五', '六号', '小六', '七号', '八号']
        size_int = [42,36,26,24,22,18,16,15,14,12,10.5,9,7.5,6.5,5.5,5]
        index = size_str.index(font_size_str)
        if index != -1:
            font_size = size_int[index]
        else:
            try:
                font_size = float(size_int[index])
            except ValueError as e:
                raise e
        
        # 段前
        space_before_value = 0.0
        space_before_value_str = self.space_before_value.text().replace(' ','')
        if space_before_value_str[-1] == '行':
            space_lines = float(space_before_value_str[0:-1])
            space_before_value = space_lines * 12
        elif space_before_value_str[-1] == '磅':
            space_before_value = float(space_before_value_str[0:-1])
        else:
            try:
                space_before_value = float(space_before_value_str)
            except ValueError as e:
                raise e

        # 段后
        space_after_value = 0.0
        space_after_value_str = self.space_after_value.text().replace(' ','')
        if space_after_value_str[-1] == '行':
            space_lines = float(space_after_value_str[0:-1])
            space_after_value = space_lines * 12
        elif space_after_value_str[-1] == '磅':
            space_after_value = float(space_after_value_str[0:-1])
        else:
            try:
                space_after_value = float(space_after_value_str)
            except ValueError as e:
                raise e
        
        # 行距
        line_space_value = 0.0
        line_space_rule = -1
        index = self.line_space_value.currentIndex()
        if index == 0:
            line_space_value = -1
            line_space_rule = WD_LINE_SPACING.SINGLE
        elif index == 1:
            line_space_value = -1
            line_space_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        elif index == 2:
            line_space_value = -1
            line_space_rule = WD_LINE_SPACING.DOUBLE
        elif index == 3:
            line_space_rule = WD_LINE_SPACING.AT_LEAST
            try:
                line_space_value = float(self.line_space_value_space.text().replace(' ',''))
            except ValueError as e:
                raise e
        elif index == 4:
            line_space_rule = WD_LINE_SPACING.EXACTLY
            try:
                line_space_value = float(self.line_space_value_space.text().replace(' ',''))
            except ValueError as e:
                raise e
        elif index == 5:
            line_space_rule = WD_LINE_SPACING.MULTIPLE
            try:
                line_space_value = float(self.line_space_value_space.text().replace(' ',''))
            except ValueError as e:
                raise e

        # 首行缩进
        first_line_indent_value = 0.0
        first_line_indent_str = self.first_line_indent_value.text().replace(' ','')
        if first_line_indent_str == '无':
            first_line_indent_value = 0.0
        else:
            if first_line_indent_str[-2:] == '字符':
                try:
                    first_line_indent_value = float(first_line_indent_str[:-2]) * font_size
                except ValueError as e:
                    raise e
            elif first_line_indent_str[-2:] == '厘米':
                try:
                    first_line_indent_value = float(first_line_indent_str[:-2]) * 28.346
                except ValueError as e:
                    raise e
            elif first_line_indent_str[-1] == '磅':
                try:
                    first_line_indent_value = float(first_line_indent_str[:-1])
                except ValueError as e:
                    raise e
            else:
                try:
                    first_line_indent_value = float(first_line_indent_str)
                except ValueError as e:
                    raise e

        ts = TextSetter(font_name_ch,font_name_en,font_size,is_bold,is_italic,space_before_value,space_after_value,line_space_value,line_space_rule,alignment,first_line_indent_value)
        return ts
