from docx.text.paragraph import Paragraph
from docx.shared import Pt,Inches
from docx.oxml.ns import qn
from docx.enum.text import WD_LINE_SPACING

class TextSetter():
    def __init__(self, font_name_ch: str=u"宋体", font_name_en: str="Times New Roman", font_size: float=12, is_bold: bool=False,\
        is_italic: bool=False, space_before_value: float=0.0, space_after_value: float=0.0, line_space_value: float=12.0,\
            line_space_rule: int=WD_LINE_SPACING.SINGLE, alignment: int=0, first_line_indent_value: float=0.0,) -> None:
        self.__font_name_ch = font_name_ch # 中文字体
        self.__font_name_en = font_name_en # 英文字体
        self.__font_size = font_size # 字体大小（磅）
        self.__is_bold = is_bold # 是否加粗
        self.__is_italic = is_italic # 是否斜体
        self.__space_before_value = space_before_value # 段前距 （磅）
        self.__space_after_value = space_after_value # 段后距 （磅）
        self.__line_space_rule = line_space_rule
        self.__line_space_value = line_space_value # 行距 （磅）
        self.__alignment = alignment # 对齐方式
        self.__first_line_indent_value = first_line_indent_value # 首行缩进

    def run(self, para : Paragraph, content : str="") -> None:
        # text = para.add_run(content)
        for text in para.runs:
            text.bold = self.__is_bold
            text.italic = self.__is_italic
            font = text.font
            font.name = self.__font_name_en # 英文字体
            font.size = Pt(self.__font_size)
            text._element.rPr.rFonts.set(qn('w:eastAsia'), self.__font_name_ch) # 中文字体

        # 设置行距 断距
        paragraph_format = para.paragraph_format
        paragraph_format.space_before = Pt(self.__space_before_value)    #上行间距
        paragraph_format.space_after = Pt(self.__space_after_value)    #下行间距
        paragraph_format.line_spacing_rule = self.__line_space_rule
        if self.__line_space_value != -1:
            paragraph_format.line_spacing = Pt(self.__line_space_value) #行距s
        paragraph_format.alignment = self.__alignment # 居中？左右？
        paragraph_format.first_line_indent = Pt(self.__first_line_indent_value) # 首行缩进

