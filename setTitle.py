from docx import *
from docx.oxml.ns import qn
from docx.shared import Pt,Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT,WD_LINE_SPACING
 
# # 创建一个已存在的 word 文档的对象
# file = Document('test.docx')
# new_file = Document()
# # print(file)
# # print(file.paragraphs)
# print(file.paragraphs[1].text)


# title_level = 1 # 标题等级
# title_text = file.paragraphs[1].text  # 标题内容
# font_name_ch = u"宋体"  # 标题中文字体
# font_name_en = "Times New Roman"  # 标题英文字体
# font_size = 12 # 标题大小，单位磅
# is_bold = True # 加粗？
# is_italic = False # 斜体？

# space_before_value = 12 # 段前行距
# space_after_value = 12 # 段后行距
# line_space_value = 12 # 行距



# p = new_file.add_heading("",level=2)
# print(p)
# text = p.add_run(title_text)
# text.bold = is_bold
# text.italic = is_italic
# font = text.font
# font.name = font_name_en # 英文字体
# font.size = shared.Pt(font_size)
# # text._element.rPr.rFonts.set(qn('w:eastAsia'), font_name_ch)
# text._element.rPr.rFonts.set(qn('w:eastAsia'), font_name_ch) # 中文字体

# # 设置行距 断距
# paragraph_format=p.paragraph_format
# paragraph_format.space_before=Pt(18)    #上行间距
# paragraph_format.space_after=Pt(12)    #下行间距
# paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE#=Pt(18)  #行距s
# paragraph_format.line_spacing=2.5  #行距s
# paragraph_format.alignment = 0#WD_PARAGRAPH_ALIGNMENT.RIGHT # 居中
# paragraph_format.first_line_indent = 406400 # 首行缩进

# # obj = new_file.add_picture('字号磅数对比.png', width=Inches(1.25))
# # print(type(obj))

# new_file.save('res.docx')

file = Document("test.docx")
paras = file.paragraphs
for para in paras:
    print(para.text)
    print( para.paragraph_format.space_before)
    print( para.paragraph_format.space_after)

    para.paragraph_format.space_before = Pt(24)
    para.paragraph_format.space_after = Pt(6)
    print( para.paragraph_format.space_before)
    print( para.paragraph_format.space_after)

file.save("121909.docx")
