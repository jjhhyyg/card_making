from docx import Document
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENTATION
from docx.shared import Pt
import random

sign = 'Erikkson'


def read_txt(file_name):
    result = []
    with open(file_name, encoding='utf-8') as file:
        for line in file.readlines():
            line = line.rstrip()
            result.append(line)
    return result


if __name__ == '__main__':
    name_list = read_txt('names.txt')
    sentence_list = read_txt('sentences.txt')

    for i in range(len(name_list)):
        document = Document()
        # 获取本文档中的所有章节
        sections = document.sections
        # 将该章节中的纸张方向设置为横向
        for section in sections:
            # 需要同时设置width,height才能成功
            new_width, new_height = section.page_height, section.page_width
            section.orientation = WD_ORIENTATION.LANDSCAPE
            section.page_width = new_width
            section.page_height = new_height

        # 文档字体设置
        document.styles['Normal'].font.name = u'华文行楷'
        document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'华文行楷')

        name = name_list[i]
        random_index = random.randint(0, len(sentence_list) - 1)
        sentence = sentence_list[random_index]
        card_name = f"{sign}给{name}的祝福"

        # 名称左对齐
        para = document.add_paragraph()
        run = para.add_run(name + '：')
        run.font.size = Pt(16)
        para.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        para.space_after = Pt(18)

        # 祝福语首行缩进2字符
        para = document.add_paragraph()
        run = para.add_run(sentence)
        run.font.size = Pt(20)
        run.font.name = 'Bradley Hand ITC'
        para.paragraph_format.first_line_indent = 406400

        # 署名右对齐
        para = document.add_paragraph()
        run = para.add_run('--' + sign)
        run.font.name = 'Segoe Script'
        run.font.size = Pt(16)
        para.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        para.space_before = Pt(18)

        document.save(card_name + '.docx')
