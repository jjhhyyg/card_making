from docx import Document
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import random

sign = '侯阳洋'


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
        # 文档字体设置
        document.styles['Normal'].font.name = u'宋体'
        document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')

        name = name_list[i]
        random_index = random.randint(0, len(sentence_list)-1)
        sentence = sentence_list[random_index]
        card_name = f"{sign}给{name}的祝福"

        # 名称左对齐
        para = document.add_paragraph(name+'：')
        para.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

        # 祝福语首行缩进2字符
        para = document.add_paragraph(sentence)
        para.paragraph_format.first_line_indent = 406400

        # 署名右对齐
        para = document.add_paragraph('--' + sign)
        para.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

        document.save(card_name+'.docx')
