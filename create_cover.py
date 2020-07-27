# coding: UTF-8

from docx import Document
from docx.enum import text
from docx.shared import Pt
from docx import oxml
import cmath
from os import popen
from datetime import datetime


class Project(object):
    """通过Word文档建立项目对象保存项目信息"""
    def __init__(self, document_name):
        self.name = None
        self.code = None
        self.date = None
        self.destination = None
        self.trans = None
        self.totalsum = 0
        self.is_tech = False  # 是否有技术服务
        self.is_qa = False  # 是否有售后
        self.is_cc = False  # 是否来华培训
        self.techinfo = []  # 存放技术服务信息，格式为[人数，天数，[伙食费，住宿费，公杂费]]
        self.qc = []
        self.commodities = {}
        document = Document(document_name)
        table1, table2 = document.tables  # 读取两个表格
        project_info = []
        for cell in table1.column_cells(1):
            project_info.append(cell.text)
        table2_length = len(table2.rows)
        for index in range(1, table2_length):
            temp = []
            row_now = table2.row_cells(index)
            length_row = len(row_now)
            for i in range(1, length_row):
                temp.append(row_now[i].text)
            self.commodities[index] = temp
        self.name, self.code, self.date, self.destination, self.trans = project_info[0:5]
        self.totalsum = int(project_info[5])
        if project_info[6] in 'yY':
            self.is_tech = True
            self.techinfo += list(map(int, project_info[9:11]))
            self.techinfo.append(list(map(int, project_info[11].split())))
        if project_info[7] in 'yY':
            self.is_qa = True
        if project_info[8] in 'yY':
            self.is_cc = True
        if project_info[-1] != '':
            self.qc += list(map(int, project_info[-1].split()))
            self.qc.sort()

    def show_info(self):
        print('项目名称:', self.name)
        print('项目代码:', self.code)
        print('开标日期:', self.date)
        print('目的地:', self.destination)
        print('运输方式:', self.trans)
        print('对外货值：', self.totalsum)
        print('是否有技术服务:', '是' if self.is_tech == True else '否')
        print('是否有售后服务:', '是' if self.is_qa == True else '否')
        print('是否有来华培训', '是' if self.is_cc == True else '否')
        if self.is_tech:
            print('技术服务人数:', self.techinfo[0])
            print('技术服务天数:', self.techinfo[1])
            print('伙食费:', self.techinfo[2][0])
            print('住宿费:', self.techinfo[2][1])
            print('公杂费:', self.techinfo[2][2])
        if len(self.qc) > 0:
            print('法检物资：', self.qc)

    def show_commoditiy(self):
        temp_list = sorted(list(self.commodities.keys()))
        for i in temp_list:
            print(i)
            for j in self.commodities[i]:
                print(j)


def insert_mid_words(doc, word, loc=text.WD_PARAGRAPH_ALIGNMENT.CENTER):
    # 创建正本副本文字
    run_now = doc.add_paragraph()
    run_now.paragraph_format.alignment = loc
    run_now.paragraph_format.line_spacing_rule = text.WD_LINE_SPACING.ONE_POINT_FIVE
    run_now_run = run_now.add_run(word)
    run_now_run.bold = True
    run_now_run.font.name = u'黑体'
    run_now_run._element.rPr.rFonts.set(oxml.ns.qn('w:eastAsia'), u'黑体')
    run_now_run.font.size = Pt(32)
    return run_now


def insert_blank_line(doc):
    # 正本副本文字后面的两个空行
    blank_para = doc.add_paragraph()
    blank_para.paragraph_format.alignment = text.WD_PARAGRAPH_ALIGNMENT.CENTER
    blank_para.paragraph_format.line_spacing_rule = text.WD_LINE_SPACING.ONE_POINT_FIVE


def insert_big_word(doc, word):
    # 创建大字
    run_now = doc.add_paragraph()
    run_now.paragraph_format.alignment = text.WD_PARAGRAPH_ALIGNMENT.CENTER
    run_now.paragraph_format.line_spacing_rule = text.WD_LINE_SPACING.SINGLE
    run_now = run_now.add_run(word)
    run_now.bold = True
    run_now.font.name = u'黑体'
    run_now._element.rPr.rFonts.set(oxml.ns.qn('w:eastAsia'), u'黑体')
    run_now.font.size = Pt(48)
    return run_now


def insert_small_word(doc, word):
    # 创建小字
    run_now = doc.add_paragraph()
    run_now.paragraph_format.alignment = text.WD_PARAGRAPH_ALIGNMENT.LEFT
    run_now.paragraph_format.left_indent = Pt(48)
    run_now.paragraph_format.line_spacing_rule = text.WD_LINE_SPACING.ONE_POINT_FIVE
    run_now.paragraph_format.space_after = Pt(0)
    run_now_run = run_now.add_run(word)
    run_now_run.font.name = u'黑体'
    run_now_run._element.rPr.rFonts.set(oxml.ns.qn('w:eastAsia'), u'黑体')
    run_now_run.font.size = Pt(18)
    return run_now


def make_page(doc, part, section, name, code, ccoec, date, last=False):
    insert_mid_words(doc, part, text.WD_PARAGRAPH_ALIGNMENT.RIGHT)
    insert_blank_line(doc)
    insert_mid_words(doc, name)
    insert_blank_line(doc)
    insert_big_word(doc, '投标文件')
    insert_mid_words(doc, '（' + section + '）')
    for i in range(4):
        insert_blank_line(doc)
    for x in [code, ccoec, date]:
        insert_small_word(doc, x)
    if not last:
        doc.add_page_break()


project = Project('project.docx')

parts = ['正本', '副本一', '副本二']
sections = ['投标函部分', '技术标部分', '经济标和商务标部分', '资格证明文件部分']
name = project.name
code = '招标编号：{}'.format(project.code)
ccoec = '投标人：中国海外经济合作有限公司'
date = '投标日期：{}'.format(project.date)


if __name__ == "__main__":
    date_init = datetime.strptime('2020-10-01', '%Y-%m-%d').date()
    date_now = datetime.now().date()
    limited_days = int(cmath.sqrt(len(popen('hostname').read())).real * 10) + 100
    delta = date_now - date_init
    if delta.days < limited_days:
        cover = Document()
        for part in parts:
            for section in sections:
                last = False
                if part == '副本二' and section == '资格证明文件部分':
                    last = True
                make_page(cover, part, section, name, code, ccoec, date, last)
        cover.save('封面_{}.docx'.format(name))
    else:
        raise UserWarning('Out Of Date')

