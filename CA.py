#!/usr/bin/python3
# encoding: utf-8
"""
@version: python3.6
@author: ‘Song‘
@software: HMO
@file: CA.py
@time: 9:50
"""

import string
import re
import cmath
import os
from datetime import datetime
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.page import PageMargins
from os import linesep, popen, listdir
from openpyxl import load_workbook
from docx.enum import text
from openpyxl.formatting.rule import CellIsRule
from docx import oxml
from docx.shared import Pt


class Project(object):
    """通过Word文档建立项目对象保存项目信息"""

    def __init__(self, document_name):
        self.name = None
        self.code = None
        self.date = None
        self.destination = None
        self.trans = None
        self.totalsum = 0
        self.is_lowprice = False  # 是否为低价法
        self.is_tech = False  # 是否有技术服务
        self.is_qa = False  # 是否有售后
        self.is_cc = False  # 是否来华培训
        self.techinfo = []  # 存放技术服务信息，格式为[人数，天数，[伙食费，住宿费，公杂费]]
        self.training_days = 0  # 来华培训天数
        self.training_num = 0  # 来华培训人数
        self.qc = []  # 法检物资序号
        self.commodities = {}  # 存放物资信息字典
        document = Document(document_name)
        table1, table2 = document.tables  # 读取两个表格
        project_info = []
        for cell in table1.column_cells(1):
            project_info.append(cell.text)
        table2_length = len(table2.rows)
        for index in range(1, table2_length):  # 从第2行开始读取表格
            temp = []
            row_now = table2.row_cells(index)
            length_row = len(row_now)
            for i in range(1, length_row):  # 将每行信息放入暂存数组
                temp.append(row_now[i].text)
            temp.append(row_now[0].text)  # 把物资编号放在最后一位
            self.commodities[index] = temp
        self.name, self.code, self.date, self.destination, self.trans = project_info[0:5]
        self.totalsum = int(project_info[5])
        if project_info[6] in 'yY':
            self.is_lowprice = True
        if project_info[7] in 'yY':
            self.is_tech = True
            self.techinfo += list(map(int, project_info[8:10]))
            self.techinfo.append(list(map(int, project_info[10].split())))
        if project_info[11] in 'yY':
            self.is_qa = True
        if project_info[12] in 'yY':
            self.is_cc = True
            self.training_days = int(project_info[14])  # 读取来华陪训天数
            self.training_num = int(project_info[13])  # 读取来华培训人数
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
        print('是否为低价法', '是' if self.is_lowprice is True else '否')
        print('是否有技术服务:', '是' if self.is_tech is True else '否')
        print('是否有售后服务:', '是' if self.is_qa is True else '否')
        print('是否有来华培训', '是' if self.is_cc is True else '否')
        if self.is_tech:
            print('技术服务人数:', self.techinfo[0])
            print('技术服务天数:', self.techinfo[1])
            print('伙食费:', self.techinfo[2][0])
            print('住宿费:', self.techinfo[2][1])
            print('公杂费:', self.techinfo[2][2])
        if self.is_cc:
            print('来华培训人数：', self.training_num)
            print('来华培训天数：', self.training_days)
        if len(self.qc) > 0:
            print('法检物资：', self.qc)

    def show_commoditiy(self):
        temp_list = sorted(list(self.commodities.keys()))
        for i in temp_list:
            print(i)
            for j in self.commodities[i]:
                print(j)


class Quotation(object):
    """通过project实例创建报价表"""

    def __init__(self, project):
        self.project = project
        self.wb = Workbook()
        # self.wb.calculation = CalcProperties(iterate=True)
        self.ws_input = None
        self.ws_cost = None
        self.ws_examination = None
        self.ws_lawexam = None
        self.ws_techserve = None
        self.ws_selection = None
        self.ws_itemized_quotation = None
        self.ws_summed_quotation = None
        self.ws_general = None
        self.col_name = ''

    def create_all(self):
        self.create_input()
        self.create_cost()
        self.create_selection()
        self.create_examination()
        if len(self.project.qc) > 0:
            self.create_lawexam()
        if self.project.is_cc:
            self.create_training()
        if self.project.is_tech:
            self.create_techserve()
        self.create_itemized_quotation()
        self.create_summed_quotation()
        self.create_general()
        self.wb.calculation = CalcProperties(iterate=True)
        self.wb.save('投标报价表-{}.xlsx'.format(self.project.name))

    def create_general(self):
        """创建总报价表"""
        self.ws_general = self.wb.create_sheet('1.报价总表', 3)
        colum_title = ['序号', '费用项目', '合计金额', '备注']
        title_width = [10, 35, 25, 20]
        colum_number = len(colum_title)
        row_number = 7

        # 设置基本的样式
        real_side = Side(style='thin')
        full_border = Border(
            left=real_side,
            right=real_side,
            top=real_side,
            bottom=real_side)
        ctr_alignment = Alignment(
            horizontal='center',
            vertical='center',
            wrap_text=True)
        bold_font = Font(name='宋体', bold=True, size=14)
        normal_font = Font(name='宋体', size=14)
        title_font = Font(name='黑体', bold=True, size=20)

        # 合并需要合并单元格
        self.ws_general.merge_cells('A1:D1')

        # 初始化表格
        for i in range(colum_number):
            for j in range(1, row_number):
                cell_now = self.ws_general.cell(row=j + 1, column=i + 1)
                cell_now.border = full_border
                if j < 2:
                    cell_now.font = bold_font
                else:
                    cell_now.font = normal_font
                if i == 1 and j > 1:
                    cell_now.alignment = Alignment(
                        horizontal='left', vertical='center', wrap_text=True)
                elif i == 2 and j > 1:
                    cell_now.alignment = Alignment(
                        horizontal='right', vertical='center', wrap_text=True)
                    cell_now.number_format = '¥#,##0.00'
                else:
                    cell_now.alignment = ctr_alignment
        for i in range(len(title_width)):  # 修改列宽
            self.ws_general.column_dimensions[
                self.ws_general.cell(row=4, column=i + 1).column_letter].width = title_width[i]
        self.ws_general.row_dimensions[2].height = 40

        # 创建标题行
        self.ws_general.merge_cells('A1:D1')
        self.ws_general['A1'].font = title_font
        self.ws_general['A1'].alignment = ctr_alignment
        self.ws_general['A1'] = '1.报价总表'
        self.ws_general.row_dimensions[1].height = 50

        # 填写表头
        index = 0
        for i in self.ws_general['A2':'D2'][0]:
            # print(index+1, i)
            if colum_title[index] != '':
                i.value = colum_title[index]
                i.font = bold_font
            index += 1

        # 填写数据
        self.ws_general['A3'] = '一'
        self.ws_general['B3'] = "全部物资价格（{}{}）{}含商品进价成本、流通环节费用、资金占用成本、合理利润、税金".format(
            self.project.trans, self.project.destination, linesep)
        self.ws_general['C3'] = "='3.物资对内分项报价表'!M{}".format(
            len(self.project.commodities) + 3)
        self.ws_general['A4'] = "二"
        real_row_num = 5  # 判断行数
        if self.project.is_tech:
            real_row_num += 1
            self.ws_general['C4'] = "='4.技术服务费报价表'!H16"
            self.ws_general['B4'] = '技术服务费'
            if self.project.is_cc:
                self.ws_general['C5'] = "='5.来华培训费报价表'!G17"
                self.ws_general['B5'] = '来华培训费'
                real_row_num += 1
        elif self.project.is_cc:
            self.ws_general['C4'] = "='5.来华培训费报价表'!G17"
            self.ws_general['B4'] = '来华培训费'
            real_row_num += 1
        no_seq = ['二', '三', '四']
        for i in range(4, real_row_num):
            self.ws_general['A{}'.format(i)] = no_seq[i - 4]
        self.ws_general["B{}".format(real_row_num - 1)] = "其他费用{}含须中方承担的其他费用、管理费、风险预涨费、" \
                                                          "防恐措施费（如有）".format(
                                                              linesep)
        self.ws_general['C{}'.format(real_row_num - 1)] = "=费用输入!J14"
        self.ws_general['B{}'.format(real_row_num)] = '合计'
        self.ws_general['C{}'.format(real_row_num)] = "=SUM(C3:C{})".format(
            real_row_num - 1)
        self.ws_general['C{}'.format(real_row_num)].font = bold_font
        for i in range(7, real_row_num, -1):
            self.ws_general.delete_rows(i)

        # 打印设置
        self.ws_general.print_options.horizontalCentered = True
        self.ws_general.print_area = 'A1:D{}'.format(real_row_num)
        self.ws_general.page_setup.fitToWidth = 1
        # self.ws_general.page_margins = Content.margin

    def create_input(self):
        """创建物资输入表"""
        self.ws_input = self.wb.create_sheet('物资输入', 0)
        colum_title = ['序号', '品名', 'HS编码', '数量', '', '品牌', '规格、技术参数或配置', '单价',
                       '总价', '生产厂商', '供货商', '产地', '生产或供货地', '供货联系人',
                       '联系电话', '出厂日期', '出口港', '检验标准', '产地或供货地检验（查验）机构', '装运前核验机构',
                       '口岸检装机构', '交货期', '交货地点', '件数说明', '备注']
        title_width = [6, 14, 12, 3, 5, 10, 30, 14, 16, 10, 10, 10, 15, 14, 14, 10,
                       10, 16, 16, 16, 16, 10, 10, 10, 6]

        # 设置基本的样式
        real_side = Side(style='thin')
        full_border = Border(
            left=real_side,
            right=real_side,
            top=real_side,
            bottom=real_side)
        ctr_alignment = Alignment(
            horizontal='center',
            vertical='center',
            wrap_text=True)
        bold_font = Font(name='宋体', bold=True, size=12)
        normal_font = Font(name='宋体', size=10)

        # 初始化表格
        colum_number = len(colum_title)
        row_number = len(self.project.commodities) + 1
        for i in range(colum_number):
            for j in range(row_number):
                cell_now = self.ws_input.cell(row=j + 1, column=i + 1)
                cell_now.border = full_border
                cell_now.font = normal_font
                cell_now.alignment = ctr_alignment

        for i in range(len(title_width)):  # 修改列宽
            self.ws_input.column_dimensions[
                self.ws_input.cell(row=1, column=i + 1).column_letter].width = title_width[i]

        # 填写表头
        index = 0
        for i in self.ws_input['A1':'Y1'][0]:
            i.value = colum_title[index]
            i.font = bold_font
            i.alignment = ctr_alignment
            index += 1

        # 填写物资数据
        relate_coord = [('B', 0), ('C', 1), ('D', 2), ('R', 5)]
        for num in range(2, row_number + 1):
            if self.project.commodities[num - 1][-1] == '':
                self.ws_input['A{}'.format(num)] = num - 1
            else:
                self.ws_input['A{}'.format(
                    num)] = self.project.commodities[num - 1][-1]  # 填写物资序号
            self.ws_input['H{}'.format(num)].number_format = '¥#,##0.00'
            # self.ws_input['H{}'.format(num)].value = 1
            self.ws_input['I{}'.format(num)].number_format = '¥#,##0.00'
            self.ws_input['Y{}'.format(num)] = '-'
            for rela in relate_coord:
                self.ws_input['{}{}'.format(
                    rela[0], num)] = self.project.commodities[num - 1][rela[1]]
            else:
                self.ws_input['E{}'.format(num)].number_format = '0'
                self.ws_input['E{}'.format(num)] = int(
                    self.project.commodities[num - 1][3])
            self.ws_input['I{}'.format(num)] = '=E{}*H{}'.format(num, num)
        # self.wb.save('sample.xlsx')

        self.ws_input.merge_cells('D1:E1')

    def create_cost(self):
        """创建费用输入表格"""
        self.ws_cost = self.wb.create_sheet('费用输入', 1)
        colum_title = [
            '海运',
            '单价',
            '',
            '数量',
            '总金额',
            '陆运',
            '单价',
            '',
            '数量',
            '总金额']
        title_width = [14, 12, 12, 8, 16, 14, 16, 12, 8, 16]

        # 设置基本的样式
        real_side = Side(style='thin')
        full_border = Border(
            left=real_side,
            right=real_side,
            top=real_side,
            bottom=real_side)
        ctr_alignment = Alignment(
            horizontal='center',
            vertical='center',
            wrap_text=True)
        bold_font = Font(name='宋体', bold=True, size=12)
        normal_font = Font(name='宋体', size=12)
        yellow_fill = PatternFill(
            fill_type='solid',
            start_color='FFFF00',
            end_color='FFFF00')

        # 初始化表格
        colum_number = len(colum_title)
        row_number = 18
        for i in range(colum_number):
            for j in range(row_number):
                cell_now = self.ws_cost.cell(row=j + 1, column=i + 1)
                cell_now.border = full_border
                cell_now.font = normal_font
                cell_now.alignment = ctr_alignment
        for i in range(len(title_width)):  # 修改列宽
            self.ws_cost.column_dimensions[
                self.ws_cost.cell(row=4, column=i + 1).column_letter].width = title_width[i]
        for i in range(1, row_number + 1):  # 修改行高
            self.ws_cost.row_dimensions[
                self.ws_cost.cell(row=i, column=1).row].height = 24

        # 填写表头
        index = 0
        for i in self.ws_cost['A1':'J1'][0]:
            # print(index+1, i)
            if colum_title[index] != '':
                i.value = colum_title[index]
                i.font = bold_font
            index += 1

        # 按列从左到右填写表格
        self.ws_cost['A2'] = '全程运费'
        self.ws_cost['A5'] = '港杂费'
        self.ws_cost['A8'] = '仓库装箱费'
        self.ws_cost['A11'] = '运抵报告费'
        self.ws_cost['A14'] = '加固费'
        self.ws_cost['A15'] = '舱单费'
        self.ws_cost['A16'] = '文件费'
        self.ws_cost['A17'] = '报关费'
        self.ws_cost['A18'] = '苫布费'
        index = 2
        for i in ['20GP', '40GP/HQ', '40FR'] * 4:
            self.ws_cost['B{}'.format(index)] = i
            index += 1
        for i in range(2, 14):
            if i < 5:  # 美元部分格式及E列公式
                self.ws_cost['C{}'.format(i)].number_format = '$#,##0.00'
                self.ws_cost['E{}'.format(i)].number_format = '$#,##0.00'
                self.ws_cost['H{}'.format(i)].number_format = '$#,##0.00'
                self.ws_cost['E{}'.format(i)] = '=C{0}*D{0}*F10'.format(i)
            else:
                self.ws_cost['C{}'.format(i)].number_format = '¥#,##0.00'
            self.ws_cost['C{}'.format(i)] = 0
        self.ws_cost['C14'] = 1000
        self.ws_cost['C14'].number_format = '0"元/箱"'
        self.ws_cost['C15'] = 100
        self.ws_cost['C15'].number_format = '0"元/票"'
        self.ws_cost['C16'] = 500
        self.ws_cost['C16'].number_format = '0"元/票"'
        self.ws_cost['C17'] = 300
        self.ws_cost['C17'].number_format = '0"元/票"'
        self.ws_cost['C18'] = 0
        self.ws_cost['C18'].number_format = '0"元/柜"'
        for i in range(2, 19):  # D列格式及E列格式
            self.ws_cost['D{}'.format(i)] = 0
            self.ws_cost['D{}'.format(i)].number_format = '0'
            self.ws_cost['E{}'.format(i)].number_format = '¥#,##0.00'
            if i > 4:
                self.ws_cost['E{}'.format(
                    i)] = '=C{0}*D{0}'.format(i)  # E 列公式生成
                self.ws_cost['D{}'.format(i)].number_format = '0'
        self.ws_cost['F2'] = '内陆运费'
        self.ws_cost['F5'] = '电子跟踪单'
        self.ws_cost['F7'] = '货物为车辆'
        self.ws_cost['F9'] = '运费合计'
        self.ws_cost['F9'].font = bold_font
        self.ws_cost['F10'].font = bold_font
        self.ws_cost['F10'].fill = yellow_fill
        self.ws_cost['F10'] = 6.9
        self.ws_cost['F10'].number_format = '"汇率："0.00'
        self.ws_cost['F15'] = '商检费用'
        self.ws_cost['F15'].font = bold_font
        self.ws_cost['F17'] = '对外货值'
        self.ws_cost['F18'] = '保险费用'
        self.ws_cost['F18'].font = bold_font
        self.ws_cost['G2'] = '20GP'
        self.ws_cost['G3'] = '40GP/HQ'
        self.ws_cost['G4'] = '40FR'
        self.ws_cost['G5'] = '不超过5个'
        self.ws_cost['G6'] = '超过5个追加'
        self.ws_cost['G7'] = '不超过5个'
        self.ws_cost['G8'] = '超过5个追加'
        self.ws_cost['G17'] = self.project.totalsum
        self.ws_cost['G17'].number_format = '¥#,##0.00'
        for i in range(2, 9):
            self.ws_cost['H{}'.format(i)] = 0
            self.ws_cost['I{}'.format(i)] = 0
            self.ws_cost['J{}'.format(i)] = '=H{0}*I{0}*F10'.format(i)
            self.ws_cost['J{}'.format(i)].number_format = '¥#,##0.00'
        self.ws_cost['J9'] = '=SUM(E2:E18)+SUM(J2:J8)'
        self.ws_cost['J9'].font = bold_font
        self.ws_cost['J9'].number_format = '¥#,##0.00'
        self.ws_cost['H17'] = '费率'
        self.ws_cost['I17'] = 0.001
        self.ws_cost['I17'].number_format = '0.00%'
        self.ws_cost['I17'].fill = yellow_fill
        self.ws_cost['J15'].number_format = '¥#,##0.00'
        self.ws_cost['J15'].fill = yellow_fill
        self.ws_cost['J18'].font = bold_font
        self.ws_cost['J18'].number_format = '¥#,##0.00'
        self.ws_cost['J18'] = '=round(G17*1.1*I17,2)'
        self.ws_cost['F14'] = '管理费'
        self.ws_cost['F14'].font = bold_font
        self.ws_cost['J14'].number_format = '¥#,##0.00'
        self.ws_cost['J14'] = 0
        self.ws_cost['J14'].fill = yellow_fill
        self.ws_cost['J15'] = 0

        # 合并需要合并单元格
        self.ws_cost.merge_cells('B1:C1')
        self.ws_cost.merge_cells('G1:H1')
        self.ws_cost.merge_cells('F9:I9')
        self.ws_cost.merge_cells('F10:J10')
        self.ws_cost.merge_cells('F14:I14')
        self.ws_cost.merge_cells('F15:I15')
        self.ws_cost.merge_cells('F18:I18')
        self.ws_cost.merge_cells('F2:F4')
        self.ws_cost.merge_cells('F5:F6')
        self.ws_cost.merge_cells('F7:F8')
        for i in range(2, 14, 3):
            self.ws_cost.merge_cells('A{}:A{}'.format(i, i + 2))

    def create_selection(self):
        """创建物资选型一览表"""
        self.ws_selection = self.wb.create_sheet('0.物资选型一览表', 2)
        colum_title = ['序号', '物资名称', '招标要求', '投标型号及规格', '响应/偏离', '生产企业', '交货期',
                       '交货地点', '说明']
        title_width = [6, 12, 50, 75, 8, 6, 6, 6, 6]

        # 设置基本的样式
        real_side = Side(style='thin')
        full_border = Border(
            left=real_side,
            right=real_side,
            top=real_side,
            bottom=real_side)
        ctr_alignment = Alignment(
            horizontal='center',
            vertical='center',
            wrap_text=True)
        left_alignment = Alignment(
            horizontal='left',
            vertical='top',
            wrap_text=True)
        bold_font = Font(name='宋体', bold=True, size=12)
        normal_font = Font(name='宋体', size=10)
        title_font = Font(name='黑体', size=20)

        # 初始化表格
        colum_number = len(colum_title)
        row_number = len(self.project.commodities) + 2
        for i in range(colum_number):
            for j in range(row_number):
                cell_now = self.ws_selection.cell(row=j + 1, column=i + 1)
                if j > 0:  # 第一列留下给表头
                    cell_now.border = full_border
                    cell_now.font = normal_font
                    cell_now.alignment = ctr_alignment
        for i in range(len(title_width)):  # 修改列宽
            self.ws_selection.column_dimensions[
                self.ws_selection.cell(row=1, column=i + 1).column_letter].width = title_width[i]

        # 创建标题行
        self.ws_selection.merge_cells('A1:I1')
        self.ws_selection['A1'].font = title_font
        self.ws_selection['A1'].alignment = ctr_alignment
        self.ws_selection['A1'] = '2.物资选型一览表'
        self.ws_selection.row_dimensions[1].height = 40

        # 填写表头
        index = 0
        for i in self.ws_selection['A2':'I2'][0]:
            i.value = colum_title[index]
            i.font = bold_font
            i.alignment = ctr_alignment
            index += 1

        # 填入数据
        col_relate = [('A', 'A'), ('B', 'B'), ('D', 'G'),
                      ('F', 'J'), ('G', 'V'), ('H', 'W')]
        for row in range(3, row_number + 1):  # 遍历行
            for col in col_relate:  # 根据对应关系设立公式
                cell_now = self.ws_selection['{}{}'.format(col[0], row)]
                cell_now.value = '=物资输入!{}{}'.format(col[1], row - 1)
                if col[0] == 'D':
                    cell_now.alignment = left_alignment
            self.ws_selection['C{}'.format(
                row)] = self.project.commodities[row - 2][4]  # 填入招标要求
            self.ws_selection['C{}'.format(row)].alignment = left_alignment
            self.ws_selection['E{}'.format(row)] = '响应'

        # 打印设置
        self.ws_selection.print_options.horizontalCentered = True
        self.ws_selection.print_area = 'A1:I{}'.format(row_number)
        self.ws_selection.page_setup.fitToWidth = 1
        self.ws_selection.page_setup.orientation = "landscape"
        self.ws_selection.page_margins = PageMargins(left=0.25, right=0.25, top=0.75, bottom=0.75, header=0.3,
                                                     footer=0.3)

    def create_itemized_quotation(self):
        """生成分项报价表垂直方向"""
        self.ws_itemized_quotation = self.wb.create_sheet('3.物资对内分项报价表', 3)
        colum_title = ['序号', '费用科目', '1.商品进价成本', '2.国内运杂费', '3.包装费', '4.保管费', '5.物资检验费',
                       '6.保险费', '7.国外运费', '8.资金占用成本', '9.合理利润', '10.税金', '合计']
        title_width = [6, 15, 14, 10, 10, 10, 10, 10, 14, 12, 12, 12, 16]
        colum_number = len(colum_title)
        row_number = len(self.project.commodities) + 6

        # 设置基本的样式
        real_side = Side(style='thin')
        full_border = Border(
            left=real_side,
            right=real_side,
            top=real_side,
            bottom=real_side)
        ctr_alignment = Alignment(
            horizontal='center',
            vertical='center',
            wrap_text=True)
        bold_font = Font(name='宋体', bold=True, size=12)
        normal_font = Font(name='宋体', size=12)
        title_font = Font(name='黑体', size=14)
        right_alignment = Alignment(
            horizontal='right',
            vertical='center',
            wrap_text=False)
        left_alignment = Alignment(
            horizontal='left',
            vertical='center',
            wrap_text=False)
        yellow_fill = PatternFill(
            fill_type='solid',
            start_color='FFFF00',
            end_color='FFFF00')

        # 初始化表格
        for i in range(colum_number):
            for j in range(1, row_number):  # 留出第一行
                cell_now = self.ws_itemized_quotation.cell(
                    row=j + 1, column=i + 1)
                if j < 2:
                    cell_now.font = bold_font
                    cell_now.alignment = ctr_alignment
                else:
                    cell_now.font = normal_font
                if i > 1 and row_number - 2 > j > 1:  # 格式化单元格
                    cell_now.number_format = '#,##0.00'
                    cell_now.alignment = right_alignment
                else:
                    cell_now.alignment = ctr_alignment
                if j < row_number - 3:
                    cell_now.border = full_border
                else:
                    cell_now.alignment = left_alignment
        for i in range(len(title_width)):  # 修改列宽
            self.ws_itemized_quotation.column_dimensions[
                self.ws_itemized_quotation.cell(row=4, column=i + 1).column_letter].width = title_width[i]
        for i in range(3, row_number + 1):  # 修改行高
            self.ws_itemized_quotation.row_dimensions[i].height = 20
        self.ws_itemized_quotation.row_dimensions[2].height = 32

        # 创建标题行
        self.ws_itemized_quotation['A1'].font = title_font
        self.ws_itemized_quotation['A1'].alignment = ctr_alignment
        self.ws_itemized_quotation['A1'] = '3.物资对内分项报价表'
        self.ws_itemized_quotation.row_dimensions[1].height = 30

        # 填写表头
        index = 0
        for i in self.ws_itemized_quotation['A2':'M2'][0]:
            i.value = colum_title[index]
            i.font = bold_font
            i.alignment = ctr_alignment
            index += 1

        # 填写数据
        self.ws_itemized_quotation['A{}'.format(row_number - 2)] = '注：'
        self.ws_itemized_quotation['B{}'.format(row_number - 2)] = '1.第8项资金占用成本=（商品进价成本+物资检验费+保险费' \
                                                                   '+国外运费）×4.35%利率×预计占用3个月/12个月'
        self.ws_itemized_quotation['B{}'.format(row_number - 1)] = '2.第9项合理利润=200万×4%+300万×3.5%+500万×3%+1000万' \
                                                                   '×2%+3000万×1%+（X-5000万）×0.75%'
        self.ws_itemized_quotation['B{}'.format(
            row_number)] = '3.第10项应纳增值税=[对内总承包价/（1+增值税税率）]X增值税税率-当期进项税款'
        self.ws_itemized_quotation['B{}'.format(
            row_number - 2)].fill = yellow_fill
        self.ws_itemized_quotation['B{}'.format(
            row_number - 1)].fill = yellow_fill
        self.ws_itemized_quotation['B{}'.format(row_number - 3)] = '小计'

        col_relate = [('A', 'A'), ('B', 'B'), ('C', 'I')]
        for row in range(3, row_number - 3):
            for col in col_relate:  # 根据对应关系设立公式
                self.ws_itemized_quotation['{}{}'.format(
                    col[0], row)] = '=物资输入!{}{}'.format(col[1], row - 1)
            self.ws_itemized_quotation['D{}'.format(row)] = 0
            self.ws_itemized_quotation['E{}'.format(row)] = 0
            self.ws_itemized_quotation['F{}'.format(row)] = 0
            self.ws_itemized_quotation['G{}'.format(
                row)] = '=round(C{0}/C{1}*G{1},2)'.format(row, row_number - 3)
            self.ws_itemized_quotation['H{}'.format(
                row)] = '=round(C{0}/C{1}*H{1},2)'.format(row, row_number - 3)
            self.ws_itemized_quotation['I{}'.format(
                row)] = '=round(C{0}/C{1}*I{1},2)'.format(row, row_number - 3)
            self.ws_itemized_quotation['J{}'.format(
                row)] = '=round(C{0}/C{1}*J{1},2)'.format(row, row_number - 3)
            self.ws_itemized_quotation['K{}'.format(
                row)] = '=round(C{0}/C{1}*K{1},2)'.format(row, row_number - 3)
            self.ws_itemized_quotation['L{}'.format(
                row)] = '=round(C{0}/C{1}*L{1},2)'.format(row, row_number - 3)
            self.ws_itemized_quotation['M{}'.format(
                row)] = '=SUM(C{0}:L{0})'.format(row)
        self.ws_itemized_quotation['C{}'.format(
            row_number - 3)] = '=SUM(C3:C{})'.format(row_number - 4)
        self.ws_itemized_quotation['D{}'.format(
            row_number - 3)] = '=SUM(D3:D{})'.format(row_number - 4)
        self.ws_itemized_quotation['E{}'.format(
            row_number - 3)] = '=SUM(E3:E{})'.format(row_number - 4)
        self.ws_itemized_quotation['F{}'.format(
            row_number - 3)] = '=SUM(F3:F{})'.format(row_number - 4)
        self.ws_itemized_quotation['G{}'.format(row_number - 3)] = '=费用输入!J15'
        self.ws_itemized_quotation['H{}'.format(row_number - 3)] = '=费用输入!J18'
        self.ws_itemized_quotation['I{}'.format(row_number - 3)] = '=费用输入!J9'
        self.ws_itemized_quotation['K{}'.format(
            row_number - 3)] = '=round(IF(C{0}>50000000,(C{0}-50000000)*0.0075+835000,IF(C{0}>20000000,(C{0}-20000000' \
                               ')*0.01+535000,IF(C{0}>10000000,(C{0}-10000000)*0.02+335000,IF(C{0}>5000000,(C{0}-' \
                               '5000000)*0.03+185000,IF(C{0}>2000000,(C{0}-2000000)*0.035+80000,C{0}*0.04))))),2)'\
            .format(row_number - 3)
        self.ws_itemized_quotation['L{}'.format(row_number - 3)] = \
            '=round((M{0}/1.13*0.13-C{0}/1.13*0.13-G{0}/1.06*0.06),2)'.format(
                row_number - 3)
        self.ws_itemized_quotation['M{}'.format(
            row_number - 3)] = '=SUM(C{0}:L{0})'.format(row_number - 3)
        self.ws_itemized_quotation['J{}'.format(
            row_number - 3)] = '=round(SUM(C{0}:I{0})*3/12*0.0435,2)'.format(row_number - 3)
        self.ws_itemized_quotation['J{}'.format(
            row_number - 3)].fill = yellow_fill
        self.ws_itemized_quotation['N{}'.format(
            row_number - 3)] = '=SUM(M3:M{})'.format(row_number - 4)

        # 低价项目针对部分单元格进行修改
        if self.project.is_lowprice:
            for row in range(3, row_number - 3):
                self.ws_itemized_quotation['G{}'.format(row)] = 0.01
                self.ws_itemized_quotation['H{}'.format(row)] = 0.01
                self.ws_itemized_quotation['I{}'.format(row)] = 0.01
            self.ws_itemized_quotation['G{}'.format(row_number - 3)] = '=sum(G3:G{})'.format(row_number - 4)
            self.ws_itemized_quotation['H{}'.format(row_number - 3)] = '=sum(H3:H{})'.format(row_number - 4)
            self.ws_itemized_quotation['I{}'.format(row_number - 3)] = '=sum(I3:I{})'.format(row_number - 4)
            self.ws_itemized_quotation['J{}'.format(
                row_number - 3)] = '=round((SUM(C{0}:I{0})*3/12*0.0435)*0.8,2)'.format(row_number - 3)
            self.ws_itemized_quotation['K{}'.format(
                row_number - 3)] = '=round(IF(C{0}>50000000,(C{0}-50000000)*0.0075+835000,IF(C{0}>20000000,' \
                                   '(C{0}-20000000)*0.01+535000,IF(C{0}>10000000,(C{0}-10000000)*0.02+335000,' \
                                   'IF(C{0}>5000000,(C{0}-5000000)*0.03+185000,IF(C{0}>2000000,' \
                                   '(C{0}-2000000)*0.035+80000,C{0}*0.04)))))*0.8,2)'.format(row_number - 3)
            self.ws_itemized_quotation['L{}'.format(row_number - 3)] = \
                '=round((M{0}/1.13*0.13-C{0}/1.13*0.13-G{0}/1.06*0.06)*0.9,2)'.format(
                    row_number - 3)

        # 增加条件格式判断
        red_fill = PatternFill(
            start_color='EE1111',
            end_color='EE1111',
            fill_type='solid')
        self.ws_itemized_quotation.conditional_formatting.add('N{}'.format(row_number - 3), CellIsRule(
            operator='notEqual', formula=['M{}'.format(row_number - 3)], fill=red_fill))

        # 合并需要合并单元格
        self.ws_itemized_quotation.merge_cells('A1:M1')
        self.ws_itemized_quotation.merge_cells(
            'B{0}:M{0}'.format(row_number - 1))
        self.ws_itemized_quotation.merge_cells(
            'B{0}:M{0}'.format(row_number - 2))
        self.ws_itemized_quotation.merge_cells('B{0}:M{0}'.format(row_number))

        # 打印设置
        self.ws_itemized_quotation.print_options.horizontalCentered = True
        self.ws_itemized_quotation.print_area = 'A1:M{}'.format(row_number)
        self.ws_itemized_quotation.page_setup.fitToWidth = 1
        self.ws_itemized_quotation.page_setup.orientation = "landscape"
        self.ws_itemized_quotation.page_margins = PageMargins(left=0.25, right=0.25, top=0.75, bottom=0.75, header=0.3,
                                                              footer=0.3)

    def create_summed_quotation(self):
        """创建对内总报价表"""
        self.ws_summed_quotation = self.wb.create_sheet('2.物资对内总报价表', 3)
        colum_title = ['序号', '品名', '规格', '商标', '产地', '生产年份', '{}单价{}(元人民币)'.format(
            self.project.trans, linesep), '数量', '', '{}总价{}(元人民币)'.format(self.project.trans, linesep), '备注']
        title_width = [6, 8, 50, 6, 6, 6, 14, 6, 6, 14, 10]
        colum_number = len(colum_title)
        row_number = len(self.project.commodities) + 4

        # 设置基本的样式
        real_side = Side(style='thin')
        full_border = Border(
            left=real_side,
            right=real_side,
            top=real_side,
            bottom=real_side)
        ctr_alignment = Alignment(
            horizontal='center',
            vertical='center',
            wrap_text=True)
        bold_font = Font(name='宋体', bold=True, size=12)
        normal_font = Font(name='宋体', size=12)
        title_font = Font(name='黑体', size=20)
        left_alignment = Alignment(
            horizontal='left',
            vertical='top',
            wrap_text=True)

        # 初始化表格
        for i in range(colum_number):
            for j in range(1, row_number):
                cell_now = self.ws_summed_quotation.cell(
                    row=j + 1, column=i + 1)
                if j < row_number - 1:
                    cell_now.border = full_border
                if j < 2:
                    cell_now.font = bold_font
                else:
                    cell_now.font = normal_font
                cell_now.alignment = ctr_alignment
        for i in range(len(title_width)):  # 修改列宽
            self.ws_summed_quotation.column_dimensions[
                self.ws_summed_quotation.cell(row=4, column=i + 1).column_letter].width = title_width[i]
        for i in range(2, row_number):  # 修改行高
            self.ws_summed_quotation.row_dimensions[
                self.ws_summed_quotation.cell(row=i, column=1).row].height = 30
        self.ws_summed_quotation.row_dimensions[row_number].height = 45

        # 创建标题行
        self.ws_summed_quotation['A1'].font = title_font
        self.ws_summed_quotation['A1'].alignment = ctr_alignment
        self.ws_summed_quotation['A1'] = '2.物资对内总报价表'
        self.ws_summed_quotation.row_dimensions[1].height = 30

        # 填写表头
        index = 0
        for i in self.ws_summed_quotation['A2':'k2'][0]:
            # print(index+1, i)
            if colum_title[index] != '':
                i.value = colum_title[index]
                # i.font = bold_font
            index += 1

        # 填入数据
        col_relate = [('A', 'A'), ('B', 'B'), ('C', 'G'), ('D', 'F'), ('E', 'L'), ('F', 'P'), ('H', 'D'),
                      ('I', 'E'), ('K', 'X')]
        for row in range(3, row_number - 1):  # 遍历行
            for col in col_relate:  # 根据对应关系设立公式
                cell_now = self.ws_summed_quotation['{}{}'.format(col[0], row)]
                cell_now.value = '=物资输入!{}{}'.format(col[1], row - 1)
                if col[0] == 'C':
                    cell_now.alignment = left_alignment
            self.ws_summed_quotation['J{}'.format(
                row)] = "='3.物资对内分项报价表'!M{}".format(row)
            self.ws_summed_quotation['J{}'.format(
                row)].number_format = '#,##0.00'
            self.ws_summed_quotation['G{}'.format(
                row)] = '=round(J{0}/I{0},2)'.format(row)
            self.ws_summed_quotation['G{}'.format(
                row)].number_format = '#,##0.00'
        self.ws_summed_quotation['A{}'.format(row_number - 1)] = '合计金额'
        self.ws_summed_quotation['C{}'.format(
            row_number - 1)] = '=SUM(J3:J{})'.format(row_number - 1)
        self.ws_summed_quotation['C{}'.format(row_number - 1)].font = bold_font
        self.ws_summed_quotation['C{}'.format(
            row_number - 1)].number_format = '¥#,##0.00'
        self.ws_summed_quotation['C{}'.format(row_number - 1)].alignment = Alignment(
            horizontal='left', vertical='center', wrap_text=True)
        self.ws_summed_quotation['A{}'.format(row_number)] = '注：'
        self.ws_summed_quotation['B{}'.format(row_number)] = "（1）本表所列{}{}单价和总价包括投标人提供上述全部物资" \
            "并承担合同规定全部义务所需的一切费用；{}（2）在备注栏中注明包装" \
            "情况,即包装的单位和数量".format(self.project.trans,
                                  self.project.destination, linesep)
        self.ws_summed_quotation['B{}'.format(row_number)].alignment = Alignment(horizontal='left',
                                                                                 vertical='center', wrap_text=True)

        # 合并需要合并单元格
        self.ws_summed_quotation.merge_cells('A1:K1')
        self.ws_summed_quotation.merge_cells('H2:I2')
        self.ws_summed_quotation.merge_cells(
            'A{0}:B{0}'.format(row_number - 1))
        self.ws_summed_quotation.merge_cells(
            'C{0}:K{0}'.format(row_number - 1))
        self.ws_summed_quotation.merge_cells('B{0}:K{0}'.format(row_number))

        # 打印设置
        self.ws_summed_quotation.print_options.horizontalCentered = True
        self.ws_summed_quotation.print_area = 'A1:K{}'.format(row_number)
        self.ws_summed_quotation.page_setup.fitToWidth = 1
        self.ws_summed_quotation.page_setup.orientation = "landscape"
        self.ws_summed_quotation.page_margins = PageMargins(left=0.25, right=0.25, top=0.75, bottom=0.75, header=0.3,
                                                            footer=0.3)

    def create_examination(self):
        """创建物资选型一览表（非法检）"""
        self.ws_examination = self.wb.create_sheet('7.物资检验一览表（非法检物资）', 3)
        colum_title = ['序号', '品名', 'HS编码', '数量及单位', '', '品牌', '规格、技术参数或配置', '金额', '生产厂商',
                       '供货商', '生产或供货地', '供货联系人及联系电话', '', '出厂日期', '出口港', '检验标准', '检验机构名称',
                       '', '', '备注']
        subcol_title = ['产地或供货地检验（查验）机构', '装运前核验机构', '口岸监装机构']
        title_width = [6, 14, 12, 3, 5, 10, 30, 16,
                       16, 16, 10, 10, 10, 10, 8, 18, 13, 8, 7, 6]

        # 设置基本的样式
        real_side = Side(style='thin')
        full_border = Border(
            left=real_side,
            right=real_side,
            top=real_side,
            bottom=real_side)
        ctr_alignment = Alignment(
            horizontal='center',
            vertical='center',
            wrap_text=True)
        left_alignment = Alignment(
            horizontal='left',
            vertical='top',
            wrap_text=True)
        bold_font = Font(name='宋体', bold=True, size=10)
        normal_font = Font(name='宋体', size=10)
        title_font = Font(name='黑体', size=20)

        # 合并需要合并单元格
        self.ws_examination.merge_cells('D2:E3')
        self.ws_examination.merge_cells('Q2:S2')
        self.ws_examination.merge_cells('L2:M3')
        # self.ws_examination['M2'].border = full_border
        for col in range(1, 21):
            if col not in [4, 5, 12, 13, 17, 18, 19]:
                self.ws_examination.merge_cells(
                    start_row=2, start_column=col, end_row=3, end_column=col)
        for i in range(len(title_width)):  # 修改列宽
            self.ws_examination.column_dimensions[
                self.ws_examination.cell(row=4, column=i + 1).column_letter].width = title_width[i]

        # 初始化表格
        colum_number = len(colum_title)
        row_number = len(self.project.commodities) + 3
        for i in range(colum_number):
            for j in range(row_number):
                cell_now = self.ws_examination.cell(row=j + 1, column=i + 1)
                if j > 0:  # 第一列留下给表头
                    cell_now.border = full_border
                    cell_now.font = normal_font
                    cell_now.alignment = ctr_alignment

        # 创建标题行
        self.ws_examination.merge_cells('A1:T1')
        self.ws_examination['A1'].font = title_font
        self.ws_examination['A1'].alignment = ctr_alignment
        index = 4  # 计算表格序号
        if self.project.is_tech:
            index += 1
        if len(self.project.qc) > 0:
            index += 1
        self.ws_examination['A1'] = '{}.物资检验一览表（非法检物资）'.format(index)
        self.ws_examination.row_dimensions[1].height = 30

        # 填写表头
        index = 0
        for i in self.ws_examination['A2':'T2'][0]:
            # print(index+1, i)
            if colum_title[index] != '':
                i.value = colum_title[index]
                i.font = bold_font
            index += 1
        index = 0
        for cell in self.ws_examination['Q3':'S3'][0]:
            cell.value = subcol_title[index]
            cell.font = bold_font
            index += 1

        # 填入数据
        col_relate = [('A', 'A'), ('B', 'B'), ('C', 'C'), ('D', 'D'), ('E', 'E'), ('F', 'F'), ('G', 'G'), ('H', 'I'),
                      ('I', 'J'), ('J', 'K'), ('K', 'M'), ('L', 'N'), ('M',
                                                                       'O'), ('N', 'P'), ('O', 'Q'), ('P', 'R'),
                      ('Q', 'S'), ('R', 'T'), ('S', 'U'), ('T', 'Y')]
        for row in range(4, row_number + 1):  # 遍历行
            for col in col_relate:  # 根据对应关系设立公式
                cell_now = self.ws_examination['{}{}'.format(col[0], row)]
                cell_now.value = '=物资输入!{}{}'.format(col[1], row - 2)
                if col[0] == 'G':
                    cell_now.alignment = left_alignment
                if col[0] == 'H':
                    cell_now.number_format = '¥#,##0.00'

        # 打印设置
        self.ws_examination.print_area = 'A1:T{}'.format(row_number)
        self.ws_examination.page_setup.fitToWidth = 1
        self.ws_examination.page_setup.orientation = "landscape"
        self.ws_examination.page_margins = PageMargins(left=0.25, right=0.25, top=0.75, bottom=0.7, header=0.3,
                                                       footer=0.3)

    def create_techserve(self):
        """创建技术服务费报价表"""
        self.ws_techserve = self.wb.create_sheet('4.技术服务费报价表', 3)
        colum_title = [
            '序号',
            '费用名称',
            '外币单价',
            '人民币单价',
            '人数',
            '天/次数',
            '外币合计',
            '人民币合计']
        title_width = [6, 17, 14, 14, 8, 10, 14, 20]

        # 设置基本的样式
        real_side = Side(style='thin')
        full_border = Border(
            left=real_side,
            right=real_side,
            top=real_side,
            bottom=real_side)
        slash_border = Border(diagonal=real_side, diagonalDown=True, left=real_side, right=real_side,
                              top=real_side, bottom=real_side)
        ctr_alignment = Alignment(
            horizontal='center',
            vertical='center',
            wrap_text=True)
        bold_font = Font(name='宋体', bold=True, size=12)
        normal_font = Font(name='宋体', size=12)
        title_font = Font(name='黑体', size=20)
        yellow_fill = PatternFill(
            fill_type='solid',
            start_color='FFFF00',
            end_color='FFFF00')

        # 合并需要合并单元格
        self.ws_techserve.merge_cells('B9:H9')
        self.ws_techserve.merge_cells('C14:G14')
        self.ws_techserve.merge_cells('C15:F15')
        self.ws_techserve.merge_cells('C16:F16')
        self.ws_techserve.merge_cells('B19:H19')

        # 初始化表格
        colum_number = len(colum_title)
        row_number = 19
        for i in range(colum_number):
            for j in range(1, row_number):
                cell_now = self.ws_techserve.cell(row=j + 1, column=i + 1)
                if j < 16:  # 第一列留下给表头
                    cell_now.border = full_border
                    cell_now.font = normal_font
                    cell_now.alignment = ctr_alignment
                else:
                    cell_now.font = normal_font
                    cell_now.alignment = Alignment(
                        horizontal='left', vertical='center', wrap_text=True)
        for i in range(len(title_width)):  # 修改列宽
            self.ws_techserve.column_dimensions[
                self.ws_techserve.cell(row=4, column=i + 1).column_letter].width = title_width[i]
        for i in range(2, row_number):  # 修改行高
            self.ws_techserve.row_dimensions[
                self.ws_techserve.cell(row=i, column=1).row].height = 20
        self.ws_techserve.row_dimensions[19].height = 40

        # 打上斜线
        cell_coor = ['C3', 'C4', 'C5', 'C14', 'C15', 'C16', 'D6', 'D7', 'D8', 'D10', 'D11', 'D12',
                     'D13', 'F4', 'F5', 'G3', 'G4', 'G5']
        for cell in cell_coor:
            self.ws_techserve[cell].border = slash_border

        # 创建标题行
        self.ws_techserve.merge_cells('A1:H1')
        self.ws_techserve['A1'].font = title_font
        self.ws_techserve['A1'].alignment = ctr_alignment
        self.ws_techserve['A1'] = '4.技术服务费报价表'
        self.ws_techserve.row_dimensions[1].height = 30

        # 填写表头
        index = 0
        for i in self.ws_techserve['A2':'H2'][0]:
            # print(index+1, i)
            if colum_title[index] != '':
                i.value = colum_title[index]
                i.font = bold_font
            index += 1

        # 填写数据
        col_a = [
            1,
            2,
            3,
            4,
            5,
            6,
            7,
            '7-1',
            '7-2',
            '7-3',
            '7-4',
            8,
            9,
            '']  # 序号
        col_b = ['技术服务补贴', '国内差旅费', '国际机票费', '国外伙食费', '国外住宿费', '国外公杂费', '国际旅途中转费',
                 '中转伙食费', '中转住宿费', '中转公杂费', '中转机场费', '管理费', '其他费用', '共计']  # 费用名称
        for index in range(3, 17):
            self.ws_techserve['A{}'.format(index)] = col_a[index - 3]
            self.ws_techserve['B{}'.format(index)] = col_b[index - 3]
        for num in range(3):  # 填写技术服务费单价
            self.ws_techserve['C{}'.format(
                num + 6)].number_format = '$#,##0.00'
            self.ws_techserve['C{}'.format(
                num + 6)] = self.project.techinfo[2][num]
        # 格式化单元格
        for i in [3, 4, 5, 6, 7, 8, 14, 16]:
            self.ws_techserve['H{}'.format(i)].number_format = '¥#,##0.00'
            if i < 9:
                self.ws_techserve['E{}'.format(i)].number_format = '0'
                self.ws_techserve['E{}'.format(i)] = self.project.techinfo[0]
                self.ws_techserve['F{}'.format(i)].number_format = '0'
                if i in [3, 6, 7, 8]:
                    self.ws_techserve['F{}'.format(
                        i)] = self.project.techinfo[1]
            if i < 6:
                self.ws_techserve['D{}'.format(i)].number_format = '¥#,##0.00'
            if i > 5:
                self.ws_techserve['G{}'.format(i)].number_format = '$#,##0.00'
        if self.project.is_lowprice:
            self.ws_techserve['F7'] = self.project.techinfo[1] - 1
        # 填充备注
        self.ws_techserve['A18'] = '注：'
        self.ws_techserve['B18'] = '（1）100美元='
        self.ws_techserve['C18'].number_format = '0.00"元人民币"'
        self.ws_techserve['C18'] = 700
        self.ws_techserve['C18'].fill = yellow_fill

        self.ws_techserve['B19'] = '（2）国外伙食、住宿、工杂费和中转伙食、住宿、工杂费须按照财行[2013]516号文' \
                                   '和财行[2017]434号文规定的该国（地区）费用标准和币种填报。计算出外币合计后，' \
                                   '应折算出相应人民币合计数'

        # 填充表格
        self.ws_techserve['D3'] = 3000
        if not self.project.is_lowprice:
            self.ws_techserve['D4'] = 2000
            self.ws_techserve['D5'] = 10000
        else:
            self.ws_techserve['D4'] = 0
            self.ws_techserve['D5'] = 0
        self.ws_techserve['D5'].fill = yellow_fill
        for i in range(6, 9):
            self.ws_techserve['G{}'.format(i)] = '=C{0}*E{0}*F{0}'.format(i)
            self.ws_techserve['H{}'.format(i)] = '=G{}*C18/100'.format(i)
        self.ws_techserve['H3'] = '=D3*F3*E3/30'
        self.ws_techserve['H4'] = '=D4*E4'
        self.ws_techserve['H5'] = '=D5*E5'
        self.ws_techserve['H14'] = '=round((H3+H4+H6+H10)*IF(E7<6,0.21,IF(E7<11,0.18,IF(E7<21,0.15,IF(E7<41,0.12,0.09' \
                                   ')))),2)'
        self.ws_techserve['H16'] = '=SUM(H3:H8, H10:H15)'
        self.ws_techserve['G16'] = '=SUM(G6:G8)'

        # 合并需要合并单元格
        self.ws_techserve.merge_cells('B9:H9')
        self.ws_techserve.merge_cells('C14:G14')
        self.ws_techserve.merge_cells('C15:F15')
        self.ws_techserve.merge_cells('C16:F16')
        self.ws_techserve.merge_cells('B19:H19')
        self.ws_techserve.merge_cells('C18:D18')

        # 打印设置
        self.ws_techserve.print_options.horizontalCentered = True
        self.ws_techserve.print_area = 'A1:H{}'.format(row_number)
        self.ws_techserve.page_setup.fitToWidth = 1
        # self.ws_techserve.page_setup.orientation = "landscape"
        self.ws_techserve.page_margins = PageMargins(left=0.7, right=0.7, top=0.75, bottom=0.75, header=0.3,
                                                     footer=0.3)

    def create_lawexam(self):
        """创建物资选型一览表（法检物资）"""
        self.ws_lawexam = self.wb.create_sheet('6.物资检验一览表（法检物资）', 3)
        colum_title = ['序号', '品名', 'HS编码', '数量及单位', '', '品牌', '规格、技术参数或配置', '金额', '生产厂商',
                       '供货商', '生产或供货地', '供货联系人及联系电话', '', '出厂日期', '出口港', '检验标准',
                       '口岸检验机构', '备注']
        title_width = [6, 14, 12, 3, 5, 10, 30, 16,
                       16, 16, 10, 10, 10, 10, 8, 18, 7, 6]

        # 设置基本的样式
        real_side = Side(style='thin')
        full_border = Border(
            left=real_side,
            right=real_side,
            top=real_side,
            bottom=real_side)
        ctr_alignment = Alignment(
            horizontal='center',
            vertical='center',
            wrap_text=True)
        left_alignment = Alignment(
            horizontal='left',
            vertical='top',
            wrap_text=True)
        bold_font = Font(name='宋体', bold=True, size=10)
        normal_font = Font(name='宋体', size=10)
        title_font = Font(name='黑体', size=20)

        # 合并需要合并单元格
        self.ws_lawexam.merge_cells('D2:E2')
        self.ws_lawexam.merge_cells('L2:M2')
        # self.ws_examination['M2'].border = full_border
        for i in range(len(title_width)):  # 修改列宽
            self.ws_lawexam.column_dimensions[
                self.ws_lawexam.cell(row=4, column=i + 1).column_letter].width = title_width[i]

        # 初始化表格
        colum_number = len(colum_title)
        row_number = len(self.project.qc) + 2
        for i in range(colum_number):
            for j in range(row_number):
                cell_now = self.ws_lawexam.cell(row=j + 1, column=i + 1)
                if j > 0:  # 第一列留下给表头
                    cell_now.border = full_border
                    cell_now.font = normal_font
                    cell_now.alignment = ctr_alignment

        # 创建标题行
        self.ws_lawexam.merge_cells('A1:R1')
        self.ws_lawexam['A1'].font = title_font
        self.ws_lawexam['A1'].alignment = ctr_alignment
        index = 4  # 计算表格序号
        if self.project.is_tech:
            index += 1
        self.ws_lawexam['A1'] = '{}.物资检验一览表（法检物资）'.format(index)
        self.ws_lawexam.row_dimensions[1].height = 30

        # 填写表头
        index = 0
        for i in self.ws_lawexam['A2':'R2'][0]:
            # print(index+1, i)
            if colum_title[index] != '':
                i.value = colum_title[index]
                i.font = bold_font
            index += 1

        # 填入数据
        col_relate = [('A', 'A'), ('B', 'B'), ('C', 'C'), ('D', 'D'), ('E', 'E'), ('F', 'F'), ('G', 'G'), ('H', 'I'),
                      ('I', 'J'), ('J', 'K'), ('K', 'M'), ('L', 'N'), ('M',
                                                                       'O'), ('N', 'P'), ('O', 'Q'), ('P', 'R'),
                      ('Q', 'U'), ('R', 'Y')]
        index = 0
        for row in range(3, row_number + 1):  # 遍历行
            for col in col_relate:  # 根据对应关系设立公式
                cell_now = self.ws_lawexam['{}{}'.format(col[0], row)]
                cell_now.value = '=物资输入!{}{}'.format(
                    col[1], self.project.qc[index] + 1)
                if col[0] == 'G':
                    cell_now.alignment = left_alignment
                if col[0] == 'H':
                    cell_now.number_format = '¥#,##0.00'
            index += 1
        num = 0
        for row in self.project.qc:  # 在非法检物资中删除
            self.ws_examination.delete_rows(row - num + 3)
            num += 1

        # 打印设置
        self.ws_lawexam.print_options.horizontalCentered = True
        self.ws_lawexam.print_area = 'A1:R{}'.format(row_number)
        self.ws_lawexam.page_setup.fitToWidth = 1
        self.ws_lawexam.page_setup.orientation = "landscape"
        self.ws_lawexam.page_margins = PageMargins(left=0.25, right=0.25, top=0.75, bottom=0.75, header=0.3,
                                                   footer=0.3)

    def create_training(self):
        """创建来华培训费报价表"""
        self.ws_training = self.wb.create_sheet('5.来华培训费报价表', 3)
        colum_title = [
            '序号',
            '费用名称',
            '',
            '费用计算方式',
            '',
            '',
            '人民币（元）',
            '其中含购汇人民币限额']
        title_width = [6, 14, 7, 14, 7, 12, 14, 12]
        # 设置基本的样式
        real_side = Side(style='thin')
        full_border = Border(
            left=real_side,
            right=real_side,
            top=real_side,
            bottom=real_side)
        slash_border = Border(diagonal=real_side, diagonalDown=True, left=real_side, right=real_side,
                              top=real_side, bottom=real_side)
        ctr_alignment = Alignment(
            horizontal='center',
            vertical='center',
            wrapText=True)
        right_alignment = Alignment(
            horizontal='right',
            vertical='center',
            wrapText=True)
        left_alignment = Alignment(
            horizontal='left',
            vertical='center',
            wrapText=True)
        bold_font = Font(name='宋体', bold=True, size=12)
        normal_font = Font(name='宋体', size=12)
        title_font = Font(name='黑体', bold=True, size=14)
        yellow_fill = PatternFill(
            fill_type='solid',
            start_color='FFFF00',
            end_color='FFFF00')

        # 初始化表格
        colum_number = len(colum_title)
        row_number = 20
        for i in range(colum_number):
            for j in range(1, row_number):  # 第一列留下给表头
                cell_now = self.ws_training.cell(row=j + 1, column=i + 1)
                if j < 17:  # 给主体cell设置样式
                    cell_now.border = full_border
                    cell_now.font = normal_font
                    if i == 6:  # 数字列右对齐
                        cell_now.alignment = right_alignment
                    else:
                        cell_now.alignment = ctr_alignment
                else:  # 最后两行左对齐
                    cell_now.font = normal_font
                    cell_now.alignment = left_alignment
        for i in range(len(title_width)):  # 修改列宽
            self.ws_training.column_dimensions[
                self.ws_training.cell(row=4, column=i + 1).column_letter].width = title_width[i]
        for i in range(2, row_number):  # 修改行高
            self.ws_training.row_dimensions[
                self.ws_training.cell(row=i, column=1).row].height = 20
        self.ws_training.row_dimensions[20].height = 30

        # 打上斜线
        cell_coor = ['D14', 'E14', 'F14']
        for cell in cell_coor:
            self.ws_training[cell].border = slash_border

        # 创建标题行
        self.ws_training['A1'].font = title_font
        self.ws_training['A1'].alignment = ctr_alignment
        index = 4
        if self.project.is_tech:
            index += 1
        self.ws_training['A1'] = '{}.来华培训费报价表'.format(index)
        self.ws_training.row_dimensions[1].height = 30
        self.ws_training.merge_cells('A1:H1')

        # 填写表头
        index = 0
        for i in self.ws_training['A2':'H2'][0]:
            # print(index, i)
            if colum_title[index] != '':
                i.value = colum_title[index]
                i.font = bold_font
            index += 1
        self.ws_training['D3'].value = '标准'
        self.ws_training['E3'].value = '人数'
        self.ws_training['F3'].value = '天（次）数'
        self.ws_training['D3'].font = bold_font
        self.ws_training['E3'].font = bold_font
        self.ws_training['F3'].font = bold_font

        # 填写数据
        col_a = ['一', '二', 1, 2, 3, 4, 5, 6, '三', '四', 1, 2, '', '五', '']  # 序号
        col_b = ['培训费', '接待费', '日常伙食费', '住宿费', '宴请费', '零用费', '小礼品费', '人身意外伤害保险',
                 '国际旅费', '管理费', '承办管理费', '管理人员费', '', '合计']  # 费用名称
        for index in range(4, 18):  # 填写前两列
            self.ws_training['A{}'.format(index)] = col_a[index - 4]
            if isinstance(col_a[index - 4], int):
                self.ws_training['A{}'.format(
                    index)].alignment = right_alignment
            self.ws_training['B{}'.format(index)] = col_b[index - 4]
        self.ws_training['C15'].value = '伙食费'
        self.ws_training['C16'].value = '住宿费'
        # 填写E列
        for num in range(6, 13):  # 填写培训人数
            self.ws_training['E{}'.format(num)].number_format = '0'
            self.ws_training['E{}'.format(
                num)].value = self.project.training_num
        self.ws_training['E4'].number_format = '0'
        self.ws_training['E4'].value = self.project.training_num
        self.ws_training['E15'].number_format = '0'
        num = self.project.training_num
        if num < 10:
            res = 1
        else:
            if num % 10 == 0:
                res = num / 10
            else:
                res = num / 10 + 1
        self.ws_training['E15'].value = res
        self.ws_training['E16'].number_format = '0'
        self.ws_training['E16'].value = '=E15'
        # 填写D列
        for num in [4, 6, 7, 9, 15, 16]:
            self.ws_training['D{}'.format(num)].number_format = '0"元/人*天"'
        for num in range(10, 13):
            self.ws_training['D{}'.format(num)].number_format = '0"元/人"'
        self.ws_training['D8'].number_format = '0"元/人*次"'
        self.ws_training['D4'] = 320
        self.ws_training['D6'] = 140
        self.ws_training['D7'] = 300
        self.ws_training['D8'] = 150
        self.ws_training['D9'] = 80
        self.ws_training['D10'] = 200
        self.ws_training['D11'] = 100
        self.ws_training['D12'] = 0
        self.ws_training['D12'].fill = yellow_fill
        self.ws_training['D15'] = 140
        self.ws_training['D16'] = 300

        # 填写F列
        for i in [4, 6, 7, 8, 9, 10, 11, 12, 15, 16]:
            self.ws_training['F{}'.format(i)].number_format = '0'
        self.ws_training['F4'] = self.project.training_days
        self.ws_training['F6'] = self.project.training_days
        self.ws_training['F7'] = self.project.training_days - 1
        self.ws_training['F9'] = self.project.training_days
        self.ws_training['F11'] = '-'
        self.ws_training['F12'] = '-'
        self.ws_training['F15'] = self.project.training_days
        self.ws_training['F16'] = self.project.training_days
        if self.project.is_lowprice:
            self.ws_training['F8'] = 0
            self.ws_training['F10'] = 0
        else:
            self.ws_training['F8'] = 1
            self.ws_training['F10'] = 1


        # 填写G列
        for i in range(4, 18):
            self.ws_training['G{}'.format(i)].number_format = '¥#,##0.00'
            if i in [5, 13]:
                pass
            elif i in (11, 12):
                self.ws_training['G{}'.format(i)] = '=D{0}*E{0}'.format(i)
            elif i == 14:
                self.ws_training['G{}'.format(
                    i)] = '=ROUND((SUM(G4,G6:G11))*0.06,2)'
            elif i == 17:
                self.ws_training['G{}'.format(i)] = '=SUM(G4:G16)'
            else:
                self.ws_training['G{}'.format(i)] = '=D{0}*E{0}*F{0}'.format(i)

        # 填充备注
        self.ws_training['A19'] = '注：'
        self.ws_training['B19'] = '（1）100美元='
        self.ws_training['C19'].number_format = '0.00"元人民币"'
        self.ws_training['C19'] = 700
        self.ws_training['C19'].fill = yellow_fill
        self.ws_training['B20'] = '（2）上述费用参照财政部（2008）第2号文举办援外培训班费用开支标准和财务管理办法给定的费用标准报价'

        # 合并需要合并单元格
        self.ws_training.merge_cells('D2:F2')
        self.ws_training.merge_cells('A2:A3')
        self.ws_training.merge_cells('B2:C3')
        self.ws_training.merge_cells('G2:G3')
        self.ws_training.merge_cells('H2:H3')
        self.ws_training.merge_cells('D5:G5')
        self.ws_training.merge_cells('D13:G13')
        self.ws_training.merge_cells('D14:F14')
        self.ws_training.merge_cells('B17:F17')
        self.ws_training.merge_cells('C19:D19')
        self.ws_training.merge_cells('B20:H20')
        for i in range(4, 15):
            self.ws_training.merge_cells('B{0}:C{0}'.format(i))
        self.ws_training.merge_cells('B15:B16')
        self.ws_training.merge_cells('A15:A16')

        # 打印设置
        self.ws_training.print_options.horizontalCentered = True
        self.ws_training.print_area = 'A1:H{}'.format(row_number)
        self.ws_training.page_setup.fitToWidth = 1
        # self.ws_training.page_setup.orientation = "landscape"
        self.ws_training.page_margins = PageMargins(left=0.25, right=0.25, top=0.75, bottom=0.75, header=0.3,
                                                    footer=0.3)


class Content(object):
    """通过project实例创建目录"""

    # 设置公用样式
    title_font = Font(name='宋体', size=24, bold=True)
    header_font = Font(name='仿宋_GB2312', size=14, bold=True)
    normal_font = Font(name='仿宋_GB2312', size=14)
    header_border = Border(bottom=Side(style='medium'))
    normal_border = Border(bottom=Side(style='thin', color='80969696'))
    ctr_alignment = Alignment(
        horizontal='center',
        vertical='center',
        wrap_text=True)
    left_alignment = Alignment(
        horizontal='left',
        vertical='center',
        wrap_text=True,
        indent=1)
    margin = PageMargins()

    def __init__(self, project):
        self.project = project
        self.wb = Workbook()
        self.ws_lob = None
        self.ws_tech = None
        self.ws_qual = None
        self.ws_eco_com = None
        self.my_title = 'content'

    def create_all(self):
        """生成目录总方法"""

        self.create_lob()
        self.create_tech()
        self.create_qual()
        self.create_eco_com()
        self.wb.save('目录—{}.xlsx'.format(self.project.name))

    def create_lob(self):
        """创建投标函目录"""
        self.ws_lob = self.wb.create_sheet('投标函', 0)
        col_titles = ['序号', '内容', '页码']
        content = [['一', '投标函'], ['二', '法定代表人身份证明书'], ['三', '法定代表人授权书'],
                   ['四', '援外物资项目投标廉政承诺书'], ['五', '企业内控承诺'], ['六', '投标保证金银行保函']]
        col_width = [10, 60, 10]
        col_num = 3
        row_num = 8

        # 初始化表格
        for i in range(row_num):
            for j in range(col_num):
                cell_now = self.ws_lob.cell(row=i + 1, column=j + 1)
                self.ws_lob.row_dimensions[i + 1].height = 45  # 修改行高
                if i > 0:
                    if i == 1:
                        cell_now.font = Content.header_font
                        cell_now.alignment = Content.ctr_alignment
                        cell_now.border = Content.header_border
                        cell_now.value = col_titles[j]
                    else:
                        cell_now.font = Content.normal_font
                        if j == 1:
                            cell_now.alignment = Content.left_alignment
                            cell_now.value = content[i - 2][1]
                        else:
                            cell_now.alignment = Content.ctr_alignment
                            if j == 0:
                                cell_now.value = content[i - 2][0]
                            elif j == 2:
                                cell_now.value = i - 1
                        if i != row_num - 1:
                            cell_now.border = Content.normal_border
        letters = string.ascii_uppercase
        for i in range(col_num):  # 修改列宽
            self.ws_lob.column_dimensions[letters[i]].width = col_width[i]

        # 填写抬头
        self.ws_lob.merge_cells('A1:C1')
        header = self.ws_lob['A1']
        header.font = Content.title_font
        header.alignment = Content.ctr_alignment
        header.value = '目  录'
        self.ws_lob.row_dimensions[1].height = 50

        # 打印设置
        self.ws_lob.print_options.horizontalCentered = True
        self.ws_lob.print_area = 'A1:C9'
        self.ws_lob.page_setup.fitToWidth = 1
        self.ws_lob.page_margins = Content.margin

    def create_tech(self):
        """创建技术标目录"""
        self.ws_tech = self.wb.create_sheet('技术标', 0)
        col_titles = ['序号', '内容', '页码']
        # 存放固定内容
        content = [
            '技术偏离表',
            '物资选型部分',
            '（一）物资选型一览表',
            '（二）物资生产供货企业信息及技术资料',
            '包装方案',
            '运输方案和计划',
            '附：承运人出具的书面运输承诺书',
            '物资检验服务方案',
            '附：检验机构服务方案及实力情况',
            '重点和难点问题应对方案']
        # 存放中文序号
        num = ['一', '二', '三', '四', '', '五', '', '六', '七', '八', '九', '十']
        col_width = [10, 60, 10]
        col_num = 3

        # 确定行数
        com_num = len(self.project.commodities)
        row_num = com_num + 12
        if self.project.is_tech:
            row_num += 1
            content.append('技术服务方案及相关材料')
        if self.project.is_qa:
            row_num += 1
            content.append('售后服务方案及相关材料')
        if self.project.is_cc:
            row_num += 1
            content.append('来华培训方案及相关材料')

        # 创建专用样式
        third_alignment = Alignment(
            horizontal='left',
            vertical='center',
            wrap_text=True,
            indent=3)
        third_font = Font(name='仿宋_GB2312', size=12)

        # 填写抬头
        self.ws_tech.merge_cells('A1:C1')
        header = self.ws_tech['A1']
        header.font = Content.title_font
        header.alignment = Content.ctr_alignment
        header.value = '目  录'
        self.ws_tech.row_dimensions[1].height = 50

        # 初始化表格,双循环扫描先行后列扫描表格
        for i in range(1, row_num):
            for j in range(col_num):
                cell_now = self.ws_tech.cell(row=i + 1, column=j + 1)
                self.ws_tech.row_dimensions[i + 1].height = 30  # 修改行高
                # 判断行数来确定应用的字体和样式
                if i == 1:  # 表头行样式填写
                    cell_now.font = Content.header_font
                    cell_now.alignment = Content.ctr_alignment
                    cell_now.border = Content.header_border
                    cell_now.value = col_titles[j]
                elif 1 < i < 4:  # 头两行
                    cell_now.font = Content.normal_font
                    cell_now.border = Content.normal_border
                    if j == 1:
                        cell_now.alignment = Content.left_alignment
                        cell_now.value = content[i - 2]
                    else:
                        cell_now.alignment = Content.ctr_alignment
                        if j == 0:
                            cell_now.value = num[i - 2]
                elif i == 4 or i == 5:  # 3、4行
                    cell_now.font = Content.normal_font
                    cell_now.border = Content.normal_border
                    if j == 1:
                        cell_now.alignment = Content.left_alignment
                        cell_now.value = content[i - 2]
                    else:
                        cell_now.alignment = Content.ctr_alignment
                elif 5 < i < com_num + 6:  # 填写物资名称
                    cell_now.font = third_font
                    cell_now.border = Content.normal_border
                    if j == 1:
                        cell_now.alignment = third_alignment
                        cell_now.value = '{}、{}'.format(
                            i - 5, self.project.commodities[i - 5][0])
                    else:
                        cell_now.alignment = Content.ctr_alignment
                else:   # 其余的一起填写
                    cell_now.font = Content.normal_font
                    if j == 1:
                        cell_now.alignment = Content.left_alignment
                        cell_now.value = content[i - com_num - 2]
                    else:
                        cell_now.alignment = Content.ctr_alignment
                        if j == 0:
                            cell_now.value = num[i - com_num - 4]
                    if i != row_num - 1:
                        cell_now.border = Content.normal_border
        for i in (9, 11):  # 修改两处格式
            self.ws_tech.cell(row=com_num + i, column=2).font = third_font
            self.ws_tech.cell(
                row=com_num + i,
                column=2).alignment = third_alignment
        letters = string.ascii_uppercase
        for i in range(col_num):  # 修改列宽
            self.ws_tech.column_dimensions[letters[i]].width = col_width[i]

        # 打印设置
        self.ws_tech.print_options.horizontalCentered = True
        self.ws_tech.print_area = 'A1:C{}'.format(row_num)
        self.ws_tech.page_setup.fitToWidth = 1
        self.ws_tech.page_margins = PageMargins(
            top=0.5, bottom=0.5, header=0.1, footer=0.1)

    def create_eco_com(self):
        self.ws_eco_com = self.wb.create_sheet('经济和商务', 0)
        col_titles = ['序号', '内容', '页码']
        content = [
            '',
            '报价总表',
            '物资对内总报价表',
            '物资对内分项报价表',
            '',
            '守信企业确认书',
            '上一年度进出口额证明材料',
            '项目负责人援外物资项目或境外项目业绩说明',
            '同类物资业绩一览表'
        ]
        col_width = [10, 60, 10]
        num = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十']
        col_num = 3

        # 确定行数
        row_num = 11
        if len(self.project.qc) == 0:
            row_num += 1
            content.insert(4, '物资检验一览表（非法检物资）')
        else:
            if len(self.project.qc) == len(self.project.commodities):
                row_num += 1
                content.insert(4, '物资检验一览表（法检物资）')
            else:
                row_num += 2
                content.insert(4, '物资检验一览表（法检物资）')
                content.insert(4, '物资检验一览表（非法检物资）')
        if self.project.is_cc:
            row_num += 1
            content.insert(4, '来华培训费报价表')
        if self.project.is_tech:
            row_num += 1
            content.insert(4, '技术服务费报价表')
        if self.project.is_lowprice:
            row_num -= 5
            for i in range(5):
                content.pop()

        # 初始化表格
        for i in range(1, row_num):
            for j in range(col_num):
                cell_now = self.ws_eco_com.cell(row=i + 1, column=j + 1)
                self.ws_eco_com.row_dimensions[i + 1].height = 45  # 修改行高
                # 判断行数来确定应用的字体和样式
                if i == 1:  # 表头行样式填写
                    cell_now.font = Content.header_font
                    cell_now.alignment = Content.ctr_alignment
                    cell_now.border = Content.header_border
                    cell_now.value = col_titles[j]
                else:  # 其余的一起填写
                    cell_now.font = Content.normal_font
                    if j == 1:
                        cell_now.alignment = Content.left_alignment
                        cell_now.value = content[i - 2]
                    else:
                        cell_now.alignment = Content.ctr_alignment
                    if i in (2, row_num - 5, row_num - 6) and not self.project.is_lowprice:
                        cell_now.border = Content.header_border
                    elif i != row_num - 1:
                        cell_now.border = Content.normal_border
        letters = string.ascii_uppercase
        for i in range(col_num):  # 修改列宽
            self.ws_eco_com.column_dimensions[letters[i]].width = col_width[i]

        # 填写序号
        self.ws_eco_com['A3'] = '经济标部分'
        self.ws_eco_com['A3'].font = Content.header_font
        if not self.project.is_lowprice:
            self.ws_eco_com['A{}'.format(row_num - 4)] = '商务标部分'
            self.ws_eco_com['A{}'.format(
                row_num - 4)].font = Content.header_font

        # 填写序号
        if self.project.is_lowprice:
            for i in range(4, row_num + 1):
                self.ws_eco_com['A{}'.format(i)] = num[i - 4]
        else:
            for i in range(4, row_num - 4):
                self.ws_eco_com['A{}'.format(i)] = num[i - 4]
            for i in range(row_num - 3, row_num + 1):
                self.ws_eco_com['A{}'.format(i)] = num[i - row_num + 3]

        # 合并小标题
        self.ws_eco_com.merge_cells('A3:C3')
        if not self.project.is_lowprice:
            self.ws_eco_com.merge_cells('A{0}:C{0}'.format(row_num - 4))

        # 填写抬头
        self.ws_eco_com.merge_cells('A1:C1')
        header = self.ws_eco_com['A1']
        header.font = Content.title_font
        header.alignment = Content.ctr_alignment
        header.value = '目  录'
        self.ws_eco_com.row_dimensions[1].height = 50

        # 打印设置
        self.ws_eco_com.print_options.horizontalCentered = True
        self.ws_eco_com.print_area = 'A1:C{}'.format(row_num)
        self.ws_eco_com.page_setup.fitToWidth = 1
        # self.ws_eco_com.page_margins = PageMargins(
        #     top=0.5, bottom=0.5, header=0.1, footer=0.1)

    def create_qual(self):
        self.ws_qual = self.wb.create_sheet('资格后审', 0)
        col_titles = ['序号', '内容', '页码']
        content = [['一', '资格后审申请函'], ['二', '证明文件']]
        content2 = [
            '投标人的法人营业执照、援外物资项目实施企业资格证明文件',
            '法定代表人证明书和授权书（复印件）',
            '未受到行政处罚且无不良行为记录的声明函',
            '技术资质证明文件',
            '财务审计报告',
            '依法缴纳社会保障资金的证明和依法缴纳税收的证明',
            '关联企业声明',
            '受托检验机构营业执照副本复印件',
            '受托检验机构从事进出口商品检验鉴定业务的许可文件的复印件',
            '受托检验机构获得CNAS认可和CMA计量认证资质文件的复印件',
            '受托检验机构获得ISO / IEC17020检验机构运行体系认证证书复印件',
            '受托检验机构分公司 / 分支机构或实验室 / 合作实验室的证明材料',
            '受托检验机构向我公司出具的《物资检验、核验和监装承诺书》',
            '受托检验机构向你中心出具的《受托检验机构声明函》',
            '其它',
        ]
        col_width = [8, 75, 8]
        col_num = 3
        row_num = 19

        # 创建专用样式
        special_alignment = Alignment(
            horizontal='left',
            vertical='center',
            wrap_text=True,
            indent=0)
        special_font = Font(name='仿宋_GB2312', size=12)

        # 初始化表格
        for i in range(1, row_num):
            for j in range(col_num):
                cell_now = self.ws_qual.cell(row=i + 1, column=j + 1)
                self.ws_qual.row_dimensions[i + 1].height = 35  # 修改行高
                # 判断行数来确定应用的字体和样式
                if i == 1:  # 表头行样式填写
                    cell_now.font = Content.header_font
                    cell_now.alignment = Content.ctr_alignment
                    cell_now.border = Content.header_border
                    cell_now.value = col_titles[j]
                elif 1 < i < 4:  # 头两行
                    cell_now.font = Content.normal_font
                    cell_now.border = Content.normal_border
                    if j == 1:
                        cell_now.alignment = Content.left_alignment
                        cell_now.value = content[i - 2][1]
                    else:
                        cell_now.alignment = Content.ctr_alignment
                        if j == 0:
                            cell_now.value = content[i - 2][0]
                else:   # 其余的一起填写
                    cell_now.font = special_font
                    if j == 1:
                        cell_now.alignment = special_alignment
                        cell_now.value = '{}、{}'.format(i - 3, content2[i - 4])
                    else:
                        cell_now.alignment = Content.ctr_alignment
                    if i != row_num - 1:
                        cell_now.border = Content.normal_border
        letters = string.ascii_uppercase
        for i in range(col_num):  # 修改列宽
            self.ws_qual.column_dimensions[letters[i]].width = col_width[i]

        # 填写抬头
        self.ws_qual.merge_cells('A1:C1')
        header = self.ws_qual['A1']
        header.font = Content.title_font
        header.alignment = Content.ctr_alignment
        header.value = '目  录'
        self.ws_qual.row_dimensions[1].height = 50

        # 打印设置
        self.ws_qual.print_options.horizontalCentered = True
        self.ws_qual.print_area = 'A1:C{}'.format(row_num)
        self.ws_qual.page_setup.fitToWidth = 1
        self.ws_qual.page_margins = PageMargins(
            top=0.5, bottom=0.5, header=0.1, footer=0.1)


class Cover(object):
    """通过project实例创建封面"""
    # parts = ['正本', '副本一', '副本二']
    # sections = ['投标函部分', '技术标部分', '经济标和商务标部分', '资格证明文件部分']
    # name = self.project.name
    # code = '招标编号：{}'.format(self.project.code)
    # ccoec = '投标人：中国海外经济合作有限公司'
    # date = '投标日期：{}'.format(self.project.date)

    def __init__(self, project):
        self.project = project

        self.parts = ['正本', '副本一', '副本二']
        self.sections = ['投标函部分', '技术标部分', '经济标和商务标部分', '资格证明文件部分']
        self.name = self.project.name
        self.code = '招标编号：{}'.format(self.project.code)
        self.ccoec = '投标人：中国海外经济合作有限公司'
        self.date = '投标日期：{}'.format(self.project.date)
        if self.project.is_lowprice:
            self.sections.pop(2)
            self.sections.insert(2, '经济标部分')
        self.doc = Document()

    def insert_mid_words(self, doc, word, loc=text.WD_PARAGRAPH_ALIGNMENT.CENTER):
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

    def insert_blank_line(self, doc):
        # 正本副本文字后面的两个空行
        blank_para = doc.add_paragraph()
        blank_para.paragraph_format.alignment = text.WD_PARAGRAPH_ALIGNMENT.CENTER
        blank_para.paragraph_format.line_spacing_rule = text.WD_LINE_SPACING.ONE_POINT_FIVE

    def insert_big_word(self, doc, word):
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

    def insert_small_word(self, doc, word):
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

    def make_page(self, doc, part, section, name, code, ccoec, date, last=False):
        self.insert_mid_words(doc, part, text.WD_PARAGRAPH_ALIGNMENT.RIGHT)
        self.insert_blank_line(doc)
        self.insert_mid_words(doc, name)
        self.insert_blank_line(doc)
        self.insert_big_word(doc, '投标文件')
        self.insert_mid_words(doc, '（' + section + '）')
        for i in range(4):
            self.insert_blank_line(doc)
        for x in [code, ccoec, date]:
            self.insert_small_word(doc, x)
        if not last:
            doc.add_page_break()

    def generate(self):
        for part in self.parts:
            for section in self.sections:
                last = False
                if part == '副本二' and section == '资格证明文件部分':
                    last = True
                self.make_page(self.doc, part, section, self.name, self.code, self.ccoec, self.date, last)
        self.doc.save('封面_{}.docx'.format(self.name))


def separate_wb():
    wb_pattern = re.compile('^投标报价表\-?\w*(\.xlsx)$')
    for doc in listdir():
        if re.match(wb_pattern, doc):
            filename = doc
    sheet_pattern = re.compile('^[0-9]\.\w*')
    my_wb = load_workbook(filename, data_only=True)
    name_list = []
    for sheet in my_wb:
        if re.match(sheet_pattern, sheet.title):
            name_list.append(sheet.title)
    for name in name_list:
        wb_now = load_workbook(filename, data_only=True)
        ws_now = wb_now[name]
        for sheet in wb_now:
            if sheet.title != ws_now.title:
                wb_now.remove(sheet)
        wb_now.save('{}.xlsx'.format(name))


def make_dir():
    project = Project('\\'.join([os.path.abspath(''), 'project.docx']))
    project_name = project.name
    goods = []

    level_1 = [u'1.投标函部分', u'2.技术标部分', u'3.经济标部分']
    level_2 = [u'.技术偏离表', u'.物资选型部分', u'.包装方案', u'.运输方案和计划', u'.物资检验方案', u'.重点和难点问题应对方案']
    # level_3 = [u'1.物资选型一览表', u'2.各项物资参数响应表', u'3.各项物资供货授权及质量保证书',
    # u'4.各项物资生产企业信息表',u'5.各项物资选型技术资料']

    if project.is_tech:
        level_2.append(u'.技术服务方案')
    if project.is_qa:
        level_2.append(u'.售后服务方案')
    if project.is_cc:
        level_2.append(u'.来华培训方案')
    if not project.is_lowprice:
        level_1.append(u'4.商务标部分')


    Path1 = '\\'.join([os.path.abspath(''), u'空白本-{}'.format(project_name)])
    Path2 = '\\'.join([Path1, u'2.技术标部分'])
    Path3 = '\\'.join([Path2, u'2.物资选型部分'])
    # Path4 = '\\'.join([Path3, u'3.各项物资选型技术资料'])

    for i in level_1:
        os.makedirs('\\'.join([Path1, i]))

    for i in range(len(level_2)):
        temp = str(i + 1) + level_2[i]
        PathHere = '\\'.join([Path2, temp])
        os.mkdir(PathHere)

    # for i in level_3:
    # 	os.mkdir('\\'.join([Path3, i]).encode('gbk'))
    PathHere = '\\'.join([Path3, u'0.物资选型一览表'])
    os.mkdir(PathHere)

    seq = list(project.commodities.keys())
    seq.sort()
    for i in seq:
        Seq = str(i) + '.'
        temp = Seq + project.commodities[i][0].split('\n')[0]
        PathHere = '\\'.join([Path3, temp])
        os.mkdir(PathHere)


def main_loop(tips):
    while 1:
        user_input = input(tips)
        if user_input == '':
            print('<<< 您输入的指令为空，请重新输入 >>>')
        else:
            for i in user_input:
                if i not in '12345':
                    print('<<< 您输入非法指令，请重新输入 >>>')
                    break
            else:
                break
    project = Project('project.docx')
    quota = Quotation(project)
    content = Content(project)
    cover = Cover(project)
    func_dict = {'1': quota.create_all, '2': content.create_all, '3': cover.generate, '4': make_dir, '5': separate_wb}
    for func in user_input:
        if func == '1' and os.path.exists('投标报价表-{}.xlsx'.format(project.name)):
            temp_input = input('!!!该报价表已存在，请确认是否需要覆盖（Y/N）!!! >>> ')
            if temp_input in 'Yy':
                func_dict['1']()
        else:
            func_dict[func]()
    input('<<< 程序已经运行完成，按任意键退出 >>>')


def main_func(tips):
    date_init = datetime.strptime('2020-10-01', '%Y-%m-%d').date()
    date_now = datetime.now().date()
    limited_days = int(cmath.sqrt(len(popen('hostname').read())).real * 10) + 100
    delta = date_now - date_init
    if delta.days < limited_days:
        try:
            main_loop(tips)
        except Exception as e:
            input('<<< 出现异常：{} >>>'.format(e))
    else:
        input('<<< 出现异常：Out Of Date >>>')

tips = """
请按照序号选择你需要的功能：
1、生成报价表
2、生成目录
3、生成封面
4、生成空白本文件夹结构
5、拆分报价表
>>> """


if __name__ == "__main__":
    main_func(tips)
