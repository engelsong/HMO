# coding:utf-8

import os.path
from docx import Document
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

if __name__ == "__main__":
    date_init = datetime.strptime('2020-10-01', '%Y-%m-%d').date()
    date_now = datetime.now().date()
    limited_days = int(cmath.sqrt(len(popen('hostname').read())).real * 10) + 100
    delta = date_now - date_init
    if delta.days < limited_days:
        make_dir()
    else:
        raise UserWarning('Out Of Date')
