from django.shortcuts import render

# Create your views here.
from django.http import HttpResponse
import xlrd
import os
from docx import Document

file_dir = os.path.join(os.path.abspath(os.path.dirname(__file__)),'files')
def index(request):
    read_execl()
    return HttpResponse("Hello, world. You're at the polls index.")

def execl2word(request):
    read_execl()
    file_name = '''生成word成功,访问路径{dir}'''.format(dir=os.path.join(file_dir, 'demo.docx'))
    return HttpResponse(file_name)

def read_execl():
    filename =os.path.join(file_dir ,'person.xls')
    worksheet = xlrd.open_workbook(filename)
    sh = worksheet.sheet_by_index(0)
    table = []
    for rx in range(sh.nrows):
        if rx == 0:
            continue
        person = []
        for cx in range(sh.ncols):
            col_value = sh.cell_value(rowx=rx, colx=cx)
            person.append(col_value)
        table.append(person)

    for i in table:
        print(i)
        create_word(i[0], i[1], i[2], i[3], i[4])



def create_word(id_card,family,number,start_date,end_date):
    words="""
    广州市居民家庭经济状况核对需求书
(模板)

广州市各级核对机构：
根据相关规定，我单位需对             (身份证号码：{id_card}
	) 及 其 家 庭 成 员 (   {family}             
	) 共 (  {number}    )人的家庭经济状况进行核对。现根据
我市核对政策规定并经申请人及其家庭成员授权，特提请你
机构对上述人员  {start_y}    年  {start_m}    月至 {end_y}   年 {end_m}     月的收入、截止
查询之日所拥有的财产等家庭经济状况及其他相关情况进
行核对。
我单位承诺对提请核对的人员身份证明信息和授权文
件的有效性和真实性负责。



业务部门(盖章):
年      月



备注：涉及实物财产评估的，需提供《实物财产价格评估需求表》并加盖公章。
    """.format(id_card=id_card,family=family,number=number,
               start_y=start_date.split('.')[0],start_m=start_date.split('.')[1],
               end_y=end_date.split('.')[0], end_m=end_date.split('.')[1]
               )
    filename = os.path.join(file_dir, 'demo.docx')
    document = Document()
    p = document.add_paragraph(words)
    document.add_page_break()
    document.save(filename)


