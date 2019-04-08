#coding=gb18030
#author:zhoub
from xlrd import open_workbook,xldate
import string
import tempfile
from Tkinter import *
import os
import arrow
from pyh import *
from win32api import GetSystemMetrics,ShellExecute
from multiprocessing import Pool,freeze_support,cpu_count

def book_list():  ## 自以下文档提取结果
    book1=[r'\\SJSTORAGE\Dept_Operaton_CML_PE\APE\APE Report File\ECR & PCR 編號跟進.XLS\2019\ECR编号跟进-2019.xls','ECR']
    book2=[r'\\SJSTORAGE\Dept_Operaton_CML_PE\APE\APE Report File\ECR & PCR 編號跟進.XLS\2018\ECR编号跟进-2018.xls','ECR']
    book3=[r'\\SJSTORAGE\Dept_Operaton_CML_PE\APE\APE Report File\ECR & PCR 編號跟進.XLS\2017\ECR编号跟进-2017.xls','ECR']
    book4=[r'\\SJSTORAGE\Dept_Operaton_CML_PE\APE\APE Report File\ECR & PCR 編號跟進.XLS\2016\ECR編號跟進-2016.xls','ECR']
    book5=[r'\\SJSTORAGE\Dept_Operaton_CML_PE\APE\APE Report File\ECR & PCR 編號跟進.XLS\2019\PCR編號跟進-2019.xls','PCR']
    book6=[r'\\SJSTORAGE\Dept_Operaton_CML_PE\APE\APE Report File\ECR & PCR 編號跟進.XLS\2018\PCR編號跟進-2018.xls','PCR']
    book7=[r'\\SJSTORAGE\Dept_Operaton_CML_PE\APE\APE Report File\ECR & PCR 編號跟進.XLS\2017\PCR編號跟進-2017.xls','PCR']
    book8=[r'\\SJSTORAGE\Dept_Operaton_CML_PE\APE\APE Report File\ECR & PCR 編號跟進.XLS\2016\PCR編號跟進-2016.xls','PCR']
    book_lists=[book1,book2,book3,book4,book5,book6,book7,book8]
    return book_lists

def CreateLists(book,model):
    k = 0
    br = []
    wb=open_workbook(book[0])
    sh=wb.sheet_by_name(book[1])
    for rows_ in sh.col_values(4):
        if model in rows_:
            try:                     ## 判断有无链接
                link=sh.hyperlink_map.get((k,0))
                lists_link=link.url_or_path
                lists_value=sh.cell(k,0).value
                column_1=lists_link
            except:
                column_1=sh.cell(k,0).value
            try:
                date=xldate.xldate_as_datetime(sh.cell(k,1).value, 0)   ## 单元格日期
                da=arrow.get(date)
            except:
                da='1999/9/9'  ## 错误日期填充
            column_2=da.format('YYYY-M-D')
            column_3=sh.cell(k,2).value
            column_4=sh.cell(k,3).value
            column_5=sh.cell(k,4).value
            column_6=sh.cell(k,5).value
            column_7=sh.cell(k,6).value
            arr = [column_1,column_2,column_3,column_4,column_5,column_6,column_7]
            br.append(arr)
        k+=1
    return br

def write_html(ar):  ## 生成html页面
    page = PyH('results')
    page << head(style('a:link{text-decoration: none}\na:hover{text-decoration:none;color: #FF00FF;}',type="text/css"))
    tb = page << body() << table(width="100%")
    for k in ar:
        for kk in k.get():
            t = tb << tr(width="100%")
            if '.pdf' in kk[0]:
                filename = os.path.splitext(os.path.split(kk[0])[1])[0]
                t << td(a(filename,href=kk[0],target="_blank"),width="11%",bgcolor='#F5F5DC')
                t << td(kk[1],width="8%",bgcolor='#F5F5DC')
                t << td(kk[2],width="5%",bgcolor='#F5F5DC')
                t << td(kk[3],width="5%",bgcolor='#F5F5DC')
                t << td(kk[4],width="30%",bgcolor='#F5F5DC')
                t << td(kk[5],width="30%",bgcolor='#F5F5DC')
                t << td(kk[6],width="5%",bgcolor='#F5F5DC')
            else:
                t << td(kk[0],width="11%",bgcolor='#F0E68C')
                t << td(kk[1],width="8%",bgcolor='#F0E68C')
                t << td(kk[2],width="5%",bgcolor='#F0E68C')
                t << td(kk[3],width="5%",bgcolor='#F0E68C')
                t << td(kk[4],width="30%",bgcolor='#F0E68C')
                t << td(kk[5],width="30%",bgcolor='#F0E68C')
                t << td(kk[6],width="5%",bgcolor='#F0E68C')
    path = tempfile.gettempdir()
    page.printOut(r'{}\result.html'.format(path))

def cmd_exe(event=None):
    path = tempfile.gettempdir()
    motor_model = string.upper(e.get())
    main_function(motor_model)
    Tk.quit
    ShellExecute(0,'open',r'{}\result.html'.format(path),'','',1)

def Tk_ui():
    global e,root
    root = Tk()
    width = GetSystemMetrics(0)/2
    height = GetSystemMetrics(1)/2
    root.geometry('365x59+{}+{}'.format(width-182,height-30))
    root.resizable(0,0)
    root.title(u'更改记录查询')
    Label(root,text=u'输入马达型号:',font=('Microsoft YaHei',14),width=9,bd=2,padx=8).grid(sticky=W,row=0,column=0)
    keyword = StringVar()
    e = Entry(root,textvariable=keyword,width=35,font=('Microsoft YaHei',14))
    e.grid(sticky=W,row=0,column=1)
    Label(root,text=u'(C)周斌 保留所有权利',fg='#D7236B',width='42',relief='groove',font=('Microsoft YaHei',10),bd=2,padx=14,pady=5).grid(sticky=W,row=2,column=0,columnspan=2)
    Entry.focus_set(e)
    e.bind("<Return>",cmd_exe)
    root.mainloop()

def main_function(motor_model):
    results = []
    core_num = cpu_count()          ## CPU核心数量
    p = Pool(processes=core_num)    ## 多进程
    for obj in book_list():
        result = p.apply_async(CreateLists,(obj,motor_model,))
        results.append(result)
    p.close()
    p.join()
    write_html(results)

if __name__=='__main__':
    freeze_support()
    Tk_ui()