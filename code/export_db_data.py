# -*- coding: utf-8 -*-
import cx_Oracle
import os,sys
import shutil
import xlwt
import time
import email_main_one as send_email

def export_xls(p_date, send_hh):
    db=cx_Oracle.connect('username/password@ip/orcl') #数据库的连接信息
    cr = db.cursor()
    v_sql = "select * from table_name a where a.send_time like '%%%s%%' and a.status='T'" %(str(send_hh))
    print v_sql
    cr.execute(v_sql)
    rs = cr.fetchall()
    # p_date = time.strftime('%Y-%m-%d',time.localtime(time.time()))
    for x in rs:
        table_name = x[4].decode('gbk')
        receive_add = x[6]
        save_path = x[7]+p_date
        path = os.path.exists(save_path)
        # print save_path
        if path is False:
            os.makedirs(save_path)
        # else:
            # shutil.rmtree(save_path)
            # os.makedirs(save_path)
        v_sql = x[2].decode('gbk')
        print v_sql
        cr.execute(v_sql)
        results = cr.fetchall()
        fields = cr.description
        columnName = fields
        # 创建一个excel工作簿，编码gbk，表格中支持中文
        wb = xlwt.Workbook(encoding='gbk')
        # 创建一个sheet
        sheet = wb.add_sheet('sheet 1')
        # 获取行数
        rows = len(results)
        # 获取列数
        columns = len(columnName)

        # 创建格式style
        style = xlwt.XFStyle()
        # 创建font，设置字体
        font = xlwt.Font()
        # 字体格式
        font.name = 'Times New Roman'
        # 将字体font，应用到格式style
        style.font = font
        # 创建alignment，居中
        alignment = xlwt.Alignment()
        # 居中
        alignment.horz = xlwt.Alignment.HORZ_CENTER
        # 应用到格式style
        style.alignment = alignment

        # 创建格式style1
        style1 = xlwt.XFStyle()
        font1 = xlwt.Font()
        font1.name = 'Times New Roman'
        font1.bold = True
        style1.font = font1
        style1.alignment = alignment

        for i in range(columns):
            # 设置列的宽度
            sheet.col(i).width = 5000

        # 插入表头
        for field in range(0, len(fields)):
            # print fields[field][0].decode('gbk')
            sheet.write(0, field, fields[field][0].decode('gbk'), style1)

        # 将数据插入表格
        for i in range(1, rows+1):
            for j in range(columns):
                sheet.write(i, j, results[i-1][j], style)

        # 保存表格，并命名为****.xls
        save_path_s = save_path+'\\'
        wb.save(save_path_s+table_name+p_date+'.xls')

        # 发送邮件
        send_email.send_email(receive_add,save_path)

    cr.close()
    db.close()

if __name__ == '__main__':
    p_date_in = sys.argv[1] # 运行该脚本的第一个输入参数
    send_hh_in = sys.argv[2] # 运行该脚本的第二个输入参数
    export_xls(str(p_date_in),send_hh_in)