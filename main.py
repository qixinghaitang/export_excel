import pymysql as mysql
import xlwt 
import os
import copy
from settings import *

def del_old_file(file):
    # file = 'excel_test.xls'  # 文件路径
    if os.path.exists(file):  # 如果文件存在
        # 删除文件，可使用以下两种方法。
        os.remove(file)  
        #os.unlink(path)
        print('remove old file success!')
    else:
        print('no such file: %s!'%file)  # 则返回文件不存在

# 获取每列所占用的最大列宽
def get_max_col(max_list):
    line_list = []
    # i表示行，j代表列
    for j in range(len(max_list[0])):
        line_num = []
        for i in range(len(max_list)):
            line_num.append(max_list[i][j])  # 将每列的宽度存入line_num
        line_list.append(max(line_num))  # 将每列最大宽度存入line_list
    return line_list

def style_title():
    #创建样式
    #列标题样式
    style = xlwt.XFStyle()
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 18
    style.pattern = pattern
    font = xlwt.Font()
    font.name = 'Arial'
    font.height = 20*12
    font.colour_index = 1
    font.bold = True
    font.underline = False
    font.italic = False
    style.font = font
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    style.borders = borders
    return style

def style_body():
    #正文样式
    style = xlwt.XFStyle()
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    style.borders = borders
    return style

def add_sheet(workbook, sheet_name, cursor, sql, col_title):

    worksheet = workbook.add_sheet(sheet_name)  
    title = style_title()  
    body = style_body()

    try:
        cursor.execute(sql)
        results2 = cursor.fetchall()
        row_num = 0 
        col_list = []  # 记录每行宽度
        col_num = [0 for x in range(0, len(results2[0]))]
        for i in range(0, len(results2)):
            for j in range(0, len(results2[i])):
                worksheet.write(i+1, j, results2[i][j], body)
                col_num[j] = len(str(results2[i][j]).encode('gb18030')) # 计算每列值的大小
                row_num += 1
            col_list.append(copy.copy(col_num))  # 记录一行每列写入的长度
            
        for index in range(len(col_title)):
            worksheet.write(0, index, col_title[index], title)
            col_num[index] = len(str(col_title[index]).encode('gb18030')) # 计算每列值的大小
        col_list.append(copy.copy(col_num))  # 记录一行每列写入的长度
            
        # 获取每列最大宽度
        col_max_num = get_max_col(col_list)

        # 设置自适应列宽
        for i in range(0, len(col_max_num)):
            # 256*字符数得到excel列宽,为了不显得特别紧凑添加两个字符宽度
            worksheet.col(i).width = 256 * (col_max_num[i] + 10)
        
        print('sheet: %s export success!'%sheet_name)

    except:
        print("Error: %sUnable to fetch data!"%sheet_name)

def main(file):

    del_old_file(file)

    workbook = xlwt.Workbook(encoding='utf=8')
    style = xlwt.XFStyle()

    db = mysql.connect(host = host,
                        user = user,
                        password = password,
                        port = port,
                        database = database)

    cursor = db.cursor()
    
    add_sheet(workbook, '目录', cursor, sql1, col_title1)
    add_sheet(workbook, '数据字典', cursor, sql2, col_title2)
    
    # 关闭游标
    cursor.close()

    # 关闭数据库连接
    db.close()

    workbook.save(file)
    print('export file：%s success!'%file_name)

if __name__ == "__main__":
    main(file_name)

    