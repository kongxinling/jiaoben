#将绿盟漏扫的主机报表中的端口信息合并到同一个表excel merge中
import xlrd
import xlwt
import os
def merge(log_d):
    merge = xlwt.Workbook()  # 新建一个excel
    sheetport = merge.add_sheet('port')  # 在新建的excel中新建一个sheet
    rows = ['ip', ' ', u'端口', u'协议', u'服务', u'状态']
    for i in range(len(rows)):
        sheetport.write(0, i, rows[i])  # 在表格第一行写入标题
    row = 1
    # 判断表中是否包含端口信息
    z = row
    zz = row
    # log_d='C:/Users/xinling/Desktop/1'#文件所在的位置
    logFiles = os.listdir(log_d)
    # 遍历文件夹内所有文件
    for filename in logFiles:
        filepath = log_d + '/' + filename
        # print(filepath)
        # 打开文件夹中的文件
        book = xlrd.open_workbook(filepath)
        sheet1 = book.sheet_by_name('主机概况')  # 读取特定表信息
        sheet2 = book.sheet_by_name('其它信息')
        # 定位端口信息所在行
        for i in range(sheet2.nrows):
            text = sheet2.row_values(i)
            while "".join(tuple(text)) == '端口信息':
                j = i
                break
        count = 1
        for i in range(sheet2.nrows):
            if count > 0:
                if i > j:
                    # 判断是否存在端口信息
                    if "".join(sheet2.cell(i + 1, 0).value) == '':
                        text4 = sheet2.row_values(i + 1)
                        col = 1
                        # 将端口信息写入表中
                        for k in text4:
                            sheetport.write(row, col, k)
                            col = col + 1
                            zz = zz + 1
                        # 如果存在端口信息则写入IP信息
                        if zz != z:
                            sheetport.write(row, 0, sheet1.cell(2, 1).value)
                            z = zz
                        row = row + 1
                        zz = row
                    else:
                        count = count - 1
                        # 只提取端口信息
                else:
                    continue
            else:
                break
    merge.save('merge.xls')
directory = input('请输入文件所在路径:')
merge(directory)



