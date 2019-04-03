import xlrd
import xlwt


def read_excel(file_path):
    file = file_path
    wb = xlrd.open_workbook(filename=file)# 打开文件
    # print(wb.sheet_names()) # 获取所有表格名字

    sheet1 = wb.sheet_by_index(0)#通过索引获取表格

    # rows = sheet1.row_values(1)#获取行内容
    # cols = sheet1.col_values(3)#获取列内容
    # print(rows)
    normal_data = {}
    zero_data = {}
    for i in range(1, 462):
        row = sheet1.row_values(i)
        # print(len(data))
        data_zero = {}
        data_normal = {}
        # 统计金额为0的数据
        if row[-1] <= 0:
            # print(data_zero.keys())
            print(row[1], row[-3], row[-1])
            if row[1] not in zero_data.keys():
                zero_data[row[1]] = [row[-3], row[-1]] # {云顶古树普洱陈年茶沱（尊享）:[数量， 个数]}
            else:
                zero_data[row[1]][0] += row[-3]
                zero_data[row[1]][1] += row[-1]

        # 统计正常金额的数据
        else:
            if row[1] not in normal_data.keys():
                normal_data[row[1]] = [row[-3], row[-1]] # {云顶古树普洱陈年茶沱（尊享）:[数量， 个数]}
            else:
                normal_data[row[1]][0] += row[-3]
                normal_data[row[1]][1] += row[-1]
    print(len(zero_data), zero_data)
    print(len(normal_data), normal_data)
    return zero_data, normal_data


def writer(zero_data, normal_data, out_path):
    f = xlwt.Workbook()
    sheet1 = f.add_sheet('0金额', cell_overwrite_ok=True)
    sheet2 = f.add_sheet('正常金额', cell_overwrite_ok=True)
    row0 = ["商品名称", "数量", "金额",]
    # colum0 = ["张三", "李四", "恋习Python", "小明", "小红", "无名"]
    # 写第一行

    for i in range(0, len(row0)):
        sheet1.write(0, i, row0[i])
        f.save(out_path + '销售汇总.xls')
    flag = 1
    for data in zero_data:
        # print(data, normal_data[data])

        for i in range(0, len(row0)):

            if i == 0:
                sheet1.write(flag, i, data)
            else:
                sheet1.write(flag, i, zero_data[data][i - 1])
        f.save(out_path + '销售汇总.xls')
        flag += 1
    for i in range(0, len(row0)):
        sheet2.write(0, i, row0[i])
        f.save(out_path + '销售汇总.xls')
    flag1 = 1
    for data in normal_data:
        print(data, normal_data[data])

        for i in range(0, len(row0)):

            if i == 0:
                sheet2.write(flag1, i, data)
            else:
                sheet2.write(flag1, i, normal_data[data][i-1])
        f.save(out_path + '销售汇总.xls')
        flag1 += 1
    # # 写第一列
    # for i in range(0, len(colum0)):
    #     sheet1.write(i + 1, 0, colum0[i], set_style('Times New Roman', 220, True))
    f.save(out_path + '销售汇总.xls')

def data_main(file_path, out_path):
    zero_data, normal_data = read_excel(file_path)
    writer(zero_data, normal_data, out_path)

if __name__ == '__main__':
    zero_data, normal_data = read_excel()
    writer(zero_data, normal_data)