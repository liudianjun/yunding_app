'''
使用tkinter插件制作小工具

'''
from tkinter import *
from tkinter.filedialog import *
# from sales import data_main
import xlrd
import xlwt
import re


class out_put(object):

    def __init__(self, file_path, out_path, file_name):
        self.file_path = file_path
        self.out_path = out_path
        self.file_name = file_name

    def read_excel(self):
        file = self.file_path
        wb = xlrd.open_workbook(filename=file)  # 打开文件
        # print(wb.sheet_names()) # 获取所有表格名字
        sheet1 = wb.sheet_by_index(0)  # 通过索引获取表格
        # print(sheet1.nrows) # 获取当前读取列表的行数
        # 获取 商品名称 数量 金额对应的索引
        goods_index = None
        numb_index = None
        amount_index = None
        # 获取表字段
        fields = sheet1.row_values(0)
        print(fields)
        nrows = sheet1.nrows
        ncols = sheet1.ncols
        # print('表头：', fields, nrows, ncols)
        for j,i in enumerate(fields):
            if i == '商品名称' or i == '品名':
                goods_index = j
            if i == '数量':
                numb_index = j
            if i == '金额':
                amount_index = j

        normal_data = {}
        zero_data = {}
        print(goods_index, numb_index, amount_index)
        # 因为第一行是表头所有从第二行开始判断
        for i in range(1, sheet1.nrows):
            amount = ''
            count = ''
            row = sheet1.row_values(i)
            # print(row)
            data_zero = {}
            data_normal = {}
            # 统计金额为0的数据
            # print(row[goods_index], row[numb_index], row[amount_index], type(row[amount_index]))
            # print('amount_index', amount_index, row[numb_index], type(row[numb_index]))
            # 获取金额列数据，有些数据是空字符串，改成0
            if isinstance(row[amount_index], str):
                amount = 0
            else:
                amount = row[amount_index]
            # 获取数量列数据，有些数据是空字符串，改成0
            if isinstance(row[numb_index], str):
                if row[numb_index] == '':
                    count = 0
                else:
                    count = float(re.findall(r'\d+', row[numb_index])[0])
            else:
                count = row[numb_index]
            print('count->', count)
            if amount == 0:
                # print(row[amount_index])
                # print(data_zero.keys())
                if row[goods_index] not in zero_data.keys():
                    zero_data[row[goods_index]] = [count, amount]  # {云顶古树普洱陈年茶沱（尊享）:[数量， 个数]}
                else:
                    zero_data[row[goods_index]][0] += count
                    zero_data[row[goods_index]][1] += amount

            # 统计正常金额的数据
            else:
                if row[goods_index] not in normal_data.keys():
                    normal_data[row[goods_index]] = [count, amount]  # {云顶古树普洱陈年茶沱（尊享）:[数量， 个数]}
                else:
                    normal_data[row[goods_index]][0] += count
                    normal_data[row[goods_index]][1] += amount
        # print(len(zero_data), zero_data)
        # print(len(normal_data), normal_data)
        return zero_data, normal_data

    def writer(self,zero_data, normal_data):
        f = xlwt.Workbook()
        sheet1 = f.add_sheet('0金额', cell_overwrite_ok=True)
        sheet2 = f.add_sheet('正常金额', cell_overwrite_ok=True)
        row0 = ["商品名称", "数量", "金额", ]
        # colum0 = ["张三", "李四", "恋习Python", "小明", "小红", "无名"]
        # 写第一行

        for i in range(0, len(row0)):
            sheet1.write(0, i, row0[i])
            f.save(self.out_path + '统计' + self.file_name)
        flag = 1
        for data in zero_data:
            # print(data, normal_data[data])

            for i in range(0, len(row0)):

                if i == 0:
                    sheet1.write(flag, i, data)
                else:
                    sheet1.write(flag, i, zero_data[data][i - 1])
            f.save(self.out_path + 'sale_count.xls')
            flag += 1
        for i in range(0, len(row0)):
            sheet2.write(0, i, row0[i])
            f.save(self.out_path + '统计' + self.file_name)
        flag1 = 1
        for data in normal_data:
            # print(data, normal_data[data])

            for i in range(0, len(row0)):

                if i == 0:
                    sheet2.write(flag1, i, data)
                else:
                    sheet2.write(flag1, i, normal_data[data][i - 1])
            f.save(self.out_path + '统计' + self.file_name)
            flag1 += 1
        # # 写第一列
        # for i in range(0, len(colum0)):
        #     sheet1.write(i + 1, 0, colum0[i], set_style('Times New Roman', 220, True))
        f.save(self.out_path + '统计' + self.file_name)

    def data_main(self):
        zero_data, normal_data = self.read_excel()
        self.writer(zero_data, normal_data)


file_path = None
file_name = None
file_name = None
def make_app():
    tk = Tk()
    Label(tk, text='云顶财务软件').pack()
    Listbox(tk, name='l_file', bg='#F4F2F4').pack(fill=BOTH, expand=True)
    Button(tk, text='选择文件', command=select_file).pack()
    Button(tk, text='输出文件', command=gen_excel).pack()
    tk.geometry('300x300')
    return tk

def select_file():
    '''
    选择文件
    :return:
    '''
    global file_path
    global file_name
    f_name = askopenfilename()
    l_box = app.children['l_file']
    if f_name:
        l_box.insert(END, f_name)

    file_path = f_name
    file_name = file_path.split('/')[-1]
    print('文件名：', file_name)

def gen_excel():
    out_paht = askdirectory() + '/'
    l_box = app.children['l_file']
    l_box.insert(END, "输出路径：" + out_paht)
    out_put(file_path, out_paht,file_name).data_main()
    print(out_paht)


app = make_app()
app.mainloop()