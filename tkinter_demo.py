'''
使用tkinter插件制作小工具

'''
from tkinter import *
from tkinter.filedialog import *
# from sales import data_main
import xlrd
import xlwt


class out_put(object):

    def __init__(self, file_path, out_path):
        self.file_path = file_path
        self.out_path = out_path

    def read_excel(self):
        file = self.file_path
        wb = xlrd.open_workbook(filename=file)  # 打开文件
        # print(wb.sheet_names()) # 获取所有表格名字
        sheet1 = wb.sheet_by_index(0)  # 通过索引获取表格
        print(sheet1.nrows) # 获取当前读取列表的行数
        normal_data = {}
        zero_data = {}
        # 因为第一行是表头所有从第二行开始判断
        for i in range(1, sheet1.nrows):
            row = sheet1.row_values(i)
            # print(len(data))
            data_zero = {}
            data_normal = {}
            # 统计金额为0的数据
            # print(row[1], row[-3], row[-1], type(row[-1]))
            if row[-1] <= 0:
                print(row[-1])
                # print(data_zero.keys())
                if row[1] not in zero_data.keys():
                    zero_data[row[1]] = [row[-3], row[-1]]  # {云顶古树普洱陈年茶沱（尊享）:[数量， 个数]}
                else:
                    zero_data[row[1]][0] += row[-3]
                    zero_data[row[1]][1] += row[-1]

            # 统计正常金额的数据
            else:
                if row[1] not in normal_data.keys():
                    normal_data[row[1]] = [row[-3], row[-1]]  # {云顶古树普洱陈年茶沱（尊享）:[数量， 个数]}
                else:
                    normal_data[row[1]][0] += row[-3]
                    normal_data[row[1]][1] += row[-1]
        print(len(zero_data), zero_data)
        print(len(normal_data), normal_data)
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
            f.save(self.out_path + 'sale_count.xls')
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
            f.save(self.out_path + 'sale_count.xls')
        flag1 = 1
        for data in normal_data:
            print(data, normal_data[data])

            for i in range(0, len(row0)):

                if i == 0:
                    sheet2.write(flag1, i, data)
                else:
                    sheet2.write(flag1, i, normal_data[data][i - 1])
            f.save(self.out_path + 'sale_count.xls')
            flag1 += 1
        # # 写第一列
        # for i in range(0, len(colum0)):
        #     sheet1.write(i + 1, 0, colum0[i], set_style('Times New Roman', 220, True))
        f.save(self.out_path + 'sale_count.xls')

    def data_main(self):
        zero_data, normal_data = self.read_excel()
        self.writer(zero_data, normal_data)


file_path = None
def make_app():
    tk = Tk()
    Label(tk, text='云顶财务软件').pack()
    Listbox(tk, name='l_file', bg='#F4F2F4').pack(fill=BOTH, expand=True)
    Button(tk, text='选择文件', command=select_file).pack()
    Button(tk, text='输出文件', command=gen_excel).pack()
    tk.geometry('300x300')
    return tk

def select_file():
    global file_path
    f_name = askopenfilename()
    l_box = app.children['l_file']
    if f_name:
        l_box.insert(END, f_name)

    file_path = f_name

def gen_excel():
    out_paht = askdirectory() + '/'
    l_box = app.children['l_file']
    l_box.insert(END, "输出路径：" + out_paht)
    out_put(file_path, out_paht).data_main()
    print(out_paht)


app = make_app()
app.mainloop()