import xlrd
import os


class ExcelReade(object):
    def __init__(self, excel_name, sheet_name):
        """
        # 我把excel放在工程包的data文件夹中：
        # 1.需要先获取到工程文件的地址
        # 2.再找到excel的文件地址（比写死的绝对路径灵活）

            os.path.relpath(__file__)
            1.根据系统获取绝对路径
            2.会根据电脑系统自动匹配路径：mac路径用/,windows路径用\
            3.直接使用__file__方法是不会自动适配环境的
        """
        # 获取当前.py文件所在文件夹层
        dir_path = os.path.dirname(os.path.realpath(__file__))
        # 再次向上获取工程的文件夹层
        project_path = os.path.dirname(dir_path)
        # 获取excel所在文件目录
        self.excel_path = os.path.join(project_path, "data", excel_name)
        # 打开指定的excel文件
        self.date = xlrd.open_workbook(self.excel_path)
        # 找到指定的sheet页
        self.table = self.date.sheet_by_name(sheet_name)
        self.rows = self.table.nrows  # 获取总行数
        self.cols = self.table.ncols  # 获取总列数

    def data_dict(self):
        if self.rows <= 1:
            print("总行数小于等于1，路径：", end='')
            print(self.excel_path)
            return False
        else:
            # 将列表的第一行设置为字典的key值
            keys = self.table.row_values(0)
            # 定义一个数组
            data = []
            # 从第二行开始读取数据，循环rows（总行数）-1次
            for i in range(1, self.rows):
                # 循环内定义一个字典，每次循环都会清空
                dict = {}
                # 从第一列开始读取数据，循环cols（总列数）次
                for j in range(0, self.cols):
                    # 将value值关联同一列的key值
                    dict[keys[j]] = self.table.row_values(i)[j]
                # 将关联后的字典放到数组里
                data.append(dict)
            return data

if __name__ == '__main__':
    start = ExcelReade('tips_library.xlsx', u'Sheet1')
    data = start.data_dict()
    date = xlrd.open_workbook('tips_library.xlsx')
    table = date.sheet_by_name('Sheet1')
    rows = table.nrows
    cols = table.ncols
    for i in range(0, rows):
        str = ""
        for j in range(0, cols):
            str += table.cell(i, j).value
        print(str)
    #value = table.cell(1, 2).value  # 获取第二行第三列单元格的值
    #print(value)
    #for i in range(len(data)):
     #   print(data[i])
      #  break

