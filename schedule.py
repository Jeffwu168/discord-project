import re
import openpyxl
from prettytable import PrettyTable
import pandas as pd

class ClassSchedule:
    def __init__(self):
        self.dic = {'星期一':'B','星期二':'C','星期三':'D','星期四':'E','星期五':'F','星期六':'G','星期日':'H'}

    def create(self):
        self.schedule = openpyxl.Workbook()
        sheet = self.schedule.worksheets[0]
        sheet.title = '課表'

        row_range = sheet['B1':'H1'][0]
        for i in range(7):
            row_range[i].value = list(self.dic.keys())[i]

        time_list = ['時間','0 7:10~8:20','1 8:10~9:00','2 9:10~10:00','3 10:20~11:10','4 11:20~12:10','5 12:20~13:10','6 13:20~14:10',
        '7 14:20~15:10','8 15:30~16:20','9 16:30~17:20','10 17:30~18:20','A 18:25~19:15','B 19:20~20:10','C 20:15~21:05','D 21:10~22:00']
        column_range = sheet['A1':'A16']
        for i in range(16):
            column_range[i][0].value = time_list[i]

        sheet['A17'].value = '沒有課程時間的課'
        sheet['A18'].value = '目前學分'
        sheet['B18'].value = 0

        self.schedule.save('class_schedule.xlsx')


    def add(self,CoursenameTeacher,hash):
        try:
            wb = openpyxl.load_workbook('class_schedule.xlsx')
            sheet = wb.worksheets[0]
        except FileNotFoundError:
            self.create()  #如果沒有課表就新增一個
            wb = openpyxl.load_workbook('class_schedule.xlsx')
            sheet = wb.worksheets[0]
        course_info = hash.search(CoursenameTeacher)  # attribute : Name , Teacher , Time , link
        name , teacher , when ,credit = course_info.Name , course_info.Teacher ,course_info.Time ,course_info.credit
        print((float(credit)))
        sheet['B18'].value = int(sheet['B18'].value) + int(float(credit))
        if when == None :  ##跑不進來這個loop
            print('Jeff is handsome')
            sheet['B17'].value = name + '(' + teacher + ')' + '\n'  # 把沒有課程時間的都放在B17
            wb.save('class_schedule.xlsx')

        else:
            when  = re.sub(r"\(.*?\)|\{.*?\}|\[.*?\]", " ", when)   # 星期二3,4 星期三6,7
            class_time = when.split(' ')       # Ex : class_time = ['星期一3,4','星期二6,7']
            for i in class_time :
                temp = re.split(r'(\d+)', i) # ['星期二', '3', ',', '4', '']
                temp = temp[:-1] # 去掉最後那個空白
                for k in temp :
                    if k == ',' :
                        pass
                    elif k in self.dic:
                        add_date = k
                    else :
                        # wb = openpyxl.load_workbook('class_schedule.xlsx')
                        # sheet = wb.worksheets[0]
                        add_row = int(k) + 2   # Ex : 第3節課要加在第5個row
                        add_column = self.dic[add_date] # Ex : 星期二要加在第C欄
                        class_name = name + '('  + teacher + ')'
                        if sheet[str(add_column) + str(add_row)].value == None:
                            sheet[str(add_column) + str(add_row)].value = class_name + '\n'
                            wb.save('class_schedule.xlsx')
                        elif class_name in sheet[str(add_column) + str(add_row)].value.split('\n')[:-1] :
                            return False
                        else:
                            sheet[str(add_column) + str(add_row)].value = sheet[str(add_column) + str(add_row)].value + class_name + '\n'
                            wb.save('class_schedule.xlsx')

    def delete(self,CoursenameTeacher,hash):
        course_info = hash.search(CoursenameTeacher)  # attribute : Name , Teacher , Time , link
        name , teacher , when , credit = course_info.Name , course_info.Teacher ,course_info.Time , course_info.credit
        wb = openpyxl.load_workbook('class_schedule.xlsx')
        sheet = wb.worksheets[0]
        sheet['B18'].value = int(sheet['B18'].value) - int(float(credit))
        if when == None :     #沒有上課時間的課程
            added_course = sheet['B17'].value.split('\n')
            added_course = added_course[:-1]
            delete_course = name + '(' + teacher + ')'
            if delete_course in added_course :
                added_course.remove(delete_course)
                sheet['B17'].value = '\n'.join(added_course) + '\n'
                wb.save('class_schedule.xlsx')

        else :
            when  = re.sub(r"\(.*?\)|\{.*?\}|\[.*?\]", " ", when)   # 星期二3,4 星期三6,7
            class_time = when.split(' ')       # Ex : class_time = ['星期一3,4','星期二6,7']
            for i in class_time :
                temp = re.split(r'(\d+)', i) # ['星期二', '3', ',', '4', '']
                temp = temp[:-1] # 去掉最後那個空白
                for k in temp :
                    if k == ',':
                        pass
                    elif k in self.dic:
                        delete_date = k
                    else:
                        delete_row = int(k) + 2   # Ex : 第3節課在第5個row
                        delete_column = self.dic[delete_date] # Ex : 星期二在第C欄
                        added_course = sheet[str(delete_column) + str(delete_row)].value.split('\n')
                        added_course = added_course[:-1] # ['資料結構(吳沛遠)','電子學(鍾孝文)']
                        delete_course = name + '(' + teacher + ')'
                        if delete_course in added_course :
                            added_course.remove(delete_course)
                            sheet[str(delete_column) + str(delete_row)].value = '\n'.join(added_course) + '\n'
                            wb.save('class_schedule.xlsx')
                        else :
                            return False

    def show(self):
        chart = pd.read_excel('class_schedule.xlsx')
        chart = chart.fillna('')
        table = PrettyTable()
        table.add_column('時間',chart['時間'].to_list())
        table.add_column('星期一',chart['星期一'].to_list())
        table.add_column('星期二',chart['星期二'].to_list())
        table.add_column('星期三',chart['星期三'].to_list())
        table.add_column('星期四',chart['星期四'].to_list())
        table.add_column('星期五',chart['星期五'].to_list())
        table.add_column('星期六',chart['星期六'].to_list())
        table.add_column('星期日',chart['星期日'].to_list())
        print(table)

        return table


