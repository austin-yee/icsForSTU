# @Author  : Austin_Yee

import xlrd
import re
from icalendar import Calendar, Event
from datetime import datetime, timedelta
from uuid import uuid1

s=input(r'请输入你的源数据的绝对路径，例如C:\Users\17106\Desktop\源.xls')
q=int(input('请输入你的课程数量，例：10'))
m=int(input('请输入第一周的月份，例：2'))
d=int(input('请输入第一周周一的日子，例：21'))
w=int(input('请输入你想要的格式\n格式0 [开课班号]课程名字 请输入0\n格式1 课程名字         请输入1\n格式2 [课程代码]课程名字 请输入2'))

# 新建日历，参数为日历名字
def create_calendar(name):
    my_calender = Calendar()
    # 添加属性
    my_calender.add('X-WR-CALNAME', name)
    my_calender.add('version', '2.0')
    my_calender.add('METHOD', 'PUBLISH')
    my_calender.add('CALSCALE', 'GREGORIAN')  # 历法：公历
    my_calender.add('X-WR-TIMEZONE', 'Asia/Shanghai')  # 通用扩展属性，表示时区
    return my_calender

# 更换课程名称前缀，0为课程代码换为开课班号，1为去掉课程代码,2为课程代码不变
def change_prefix_of_course_name(pattern,course_name,class_number):
    if pattern == 0:
        course_name = re.sub(r'\[.*?\]', '[' + class_number + ']', course_name)
    if pattern == 1:
        course_name = re.sub(r'\[.*?\]', '', course_name)
    if pattern == 2:
        course_name =course_name
    return course_name

# 新建一节课
def add_class(course_name, teacher, start_time, end_time, location):
    event = Event()
    event.add('uid', str(uuid1()))
    event.add('summary', course_name)
    event.add('dtstart', start_time)
    event.add('dtend', end_time)
    event.add('location', location)
    event.add('description', '授课老师：' + teacher)
    return event

# 更改上课序号为小时
def change_hour():
    dict_start_time1_hour = {'12': int(8), '34': int(10), '345': int(10), '67': int(14), '6789': int(14), '89': int(16), '890': int(16), 'AB': int(19), 'ABC': int(19)}
    return dict_start_time1_hour

# 更改上课序号为分钟
def change_minute():
    dict_start_time1_minute = {'12': int(00), '34': int(00), '345': int(00), '67': int(00), '6789': int(00),'89': int(00), '890': int(00), 'AB': int(20), 'ABC': int(20)}
    return dict_start_time1_minute

# 每节间隔的时间
def get_each_time():
    dict_each_time = {int(2): int(100), int(3): int(155), int(4): int(210)}
    return dict_each_time

# 参数为第一周的第一个周一
def get_data(first_date):

    dict_date = {int(1): int(first_date), int(2): int(first_date+1), int(3): int(first_date+2), int(4): int(first_date+3), int(5): int(first_date+4), int(6): int(first_date+5),int(7): int(first_date+6)}
    return dict_date

# 获取开课周第一周
def get_start_week(week_range):
    start_week=int(week_range[0]+week_range[1])
    return start_week

# 获取开课周最后一周
def get_end_week(week_range):
    end_week=int(week_range[-2]+week_range[-1])
    return end_week

# 实现方法
def main(source,q):

    book = xlrd.open_workbook(source)
    s1 = book.sheet_by_index(0)
    source=re.sub(r'\源.xls','',source)
    source=source+'CLassTable.ics'
    my_calender=create_calendar('课表')

    # 循环一次为一节课
    for i in range(0,q):
        # 开课班号
        class_number = str(s1.cell(i + 1, 0).value)
        # 课程名称
        course_name = str(s1.cell(i + 1, 1).value)
        # 班级
        location = str(s1.cell(i + 1, 4).value)
        # 教师
        teacher = str(s1.cell(i + 1, 3).value)
        # 上课时间序号
        which_class = s1.row_values(i + 1, 6, 14)  # 开区间
        # 删除开课班号
        course_name = change_prefix_of_course_name(w,course_name,class_number)
        # 开课周
        week_range = str(s1.cell(i + 1, 5).value)
        # 使用字典将上课时间由序号转换为小时和分钟
        dict_start_time1_hour = change_hour()
        dict_start_time1_minute = change_minute()
        # 每节间隔的时间
        dict_each_time =get_each_time()
        # 计算日期
        dict_date = get_data(d)
        # 获取开课周的第一周和最后一周
        start_week = get_start_week(week_range)
        end_week = get_end_week(week_range)

        # j为周几，j=0为周日，j>0为周一到周六
        for j in range(0,7):
            if which_class[j] != '':
                # date1=str(s1.cell(0,j + 6).value)
                which_class1 = which_class[j]
                start_time1 = datetime(2022, m, dict_date[j], dict_start_time1_hour[which_class1], dict_start_time1_minute[which_class1])

                for a in range(start_week-1, end_week):
                    start_time = start_time1 + timedelta(days=int(7 * a))
                    end_time = start_time + timedelta(minutes=dict_each_time[len(which_class1)])

                    ClassEvent = add_class(course_name, teacher, start_time, end_time, location)
                    my_calender.add_component(ClassEvent)

        with open(source, "wb") as fo:
            fo.write(my_calender.to_ical().replace(b'\r\n', b'\n').strip())

# s=r'C:\Users\17106\Desktop\源.xls'
# q=int(10)
main(s,q)