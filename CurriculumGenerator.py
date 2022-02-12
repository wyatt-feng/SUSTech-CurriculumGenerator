#!/usr/bin/env python
# -*- coding: UTF-8 -*-
import openpyxl, sys, Curriculum, re, datetime

def __course_start_time(num):
    time = {
        "1" : datetime.time(8),
        "3" : datetime.time(10, 20),
        "5" : datetime.time(14),
        "7" : datetime.time(16, 20),
        "9" : datetime.time(19)
        }
    return time.get(num)
def __course_end_time(num):
    time = {
        "2" : datetime.time(9, 50),
        "4" : datetime.time(12, 10),
        "6" : datetime.time(15, 50),
        "8" : datetime.time(18, 10),
        "10" : datetime.time(20, 50),
        "11" : datetime.time(21, 50)
        }
    return time.get(num)

workbook = openpyxl.load_workbook(sys.argv[1])
sheet = workbook.active
term_start = sys.argv[2]
term_start_date = datetime.date(int(term_start[0:4]), int(term_start[4:6]), int(term_start[6:]))
term_end = sys.argv[3]
term_end_date = datetime.date(int(term_end[0:3]), int(term_end[4:6]), int(term_start[6:]))
curriculum = Curriculum.Curriculum()
day = 0
for col in sheet.iter_cols(min_col=2, max_col=6, min_row=4, max_row=9):
    for course in col:
        if course.value:
            val = course.value.split("\n")
            for i in range(len(val)//4):
                name = val[i * 4]
                info = val[i * 4 + 3].split("][")
                num = re.findall("\d+", sheet["A%d"%course.row].value)
                if info[0][-2] == '单':
                    type = Curriculum.CourseRepetitionType.biweekly
                    start_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day), __course_start_time(num[0]))
                    end_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day), __course_end_time(num[1]))
                elif info[0][-2] == '双':
                    type = Curriculum.CourseRepetitionType.biweekly
                    start_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day+7), __course_start_time(num[0]))
                    end_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day+7), __course_end_time(num[1]))
                else:
                    type = Curriculum.CourseRepetitionType.weekly
                    start_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day), __course_start_time(num[0]))
                    end_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day), __course_end_time(num[1]))
                location = "南方科技大学-" + info[1]
                term_start_date = datetime.datetime.combine(term_start_date, datetime.time())
                term_end_date = datetime.datetime.combine(term_end_date, datetime.time.max)
                Curriculum.add_course(curriculum, name, start_time, end_time, location, type, term_end_date)
    day += 1

curriculum.save_as_ics_file()