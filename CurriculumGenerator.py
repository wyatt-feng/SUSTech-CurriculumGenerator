#!/usr/bin/env python
# -*- coding: UTF-8 -*-
import openpyxl, sys, Curriculum, re, datetime

__course_start_time = {
    "1" : datetime.time(8),
    "3" : datetime.time(10, 20),
    "5" : datetime.time(14),
    "7" : datetime.time(16, 20),
    "9" : datetime.time(19)
    }
__course_end_time = {
    "2" : datetime.time(9, 50),
    "4" : datetime.time(12, 10),
    "6" : datetime.time(15, 50),
    "8" : datetime.time(18, 10),
    "10" : datetime.time(20, 50),
    "11" : datetime.time(21, 50)
    }

def usage():
    print("使用方法：CurriculumGenerator.py <xlsx课程表> <学期开始时间> <学期结束时间>")
    print("学期开始和结束时间格式：yyyymmdd。例如：20150101。")
    print("请注意：学期开始时间为学期第一周的周一，学期结束时间为学期最后一周的周末。")
    print("请自行调整调休相关课程时间安排，以及国庆周课程安排。")
    sys.exit(-1)

if len(sys.argv) != 4:
    usage()
if len(sys.argv[1]) < 6 or sys.argv[1][-5:-1] != ".xlsx":
    usage()
if len(sys.argv[2]) != 8 or len(sys.argv[3]):
    usage()
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
                    start_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day), __course_start_time.get(num[0]))
                    end_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day), __course_end_time.get(num[1]))
                elif info[0][-2] == '双':
                    type = Curriculum.CourseRepetitionType.biweekly
                    start_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day+7), __course_start_time.get(num[0]))
                    end_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day+7), __course_end_time.get(num[1]))
                else:
                    type = Curriculum.CourseRepetitionType.weekly
                    start_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day), __course_start_time.get(num[0]))
                    end_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day), __course_end_time.get(num[1]))
                location = "南方科技大学-" + info[1]
                term_start_date = datetime.datetime.combine(term_start_date, datetime.time())
                term_end_date = datetime.datetime.combine(term_end_date, datetime.time.max)
                Curriculum.add_course(curriculum, name, start_time, end_time, location, type, term_end_date)
    day += 1

curriculum.save_as_ics_file()