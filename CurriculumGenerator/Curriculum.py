#!/usr/bin/env python
# -*- coding: UTF-8 -*-
import datetime, uuid
from enum import Enum
class CourseRepetitionType(Enum):
    weekly = 0
    biweekly = 1

class Course:
    def __init__(self, kwargs):
        self.event_data = kwargs
 
    def __turn_to_string__(self):
        self.event_text = "BEGIN:VEVENT\n"
        for item,data in self.event_data.items():
            item = str(item).replace("_","-")
            if item not in ["ORGANIZER","DTSTART","DTEND"]:
                self.event_text += "%s:%s\n"%(item,data)
            else:
                self.event_text += "%s;%s\n"%(item,data)
        self.event_text += "END:VEVENT\n"
        return self.event_text

class Curriculum:
    def __init__(self):
        self.__courses__ = {}
        self.__course_id__ = 0
        self.calendar_name = "课程表"

    def add_course(self, **kwargs):
        course = Course(kwargs)
        course_id = self.__course_id__
        self.__courses__[self.__course_id__] = course
        self.__course_id__ += 1
        return course_id

    def get_ics_text(self):
        self.__calendar_text__ = """BEGIN:VCALENDAR\nVERSION:2.0\nX-WR-CALNAME:课程表\n"""
        for key,value in self.__courses__.items():
            self.__calendar_text__ += value.__turn_to_string__()
        self.__calendar_text__ += "END:VCALENDAR"
        return self.__calendar_text__

    def save_as_ics_file(self):
        ics_text = self.get_ics_text()
        open("%s.ics"%self.calendar_name,"w",encoding="utf8").write(ics_text)

def add_course(curriculum, name, start_time, end_time, location, week, term_end):
    """
    向Curriculum对象添加事件的方法
    :param curriculum: curriculum实例
    :param name: 课程名
    :param start_time: 课程开始时间
    :param end_time: 课程结束时间
    :param location: 时间地点
    :param week: 上课周次（每周=0，隔周=1）
    :param term_end: 学期结束日期
    :return:
    """
    time_format = "TZID=Asia/Shanghai:{date.year}{date.month:0>2d}{date.day:0>2d}T{date.hour:0>2d}{date.minute:0>2d}{date.second:0>2d}"
    dt_start = time_format.format(date=start_time)
    dt_end = time_format.format(date=end_time)
    create_time = datetime.datetime.today().strftime("%Y%m%dT%H%M%SZ")
    if week == CourseRepetitionType.weekly:
        rrule = "FREQ=WEEKLY;UNTIL={date.year}{date.month:0>2d}{date.day:0>2d}T{date.hour:0>2d}{date.minute:0>2d}{date.second:0>2d}".format(date=term_end)
    else:
        rrule = "FREQ=WEEKLY;UNTIL={date.year}{date.month:0>2d}{date.day:0>2d}T{date.hour:0>2d}{date.minute:0>2d}{date.second:0>2d};INTERVAL=2".format(date=term_end)
    curriculum.add_course(
        SUMMARY=name,
        CREATED=create_time,
        DTSTART=dt_start,
        DTSTAMP=create_time,
        DTEND=dt_end,
        UID=str(uuid.uuid4()),
        SEQUENCE="0",
        LAST_MODIFIED=create_time,
        LOCATION=location,
        RRULE=rrule
    )
