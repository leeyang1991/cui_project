# coding=utf-8

import logging
from logging import handlers
import sys
import math
import time
import datetime

class Logger(object):
    level_relations = {
        'debug':logging.DEBUG,
        'info':logging.INFO,
        'warning':logging.WARNING,
        'error':logging.ERROR,
        'crit':logging.CRITICAL
    }#日志级别关系映射

    def __init__(self,filename,level='info',when='D',backCount=3,fmt='%(asctime)s - %(pathname)s[line:%(lineno)d] - %(levelname)s: %(message)s'):
        self.logger = logging.getLogger(filename)
        format_str = logging.Formatter(fmt)#设置日志格式
        self.logger.setLevel(self.level_relations.get(level))#设置日志级别
        sh = logging.StreamHandler()#往屏幕上输出
        sh.setFormatter(format_str) #设置屏幕上显示的格式
        th = handlers.TimedRotatingFileHandler(filename=filename,when=when,backupCount=backCount,encoding='utf-8')#往文件里写入#指定间隔时间自动生成文件的处理器
        #实例化TimedRotatingFileHandler
        #interval是时间间隔，backupCount是备份文件的个数，如果超过这个个数，就会自动删除，when是间隔的时间单位，单位有以下几种：
        # S 秒
        # M 分
        # H 小时、
        # D 天、
        # W 每星期（interval==0时代表星期一）
        # midnight 每天凌晨
        th.setFormatter(format_str)#设置文件里写入的格式
        self.logger.addHandler(sh) #把对象加到logger里
        self.logger.addHandler(th)



def process_bar(i,length,time_init=None,start_time=None,end_time=None,custom_print=''):
    '''
    :param i: 当前循环 i int
    :param length: 总循环长度 int
    :param time_init: 初始时间 time
    :param start_time: 每个循环开始时间 time
    :param end_time: 每个循环结束时间 time
    :param custom_print: 自定义print str
    :return:
    '''

    i = i + 1
    if time_init:
        time_delta = end_time - start_time
        eta = time_delta * length - time_delta*(i)
        time_elapse = end_time - time_init

        done = int(50 * (i) / length)
        sys.stdout.write(
            '\r%s '%changeTime(time_elapse)+ #逝去时间
            "[%s%s]" % ('=' * done + '>'+'%0.2f' % (100 * float(i) / length) + '%', '<'+'.' * (50 - done)) + #进度条+百分比
            ' eta %s'%changeTime(eta)+#剩余时间
            '\t' +str(custom_print)
        )
        # sys.stdout.flush()
    else:
        done = int(50 * (i) / length)
        sys.stdout.write(
            "\r[%s%s]" % ('=' * done + '>'+'%0.2f' % (100 * float(i) / length) + '%', '<'+'-' * (50 - done))+  # 进度条+百分比
            '\t'+str(custom_print)
        )
        sys.stdout.flush()


def changeTime(allTime):
    # print(allTime)
    day = 24*60*60
    hour = 60*60
    min = 60
    if allTime <60:
        return "%ds"%math.ceil(allTime)
    elif allTime > day:
        days = divmod(allTime,day)
        return "%dd %s"%(int(days[0]),changeTime(days[1]))
    elif allTime > hour:
        hours = divmod(allTime,hour)
        return '%dh %s'%(int(hours[0]),changeTime(hours[1]))
    else:
        mins = divmod(allTime,min)
        return "%dm %ds" % (int(mins[0]), math.ceil(mins[1]))

def main():
    P = process_bar
    time_init = time.time()
    for i in range(1000):
        start = time.time()
        time.sleep(0.01)
        end = time.time()
        P(i, 1000,time_init,start,end,i)
        # start = time.time()


    # a = changeTime(100.46546)
    # print(a)

if __name__ == '__main__':
    main()
    # for i in range(10):
    #     add_time(i)
    #     time.sleep(1)