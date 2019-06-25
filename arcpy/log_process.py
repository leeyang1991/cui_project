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
    }#��־�����ϵӳ��

    def __init__(self,filename,level='info',when='D',backCount=3,fmt='%(asctime)s - %(pathname)s[line:%(lineno)d] - %(levelname)s: %(message)s'):
        self.logger = logging.getLogger(filename)
        format_str = logging.Formatter(fmt)#������־��ʽ
        self.logger.setLevel(self.level_relations.get(level))#������־����
        sh = logging.StreamHandler()#����Ļ�����
        sh.setFormatter(format_str) #������Ļ����ʾ�ĸ�ʽ
        th = handlers.TimedRotatingFileHandler(filename=filename,when=when,backupCount=backCount,encoding='utf-8')#���ļ���д��#ָ�����ʱ���Զ������ļ��Ĵ�����
        #ʵ����TimedRotatingFileHandler
        #interval��ʱ������backupCount�Ǳ����ļ��ĸ����������������������ͻ��Զ�ɾ����when�Ǽ����ʱ�䵥λ����λ�����¼��֣�
        # S ��
        # M ��
        # H Сʱ��
        # D �졢
        # W ÿ���ڣ�interval==0ʱ��������һ��
        # midnight ÿ���賿
        th.setFormatter(format_str)#�����ļ���д��ĸ�ʽ
        self.logger.addHandler(sh) #�Ѷ���ӵ�logger��
        self.logger.addHandler(th)



def process_bar(i,length,time_init=None,start_time=None,end_time=None,custom_print=''):
    '''

    :param i: ��ǰѭ�� i int
    :param length: ��ѭ������ int
    :param time_init: ��ʼʱ�� time
    :param start_time: ÿ��ѭ����ʼʱ�� time
    :param end_time: ÿ��ѭ������ʱ�� time
    :param custom_print: �Զ���print str
    :return:
    '''

    i = i + 1
    if time_init:
        time_delta = end_time - start_time
        eta = time_delta * length - time_delta*(i)
        time_elapse = end_time - time_init

        done = int(50 * (i) / length)
        sys.stdout.write(
            '\r%s '%changeTime(time_elapse)+ #��ȥʱ��
            "[%s%s]" % ('=' * done + '>'+'%0.2f' % (100 * float(i) / length) + '%', '<'+'-' * (50 - done)) + #������+�ٷֱ�
            ' eta %s'%changeTime(eta)+#ʣ��ʱ��
            '\t' +str(custom_print)
        )
        # sys.stdout.flush()
    else:
        done = int(50 * (i) / length)
        sys.stdout.write(
            "\r[%s%s]" % ('=' * done + '>'+'%0.2f' % (100 * float(i) / length) + '%', '<'+'-' * (50 - done))+  # ������+�ٷֱ�
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
