# -*- coding: utf-8 -*-
"""
 Thread：从Unit2_tag.xlsx提取点信息和运行数据，向Redis发送数据

"""
import threading
import time
from PutData2Redis import *


class PeriodicMultiTask():

    def __init__(self, delay, tasks):
        self.next_call = time.time()
        self.delay = delay
        self.tasks = tasks

    def worker(self):
        for task in self.tasks:
            task.SendToRedisHash()

        self.next_call = self.next_call + self.delay
        threading.Timer(self.next_call - time.time(), self.worker).start()

if __name__ == "__main__":

    rowindex = 2
    colidindex = 2

    TaskList = []

    colvalueindex = 6
    TaskList.append(PeriodSensorTask(u'./unit2_tag.xlsx',
                                     rowindex, colidindex, colvalueindex, u'DCS2AI',))

    cRunIntervalSeconds = 2
    MultiTask = PeriodicMultiTask(cRunIntervalSeconds, TaskList)
    MultiTask.worker()
