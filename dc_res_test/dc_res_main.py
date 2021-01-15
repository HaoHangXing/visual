# ! /usr/bin/python

#filename : dc_res_test.py

import threading
import log

class DcGetThread(threading.Thread):
    def __init__(self, threadID, name):
        threading.Thread.__init__(self)
        self.name = name
        self.threadID = threadID

    def run(self):
        print('开始处理 dc_res_test 数据')
        DcGetMainFun()
        print('结束dc_res_test')


def DcGetMainFun():

    if 1:
        while True:
            log_file = input('请输入要处理的log：')
            log.LogHandler(log_file)
        
    else:

        log.LogHandler('test.log')
        while True:
            pass

if __name__ == '__main__':
    print('hello')
    dc_res_test_thread = DcGetThread(1, "dc_res_test_thread")

    dc_res_test_thread.start()
    dc_res_test_thread.join()
    print('main end')
    
