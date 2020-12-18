import tool 


class BaseLog():
    # private
    ch_index = 0 # 通道号
    time = 0
    cout = 0

class ChLog():
    times = 0 # 总测量次数

# 配置日志格式
class LogInfo(object):
    time_format = r'(\[\d+:\d+:\d+\])'

    
    begin_vol   = '[SUCCESS vol]'
    test_volt   = '[Volt]'
    rad_temperture = '[r_t]'
    board_temperture = '[p_t]'
    current = '[result]'
    end_vol = '[reach charge end vol]'

    start_time = '[start_time]'
    end_time = '[end_time]'
    test_ch = '[ch]'
    test_index = '[index]'
    test_cout = 'num'

    get_list = [begin_vol, test_volt, rad_temperture, board_temperture, current, end_vol] 

    #def __init__(self):
    #   self.get_list = [self.begin_vol, self.test_volt, self.rad_temperture, self.board_temperture, self.current, self.end_vol] 

# 控制日志信息的读写
class WRFile(LogInfo):
    # public
    def __init__(self, file):
        #LogInfo.__init__(self)

        self.r_file = file
        print("set read from :"+ self.r_file)
        self.r_f = open(self.r_file, 'r')

    def SetWriteFile(self, file):
        self.w_file = file
        print("set write to  :"+ self.w_file)
        self.w_f = open(self.w_file, 'w+')

    def WREnd(self):
        self.r_f.close()
        self.w_f.close()

    
    def ExtarctLogData(self):
        for line in self.r_f.readlines():
            for info in self.get_list:
                if info in line:
                    self.ExtarcOneLine(info, line)
    # private
    def ExtarcOneLine(self, info, line):
        # line = line.strip()
        
        if self.begin_vol == info:
            time_list = tool.FindTime(line, self.time_format)

           
            print(time_list)
        elif self.end_vol == info:
            time_list = tool.FindTime(line, self.time_format)
            print(time_list)

        line = tool.DeleteStr(line, self.time_format) # 删除开头的时间
        self.w_f.write(line)


# 全局控制文件
class Log(WRFile):
    # private
    log_file = ''     # input log file 
    output_dir = 'output' # output path
    

    def __init__(self, file):
        self.log_file = file
        WRFile.__init__(self,file)
        
        tool.RemoveCreatFolder(self.output_dir)
        self.SetWriteFile(f'{self.output_dir}/out_{file}')

    # public
    


    
def LogHandler(file):
    
    if not tool.IsExist(file):
        return False
        
    log = Log(file)
    print("begin ExtarcData")
    log.ExtarctLogData()
    log.WREnd()

