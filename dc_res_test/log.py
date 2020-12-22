import tool 
import openpyxl


# 配置输出的日志格式
class DataInfo(object):
    dc_start_vol = 0       # 测试前电压
    dc_end_vol = 0  # 最后一次放电电压
    e_vol = 0       # 充电完成后测得电压
    s_time_dict = {}
    e_time_dict = {}
    cell_vol_list = []
    test_vol_list = []
    rad_temperature_list = []
    board_temperature_list = []
    test_current_list = []
    ch = 0
    test_index = 0
    test_cout = 0

    
# 写log操作
class WLog(object):
    '''
    private
    '''
    # 数据格式
    title_offset = 19 # 开头格式对齐
    s_run_time = '[time]'
    s_dc_vol = '[dc vol]'
    s_ch_vol = '[ch_vol]'
    s_test_ch = '[channel]'
    s_test_index = '[index]'
    s_cell_vol = '[cell_vol]'
    s_test_vol = '[test_vol]'
    s_rad_tem = '[rad_temperature]'
    s_board_tem = '[board_temperature]'
    s_current = '[current]'

    def Writelist(self, i_list, split_char):
        for i, vol in enumerate(i_list):
            self.w_f.write(vol)
            if not i == len(i_list)-1:
                self.w_f.write(split_char)
        else:
            self.w_f.write('\n')
    '''
    public
    '''
    def __init__(self, w_file):
        self.w_f = w_file
        

    def __del__(self):
        pass

    def InputWData(self,c_data):
        self.data = c_data
    
    def WriteOutLog(self):
        # time
        str = f'%-{self.title_offset}s :%2d:%02d:%02d - %2d:%02d:%02d\n' % (
            self.s_run_time, self.data.s_time_dict['hour'], self.data.s_time_dict['min'], self.data.s_time_dict['sec']
            , self.data.e_time_dict['hour'], self.data.e_time_dict['min'], self.data.e_time_dict['sec'])
        self.w_f.write(str)

        # channel 
        str = f'%-{self.title_offset}s : %d\n' % (self.s_test_ch, self.data.ch)
        self.w_f.write(str)

        # index 
        str = f'%-{self.title_offset}s : %d\n' % (self.s_test_index, self.data.test_index)
        self.w_f.write(str)


        # start and end vol
        str = f'%-{self.title_offset}s : %.3f - %.3f  %.3f\n' % (self.s_dc_vol, self.data.dc_start_vol, self.data.dc_end_vol, self.data.e_vol)
        self.w_f.write(str)

        # cell vol
        str = f'%-{self.title_offset}s : '% (self.s_cell_vol)
        self.w_f.write(str)
        self.Writelist(self.data.cell_vol_list, '  ')

        # test vol
        str = f'%-{self.title_offset}s : '% (self.s_test_vol)
        self.w_f.write(str)
        self.Writelist(self.data.test_vol_list, '  ')

        # radiator temperature
        str = f'%-{self.title_offset}s : '% (self.s_rad_tem)
        self.w_f.write(str)
        self.Writelist(self.data.rad_temperature_list, ' ')

        # board temperature
        str = f'%-{self.title_offset}s : '% (self.s_board_tem)
        self.w_f.write(str)
        self.Writelist(self.data.board_temperature_list, ' ')

        # current
        str = f'%-{self.title_offset}s : '% (self.s_current)
        self.w_f.write(str)
        self.Writelist(self.data.test_current_list, '  ')

        self.w_f.write('\n')

# 读取的日志格式
class LogInfo(object):
    time_format = r'(\[\d+:\d+:\d+\])' # 例：[02:23:45]
    time_format_list = ('hour','min','sec') # 与time_format要对应

    s_begin_vol   = '[SUCCESS vol]'
    s_test_volt   = '[Volt]'
    s_rad_temperature = '[r_t]'
    s_board_temperature = '[p_t]'
    s_current = '[result]'
    s_end_vol = '[reach charge end vol]'

    expect_test_cout = 360 # 期望的电压总测量次数

    get_list = [s_begin_vol, s_test_volt, s_rad_temperature, s_board_temperature, s_current, s_end_vol] 

    #def __init__(self):
    #   self.get_list = [self.s_begin_vol, self.s_test_volt, self.s_rad_temperature, self.s_board_temperature, self.s_current, self.s_end_vol] 

# 控制日志信息的读写，主要是读，写操作单独一个类
class WRFile(LogInfo):
    __OneFinishFlag = False
    '''
    public
    '''
    def __init__(self, file):
        # LogInfo.__init__(self)

        self.r_file = file
        print("set read from :"+ self.r_file)
        self.r_f = open(self.r_file, 'r', encoding='utf-8')

    def SetWriteFile(self, file):
        print("set write to  :"+ file)

        # 创建要输出的文件
        self.w_f = open(file, 'w+')
        self.w_log = WLog(self.w_f)
        

    def WREnd(self):
        self.r_f.close()
        self.w_f.close()

    
    def ExtarctLogData(self):
        for line in self.r_f.readlines():
            for info in self.get_list:
                if info in line:
                    self.ExtarcOneLine(info, line) # 逐行提取数据
                    if self.__OneFinishFlag == True: # 采集一次有效数据,写一次数据
                        self.__OneFinishFlag = False
                        self.w_log.InputWData(self.data)
                        self.w_log.WriteOutLog()
    '''
    private
    '''
    # 提取一行数据
    def ExtarcOneLine(self, info, line):
        # line = line.strip() # 去掉末尾\n
        #line = tool.DeleteStr(line, self.time_format) # 删除开头的时间

        if self.s_begin_vol == info:
            # 重新初始化一个写类
            # self.data = WLog(self.w_f)
            self.data = DataInfo()

            time_list = tool.FindTime(line, self.time_format)
            self.data.s_time_dict = dict(zip(self.time_format_list, time_list))
            self.data.dc_start_vol = float((tool.FindPatternStr(line, '\d+\.\d+'))[0])

        elif self.s_test_volt == info:
            data_str =(tool.FindPatternStr(line, '\d+\,.*\,$'))[0]
            data_list = data_str.split(',')
            self.data.ch = int(data_list[0])
            self.data.test_index = int(data_list[1])
            self.data.test_cout = int(data_list[2])
            if not self.data.test_cout == self.expect_test_cout:
                return False
            self.data.cell_vol_list = data_list[3:3+200+1]
            self.data.test_vol_list = data_list[3+200+1:-1]
            self.data.dc_end_vol = float(data_list[3+200])

        elif self.s_rad_temperature == info and self.data.test_cout == self.expect_test_cout:
            data_str =(tool.FindPatternStr(line, '\d+\,.*\,$'))[0]
            data_list = data_str.split(',')
            self.data.rad_temperature_list = data_list[3:-1]

        elif self.s_board_temperature == info and self.data.test_cout == self.expect_test_cout:
            data_str =(tool.FindPatternStr(line, '\d+\,.*\,$'))[0]
            data_list = data_str.split(',')
            self.data.board_temperature_list = data_list[3:-1]

        elif self.s_current == info and self.data.test_cout == self.expect_test_cout:
            data_str =(tool.FindPatternStr(line, '\d+\,.*\,$'))[0]
            data_list = data_str.split(',')
            self.data.test_current_list = data_list[3:-1]

        elif self.s_end_vol == info and self.data.test_cout == self.expect_test_cout:
            time_list = tool.FindTime(line, self.time_format)
            self.data.e_time_dict = dict(zip(self.time_format_list, time_list))
            self.data.e_vol = float((tool.FindPatternStr(line, '\d+\.\d+'))[0])

            # 若果测试没有出现异常,应该有expect_test_cout次.
            if self.data.test_cout == self.expect_test_cout:
                self.__OneFinishFlag = True
                        

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

