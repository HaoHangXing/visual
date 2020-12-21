import tool 



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

class WLog(DataInfo):
    '''
    public
    '''
    def __init__(self, w_file):
        self.w_f = w_file
    
    def __del__(self):
        pass

    def write_self_log(self):
        pass
    
    def Writelist(self, i_list, split_char):
        for i, vol in enumerate(i_list):
            self.w_f.write(vol)
            if not i == len(i_list)-1:
                self.w_f.write(split_char)
        else:
            self.w_f.write('\n')

    def WriteOutLog(self):
        # time
        str = f'%-{self.title_offset}s :%2d:%02d:%02d - %2d:%02d:%02d\n' % (
            self.s_run_time, self.s_time_dict['hour'], self.s_time_dict['min'], self.s_time_dict['sec']
            , self.e_time_dict['hour'], self.e_time_dict['min'], self.e_time_dict['sec'])
        self.w_f.write(str)

        # channel 
        str = f'%-{self.title_offset}s : %d\n' % (self.s_test_ch, self.ch)
        self.w_f.write(str)

        # index 
        str = f'%-{self.title_offset}s : %d\n' % (self.s_test_index, self.test_index)
        self.w_f.write(str)


        # start and end vol
        str = f'%-{self.title_offset}s : %.3f - %.3f  %.3f\n' % (self.s_dc_vol, self.dc_start_vol, self.dc_end_vol, self.e_vol)
        self.w_f.write(str)

        # cell vol
        str = f'%-{self.title_offset}s : '% (self.s_cell_vol)
        self.w_f.write(str)
        self.Writelist(self.cell_vol_list, '  ')

        # test vol
        str = f'%-{self.title_offset}s : '% (self.s_test_vol)
        self.w_f.write(str)
        self.Writelist(self.test_vol_list, '  ')

        # radiator temperature
        str = f'%-{self.title_offset}s : '% (self.s_rad_tem)
        self.w_f.write(str)
        self.Writelist(self.rad_temperature_list, ' ')

        # board temperature
        str = f'%-{self.title_offset}s : '% (self.s_board_tem)
        self.w_f.write(str)
        self.Writelist(self.board_temperature_list, ' ')

        # current
        str = f'%-{self.title_offset}s : '% (self.s_current)
        self.w_f.write(str)
        self.Writelist(self.test_current_list, '  ')

        self.w_f.write('\n')

# 读取的日志格式
class LogInfo(object):
    time_format = r'(\[\d+:\d+:\d+\])' # [02:23:45]
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
    '''
    public
    '''
    def __init__(self, file):
        # LogInfo.__init__(self)

        self.r_file = file
        print("set read from :"+ self.r_file)
        self.r_f = open(self.r_file, 'r', encoding='utf-8')

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
    '''
    private
    '''
    # 提取一行数据
    def ExtarcOneLine(self, info, line):
        # line = line.strip() # 去掉末尾\n
        #line = tool.DeleteStr(line, self.time_format) # 删除开头的时间

        if self.s_begin_vol == info:
            # 重新初始化一个写类
            self.w_log = WLog(self.w_f)

            time_list = tool.FindTime(line, self.time_format)
            self.w_log.s_time_dict = dict(zip(self.time_format_list, time_list))
            self.w_log.dc_start_vol = float((tool.FindPatternStr(line, '\d+\.\d+'))[0])

        elif self.s_test_volt == info:
            data_str =(tool.FindPatternStr(line, '\d+\,.*\,$'))[0]
            data_list = data_str.split(',')
            self.w_log.ch = int(data_list[0])
            self.w_log.test_index = int(data_list[1])
            self.w_log.test_cout = int(data_list[2])
            if not self.w_log.test_cout == self.expect_test_cout:
                return False
            self.w_log.cell_vol_list = data_list[3:3+200+1]
            self.w_log.test_vol_list = data_list[3+200+1:-1]
            self.w_log.dc_end_vol = float(data_list[3+200])

        elif self.s_rad_temperature == info and self.w_log.test_cout == self.expect_test_cout:
            data_str =(tool.FindPatternStr(line, '\d+\,.*\,$'))[0]
            data_list = data_str.split(',')
            self.w_log.rad_temperature_list = data_list[3:-1]

        elif self.s_board_temperature == info and self.w_log.test_cout == self.expect_test_cout:
            data_str =(tool.FindPatternStr(line, '\d+\,.*\,$'))[0]
            data_list = data_str.split(',')
            self.w_log.board_temperature_list = data_list[3:-1]

        elif self.s_current == info and self.w_log.test_cout == self.expect_test_cout:
            data_str =(tool.FindPatternStr(line, '\d+\,.*\,$'))[0]
            data_list = data_str.split(',')
            self.w_log.test_current_list = data_list[3:-1]

        elif self.s_end_vol == info and self.w_log.test_cout == self.expect_test_cout:
            time_list = tool.FindTime(line, self.time_format)
            self.w_log.e_time_dict = dict(zip(self.time_format_list, time_list))
            self.w_log.e_vol = float((tool.FindPatternStr(line, '\d+\.\d+'))[0])

            # 若果测试没有出现异常,应该有expect_test_cout次.
            if self.w_log.test_cout == self.expect_test_cout:
                self.w_log.WriteOutLog()
            
            # del self.w_log
            

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

