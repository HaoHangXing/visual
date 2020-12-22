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


class RLog(LogInfo):
    def __init__(self, file):
        self.r_f = open(file, 'r', encoding='utf-8')

    def close(self):
        self.r_f.close()

    # 读取一行
    def ReadOneLine(self):
        line = self.r_f.readline()
        if not line:
            return False # 读到尾

        for info in self.get_list:
            if info in line:
                self.ExtarcOneLine(info, line) # 提取数据
                break

        return True

    def GetData(self):
        if self.__OneFinishFlag:
            return self.data
        else:
            return False
    '''
    private
    '''
    # 提取一行中数据
    def ExtarcOneLine(self, info, line):
        # line = line.strip() # 去掉末尾\n
        #line = tool.DeleteStr(line, self.time_format) # 删除开头的时间

        if self.s_begin_vol == info:
            # 重新初始化一个写类
            # self.data = WLog(self.w_f)
            self.data = DataInfo()
            self.__OneFinishFlag = False

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
  