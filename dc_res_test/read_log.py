import tool 
# 影响数据统计的相关变量
class VarDataInfo(object):
    except_cell_vol_num1 = 180
    except_cell_vol_num2 = 80
    except_test_vol_num  = 160
    expect_test_cout = except_cell_vol_num1+except_cell_vol_num2+except_test_vol_num # 期望的电压总测量次数

# 配置输出的日志格式
class DataInfo(VarDataInfo):
    dc_start_vol = 0       # 测试前电压
    dc_end_vol = 0  # 最后一次放电电压
    e_vol = 0       # 充电完成后测得电压
    s_time_dict = {} 
    e_time_dict = {}
    cell_vol_list = []
    test_vol_list = []
    rad_temperature_list = []
    board_temperature_list = []
    test_current_list = [] # 分别是电流、电压、电阻
    ch = 0
    test_index = 0 # 第几次测试
    test_cout = 0  # 此通道第几次测试

    

# 读取的日志格式
class LogInfo(object):
    #time_format = r'(\[\d+:\d+:\d+\])' # 例：[02:23:45]
    #time_format_list = ('hour','min','sec') # 与time_format要对应

    time_format = r'(\[\d+\/\d+\/\d+ \d+:\d+:\d+\])' # 例：[2020/12/22 20:25:36]
    time_format_list = ('year', 'month', 'day', 'hour','min','sec') # 与time_format要对应

    s_begin_vol   = '[SUCCESS vol]'
    s_test_volt   = '[Volt] '
    s_rad_temperature = '[r_t] '
    s_board_temperature = '[p_t]'
    s_current = '[result]'
    s_end_vol = '[reach charge end vol]'
    s_max_charge_vol = '[reach max charge time, end vol]'
    get_list = [s_begin_vol, s_test_volt, s_rad_temperature, s_board_temperature, s_current, s_end_vol,s_max_charge_vol] 


class RLog(LogInfo, VarDataInfo):
    __OneFinishFlag = False
    __start_flag = False
    data = None
    line_num = 0
    def __init__(self, file):
        #self.r_f = open(file, 'r', encoding='utf-8')
        self.r_f = open(file, 'r')

    def close(self):
        self.r_f.close()

    # 读取一行
    def ReadOneLine(self):
        self.line_num += 1
        print("line_num:",self.line_num)
        line = self.r_f.readline()
        if not line:
            print("line end")
            return False # 读到尾

        for info in self.get_list:
            if info in line:
                self.ExtarcOneLine(info, line) # 提取数据
                break

        return True

    def GetData(self):
        if self.__OneFinishFlag:
            self.__OneFinishFlag = False
            return self.data
        else:
            return False

    def DataHandler(self):
        tool.ListConverStrToFloat(self.data.cell_vol_list)
        tool.ListConverStrToFloat(self.data.test_vol_list)
        tool.ListConverStrToFloat(self.data.rad_temperature_list)
        tool.ListConverStrToFloat(self.data.board_temperature_list)
        tool.ListConverStrToFloat(self.data.test_current_list)
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
            self.__start_flag = True

            time_list = tool.FindTime(line, self.time_format)
            self.data.s_time_dict = dict(zip(self.time_format_list, time_list))
            self.data.dc_start_vol = float((tool.FindPatternStr(line, '\d+\.\d+'))[0])

        elif self.s_test_volt == info and self.__start_flag == True:
            data_str =(tool.FindPatternStr(line, '\d+\,.*\,$'))[0]
            data_list = data_str.split(',')
            self.data.ch = int(data_list[0])
            self.data.test_index = int(data_list[1])
            self.data.test_cout = int(data_list[2])
            if not self.data.test_cout == self.expect_test_cout:
                return False
            offset = self.except_cell_vol_num1 + self.except_cell_vol_num2 #  开关电压偏移位置
            self.data.cell_vol_list = data_list[3:3+offset+1]
            self.data.test_vol_list = data_list[3+offset+1:-1]
            self.data.dc_end_vol = float(data_list[3+offset])

        elif self.s_rad_temperature == info and self.data.test_cout == self.expect_test_cout and self.__start_flag == True and self.data:
            data_str =(tool.FindPatternStr(line, '\d+\,.*\,$'))[0]
            data_list = data_str.split(',')
            self.data.rad_temperature_list = data_list[3:-1]

        elif self.s_board_temperature == info and self.data.test_cout == self.expect_test_cout and self.__start_flag == True and self.data:
            data_str =(tool.FindPatternStr(line, '\d+\,.*\,$'))[0]
            data_list = data_str.split(',')
            self.data.board_temperature_list = data_list[3:-1]

        elif self.s_current == info and self.data.test_cout == self.expect_test_cout and self.__start_flag == True and self.data:
            data_str =(tool.FindPatternStr(line, '\d+\,.*\,$'))[0]
            data_list = data_str.split(',')
            self.data.test_current_list = data_list[3:-1]

        #elif (self.s_end_vol == info or self.s_max_charge_vol == info )and self.data.test_cout == self.expect_test_cout and self.__start_flag == True  and self.data:
            time_list = tool.FindTime(line, self.time_format)
            self.data.e_time_dict = dict(zip(self.time_format_list, time_list))
            self.data.e_vol = float((tool.FindPatternStr(line, '\d+\.\d+'))[0])
            self.__start_flag == False
            # 若果测试没有出现异常,应该有expect_test_cout次.
            if self.data.test_cout == self.expect_test_cout:
                self.DataHandler()
                self.__OneFinishFlag = True


            
  