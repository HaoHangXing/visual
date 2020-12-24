

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
            self.w_f.write(str(vol))
            if not i == len(i_list)-1:
                self.w_f.write(split_char)
        else:
            self.w_f.write('\n')
    '''
    public
    '''
    def __init__(self, w_file):
        print("set write to  :"+ w_file)
        self.file = w_file
        self.w_f = open(w_file, 'w+')
        

    def __del__(self):
        pass

    def close(self):
        self.w_f.close()

    def InputWData(self,c_data):
        self.data = c_data
    
    def WriteOutLog(self):
        # time
        str = f'%-{self.title_offset}s : %2d:%02d:%02d - %2d:%02d:%02d\n' % (
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
