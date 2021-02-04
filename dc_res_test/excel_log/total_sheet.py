from excel import base as excel_base
class StrTitle(object):
    # 数据格式
    title_offset = 19 # 开头格式对齐
    s_run_time = '[time]'
    s_dc_vol = '[dc vol]'
    s_ch_vol = '[ch_vol]'
    s_test_ch = '[channel]'
    s_test_index = '[index]'
    s_cell_vol = u'通道电压(V)'
    s_test_vol = u'开关后电压(V)'
    s_rad_tem = u'散热片温度(℃)'
    s_board_tem = u'电路板温度(℃)'
    s_current = u'电流(A)'

    s_title = u'参数'
    s_ch = u'通道'


 # 如果更改了total的数据排列或写入方式，记得更改这里
class TotalInfo(StrTitle):
    # 提供给其他sheet读取数据的起始坐标
    r_x = 2
    r_y = 2 
    line_info=[]

    def __init__(self):
        self.cell_vol_line = self.s_cell_vol
        self.test_vol_line = self.s_test_vol
        self.rad_tem_line = self.s_rad_tem
        self.board_tem_line = self.s_board_tem
        self.current_line = self.s_current
        self.line_info = [self.cell_vol_line, self.test_vol_line, self.rad_tem_line, self.board_tem_line, self.current_line]
        

# 记录total数据的sheet
class WsTotal(StrTitle, excel_base.SheetInfo, excel_base.CellAlignForamt):
    w_line = 0 # 此sheet采用一行一行写入数据

    def __init__(self, wb, name, index):
        ws = wb.create_sheet(name)
        excel_base.SheetInfo.__init__(self, ws, self.name, index)
        self.w_line = 0 


    def InitSheetBorder(self):
        self.ws.column_dimensions['A'].width = 15
        self.ws['A1'] = self.s_title
        self.ws['B1'] = self.s_ch
        self.ws['B1'].alignment = self.align_center
        self.w_line += 1

       
    def WriteData(self, data):
        self.WriteOneLine(self.s_cell_vol,  data.ch,  data.cell_vol_list+data.test_vol_list)
        self.WriteOneLine(self.s_test_vol,  data.ch,  data.test_vol_list)
        self.WriteOneLine(self.s_rad_tem,   data.ch,  data.rad_temperature_list)
        self.WriteOneLine(self.s_board_tem, data.ch,  data.board_temperature_list)
        self.WriteOneLine(self.s_current,   data.ch,  data.test_current_list)
        # self.WriteEmptyLine(self.total, 2)
  

    def WriteOneLine(self, title, ch, data_list):
        pass
        y = self.s_y + self.w_line
        x = self.s_x
        
        self.ws.cell(y, x, value = title)
        x += 1
        self.ws.cell(y, x, value = ch)
        self.ws.cell(y,x).alignment = self.align_center
        x += 1
        self.WriteRowList(x, y, data_list)
        self.SetRangeAlignment(x, y, x+len(data_list), y, self.align_center)
        self.w_line += 1


    def WriteEmptyLine(self, ws_name, num):
        self.w_line += num
