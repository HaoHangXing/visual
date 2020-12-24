import openpyxl
from openpyxl.styles import Font, colors, Alignment
from openpyxl.utils import get_column_letter

from openpyxl.chart import ScatterChart, Series, Reference
from openpyxl.chart.layout import Layout, ManualLayout

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


class BaseSheetInfo(object):
    ws = ''
    name = ''
    index = ''

    w_line = 0 
    s_x = 1
    s_y = 1
    index = 1

    def __init__(self, ws, name, index):
        self.ws = ws
        self.name = name
        self.index = index

# 记录total数据的sheet
class WsTotal(StrTitle, BaseSheetInfo):
    sheet_index = 0

    def __init__(self, wb, name):
        ws = wb.create_sheet(name)
        BaseSheetInfo.__init__(self, ws, name, self.sheet_index)


    def InitSheetBorder(self):
        self.ws.column_dimensions['A'].width = 15
        self.ws['A1'] = self.s_title
        self.ws['B1'] = self.s_ch
        self.ws['B1'].alignment = Alignment(horizontal='center', vertical='center')
        self.w_line += 1

       
    def WriteData(self, data):
        # self.ws_dict[self.total].s_x += 1
        self.WriteOneLine(self.s_cell_vol,  data.ch,  data.cell_vol_list)
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
        self.ws.cell(y,x).alignment = Alignment(horizontal='center', vertical='center')
        x += 1
        for i, vol in enumerate(data_list):
             self.ws.cell(y, i+x, value = vol)
             self.ws.cell(y, i+x).alignment = Alignment(horizontal='center', vertical='center')
        
        self.w_line += 1


    def WriteEmptyLine(self, ws_name, num):
        self.w_line += num
        
# 记录通道数据的sheet， 因为和总表很像，就直接继承使用了
class WsChannel(WsTotal):
    def __init__(self, wb, ch_title, ch):
        WsTotal.__init__(self, wb, ch_title)
        self.ch = ch
        
    def InitSheetBorder(self):
        self.ws.column_dimensions['A'].width = 15
        self.ws['A1'] = self.s_title
        self.w_line += 1

        for i in range(1, 361):
            cell = "%s%d" % (get_column_letter(i+self.s_x), self.s_y)
            self.ws[cell] = i

    def WriteData(self, data):
        self.WriteOneLine(self.s_cell_vol,  data.cell_vol_list)
        self.WriteOneLine(self.s_test_vol,  data.test_vol_list)
        chart = LineCharts(self.ws)
        # chart.SetChartsXValue(self.s_x+1, self.s_y, self.s_x+len(data.test_vol_list))
        # chart.SetChartsYValue(self.s_x, self.w_line, self.s_x+len(data.test_vol_list))
        chart.SetChartsXValue(self.s_x+1, self.s_y, self.s_x+10)
        chart.SetChartsYValue(self.s_x, self.w_line, self.s_x+10)
        GetSheetLineData(self.ws, self.s_x, self.w_line, self.s_x+10, self.w_line)
        
        self.WriteOneLine(self.s_rad_tem,   data.rad_temperature_list)
        self.WriteOneLine(self.s_board_tem, data.board_temperature_list)
        self.WriteOneLine(self.s_current,   data.test_current_list)

        chart.SetChartsTitle()
        cell = "%s%d" % (get_column_letter(self.s_x+1), self.w_line+2)
        chart.DrawCell(cell)
        self.WriteEmptyLine(self.ws, 16)
    
    def WriteOneLine(self, title, data_list):
        pass
        y = self.s_y + self.w_line
        x = self.s_x
        
        self.ws.cell(y, x, value = title)
        x += 1

        for i, vol in enumerate(data_list):
             self.ws.cell(y, i+x, value = vol)
             self.ws.cell(y, i+x).alignment = Alignment(horizontal='center', vertical='center')
        
        self.w_line += 1

# 用于画图的类
class LineCharts(object):
    def __init__(self, ws):
        self.ws = ws
        self.ch1 = ScatterChart() # 创建图s
        
    def SetChartsXValue(self, s_x=None, s_y=None, e_x=None, e_y=None):
        self.valuesx = Reference(self.ws, s_x, s_y, e_x, e_y)

    def SetChartsYValue(self, s_x=None, s_y=None, e_x=None, e_y=None):
        values = Reference(self.ws, s_x, s_y, e_x, e_y)
        series = Series(values, self.valuesx, title_from_data=True)
        self.ch1.series.append(series)
        
        #xvalues = Reference(ws, min_col=1, min_row=2, max_row=7)
        #for i in range(2, 4):
        #    values = Reference(ws, min_col=i, min_row=1, max_row=7)
        #    series = Series(values, xvalues, title_from_data=True)
        #    ch1.series.append(series)
    
    def SetChartsTitle(self):
        self.ch1.title = "前10个开关电压"
        self.ch1.style = 10
        self.ch1.x_axis.title = 'Vol'
        self.ch1.y_axis.title = 'Percentage'
        self.ch1.legend.position = 'r' # 图例位置
    #    self.ch1.layout=Layout(
    #    manualLayout=ManualLayout(
    #        # x=0.25, y=0.25,
    #        h=0.5, w=0.5,
    #    )
    #)


    def DrawCell(self, cell):
        self.ws.add_chart(self.ch1, cell)

class ELog(object):
    def __init__(self, file):
        self.wb = openpyxl.Workbook()
        print("set write to  :"+ file)
        self.save = file

        self.ws_dict = {} # 存储通道sheet类
        
        # 创建一个总表
        self.total_ws = WsTotal(self.wb, 'total')
        # self.ws_dict[self.total_ws.name] = self.total_ws 
        self.total_ws.InitSheetBorder()
     
          
    def close(self):
        self.wb.save(self.save)
        self.wb.close()


    def InputWData(self, c_data):
        self.data = c_data


    def WriteOutLog(self):
        # 每个通道单独创建sheet存储
        # 再创建一个总sheet存储所有通道
        ws_name = f'ch{self.data.ch}'
        if ws_name in self.ws_dict.keys():
            ch_ws = self.ws_dict[ws_name]
        else:
            ch_ws = WsChannel(self.wb, ws_name, self.data.ch)
            ch_ws.InitSheetBorder()
            self.ws_dict[ch_ws.name] = ch_ws

        ch_ws.WriteData(self.data)
        self.total_ws.WriteData(self.data)
    
    def SortSheet(self):
        tmp_list = sorted(self.ws_dict.values(), key=lambda x : x.ch)
        for i, sheet in enumerate(tmp_list):
            self.wb.move_sheet_by_index(sheet.name, i)

        self.wb.move_sheet_by_index(self.total_ws.name, self.total_ws.index)


def GetSheetLineData(ws, s_x, s_y, e_x, e_y):
    cell_range = ws.iter_rows(min_col=s_x, min_row=s_y, max_col=e_x, max_row=e_y)
    tmp_list = []
    for row in cell_range:
        for cell in row:
            tmp_list.append(cell.value)
    return list
