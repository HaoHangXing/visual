import openpyxl
from excel import base as excel_base
from excel import chart as excel_chart 
from excel_log import total_sheet
from excel_log import channel_sheet
from excel_log import main_sheet

class SheetNameIndex(object):
    '''
    一些固定名称的sheet
    '''
    t_main = 'main'
    t_total = 'total'
    sheet_title = [t_main, t_total]


class ELog(SheetNameIndex):
    def __init__(self, file):
        self.wb = openpyxl.Workbook()
        print("set write to  :"+ file)
        self.save = file

        self.ws_dict = {} # 存储通道sheet类
        
        # 创建一个总表
        self.total_ws = total_sheet.WsTotal(self.wb, self.t_total, self.sheet_title.index(self.t_total))
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
            ch_ws = channel_sheet.WsChannel(self.wb, ws_name, self.data.ch)
            ch_ws.InitSheetBorder()
            self.ws_dict[ch_ws.name] = ch_ws

        ch_ws.WriteData(self.data)
        self.total_ws.WriteData(self.data)

    
    def WsMainOrganizeData(self):
        self.main_ws = main_sheet.WsMain(self.wb, self.t_main, self.sheet_title.index(self.t_main))
        self.main_ws.SetReadSheet(self.total_ws.ws)
        self.main_ws.Inputdata(self.data) # 为了处理部分易变的数据
        self.main_ws.MainHandler()


    def SortSheet(self):
        tmp_list = sorted(self.ws_dict.values(), key=lambda x : x.ch) # 排序通道sheet
        for i, sheet in enumerate(tmp_list):
            self.wb.move_sheet_by_index(sheet.name, i)

        for i, sheet in enumerate(self.sheet_title):
            self.wb.move_sheet_by_index(sheet, i)
        

