from excel import base as excel_base


class WsMainData():
    def __init__(self, ch):
        self.ch = ch
        self.num = 1

    def __add__(self, other):
        c_add = WsMainData()
        c_add.s_y = self.s_y + self.num + other.num
        return c_add

# Main sheet 从total表中获取想要的数据
class WsMain(excel_base.SheetInfo):
    def __init__(self, wb, name, index):
        ws = wb.create_sheet(name)
        excel_base.SheetInfo.__init__(self, ws, name, index)

    def SetReadSheet(self, ws):
        self.r_ws = ws

    def MainHandler(self):
        r_x = 2
        r_y = 2
        c_list = []
        # 提取最后5个通道电压
        cell_vol_offset = self.data.except_cell_vol_num1+self.data.except_cell_vol_num2 - 3 # '-3'是排除开头的数据
        cell_vol_num = 5-1
        # 提取160个开关电压
        test_vol_offset = 1
        test_vol_num = self.data.except_test_vol_num - 1


        # 写坐标
        self.w_x = 1
        self.w_y = 1
        # 数据前后的空行，用于写标题和结尾
        self.title_empty_line = 1 
        self.end_empty_line = 3 
        
        while True:
            self.w_y = 1
            data = self.r_ws.cell(row=r_y, column=r_x)
            ch = data.value

            if data.value == None:
                break
 
            for i, ch_data in enumerate(c_list):
                self.w_y += ch_data.num + self.title_empty_line + self.end_empty_line
                if ch_data.ch == ch:
                    ch_data.num += 1
                    self.w_y -= self.end_empty_line
                    break;
                elif ch_data.ch > ch:
                    ch_data = WsMainData(ch)
                    c_list.insert(i, ch_data)
                    self.w_y -= ch_data.num + self.title_empty_line + self.end_empty_line
                    self.AppendTitleLine(cell_vol_num, test_vol_num)
                    self.AppendEndLine(ch)
                    break
                
            else:
                ch_data = WsMainData(ch)
                c_list.append(ch_data)
                self.AppendTitleLine(cell_vol_num, test_vol_num)
                self.AppendEndLine(ch)

            cell_vol = excel_base.GetSheetLineData(self.r_ws, r_x+cell_vol_offset, r_y, r_x+cell_vol_offset+cell_vol_num, r_y) 
            test_vol = excel_base.GetSheetLineData(self.r_ws, r_x+test_vol_offset, r_y+1, r_x+test_vol_offset+test_vol_num, r_y+1)

            tmp_list = [ch, ch_data.num]
            tmp_list += cell_vol
            tmp_list.append('')
            tmp_list += test_vol

            
            self.InsertRows(self.w_y, 1)
            for i, value in enumerate(tmp_list):
                self.ws.cell(column=self.w_x+i, row=self.w_y, value =value)

            r_y += 5
            #for c in c_list:
            #    print("ch:%d num:%d"%(c.ch, c.num))
            #print("data ch:%d num:%d self.w_y:%d"%(ch_data.ch, ch_data.num, self.w_y))
            #print("====================")
        
    def AppendTitleLine(self, cell_vol_num, test_vol_num):
        tmp_list = ['通道','序号']
        s_num = self.data.except_cell_vol_num1+self.data.except_cell_vol_num2
        tmp_list += [x for x in range(s_num-cell_vol_num, s_num+1)]
        tmp_list.append('')
        tmp_list += [x for x in range(s_num+1, s_num+1+test_vol_num+1)]
        self.InsertRows(self.w_y, 1)
        self.WriteRowList(self.w_x, self.w_y, tmp_list)
        self.w_y += self.title_empty_line

    def AppendEndLine(self, ch):
        # 1楼24节220Ah2V
        #res_list = [1.246,1.111,3.31,1.913,1.882,1.754,1.243,1.487,1.318,1.428,
        #            1.797,1.619,0.771,1.638,1.024,1.198,0.797,1.985,1.415,1.651,
        #            0.864,0.91,0.984,1.045]
        #res_list = [1.54,1.127,3.43,1.915,1.871,1.778,1.23,1.489,1.323,1.362,
        #            1.768,1.666,0.789,1.648,1.038,1.204,0.78,2.002,1.378,1.631,
        #            0.884,0.928,0.988,1.386]
        #res_list = [1.27  ,1.121 ,3.42 ,1.94  ,1.87  ,1.773 ,1.224 ,1.499 ,1.349 ,1.343 ,
        #            1.774 ,1.889 ,1.804 ,1.65  ,1.044 ,1.199 ,0.78  ,1.983 ,1.385 ,1.623 ,
        #            0.899 ,0.959 ,0.988 ,1.351]

        if 0:
            # 2楼24节200Ah2V
            res_list = [0.63  ,0.634 ,0.616 ,0.629 ,0.608 ,0.625 ,0.628 ,0.609 ,0.598 ,0.646 ,
                        0.632 ,0.602 ,0.629 ,0.644 ,0.602 ,0.621 ,0.637 ,0.606 ,0.583 ,0.573 ,0.621 ,
                        0.62  ,0.603 ,0.598 ,]
        elif 0:
            # 2楼4节100Ah12V
            res_list = [5.2, 6.11, 5.78, 5.12]
        elif 1:
            res_list = [8.34,8.46,8.4 ,8.34,8.08,8.31,8.5 ,8.45,8.33,8.34,
            8.29,8.32,8.25,8.36,8.15,8.3 ,8.42,8.42,8.37,8.1 ,
            8.44,8.25,8.14,8.38,8.4 ,8.82,8.38,8.21,8.16,8.32,]
        tmp_list = ['实际内阻']
        tmp_list.append(res_list[ch])
        

        self.InsertRows(self.w_y, self.end_empty_line)
        self.WriteRowList(self.w_x, self.w_y, tmp_list)

    def Inputdata(self,data):
        self.data = data