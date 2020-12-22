import tool 
import read_log
import write_log
import excel




# 控制日志信息的读, 在读取一次测试数据后写
class WRFile():
    '''
    public
    '''
    def __init__(self, file):
        print("set read from :"+ file)
        self.r_log = read_log.RLog(file)

    def SetWriteFile(self, file):
        # w_log
        self.w_log = write_log.WLog(file)

        # excel
        self.exl = excel.ELog(f'{file}.xlsx')

    def WREnd(self):
        self.r_log.close()

        # w_log
        self.w_log.close()

        # excel
        self.exl.close()
        
    def ExtarctLog(self):
        while self.r_log.ReadOneLine():
            data = self.r_log.GetData()
            if data:
                # log
                self.w_log.InputWData(data)
                self.w_log.WriteOutLog()
                        
                # excel
                self.exl.InputWData(data)
                self.exl.WriteOutLog()
    
                      

# 路径设置
class Log(WRFile):
    # private
    log_file = ''     # input log file 
    output_dir = 'output' # output path
    
    # public
    def __init__(self, file):
        self.log_file = file
        WRFile.__init__(self,file)
        
        tool.RemoveCreatFolder(self.output_dir)
        self.SetWriteFile(f'{self.output_dir}/out_{file}')

    
def LogHandler(file):

    if not tool.IsFile(file):
        return False
        
    log = Log(file)
    print("====== begin ExtarcData ======")
    log.ExtarctLog()
    log.WREnd()
    print('====== End ExtarcData ======')
