import openpyxl

class ELog(object):
    def __init__(self, file):
        self.wb = openpyxl.Workbook()
        self.save = file
          
    def close(self):
        self.wb.save(self.save)
        self.wb.close()

    def InputWData(self, c_data):
        pass

    def WriteOutLog(self):
        pass