'''
Created on Jan 15, 2019

@author: EnjeruTantei
'''

class ExcelHandler():
    '''
    classdocs
    '''
    import datetime

    def __init__(self, name="Candle_Wic_Order_"+datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')+".xls", sheetName="CandleWicOrder"):
        '''
        Constructor
        '''
        import xlwt
        import datetime
        self.wb = xlwt.Workbook()

        self.sheet1 = self.wb.add_sheet(sheetName)
        self.name = name

    def savebook(self):
        import os
        dir_path = os.path.dirname(os.path.realpath(__file__))
        self.wb.save(os.path.join(dir_path,self.name))

    def write_row(self, col, row, val, sheet=None):
        if (sheet == None):
            sheet = self.sheet1
        sheet.write(col, row, val)
