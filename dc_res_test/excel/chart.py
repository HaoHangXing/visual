from openpyxl.chart import ScatterChart, Series, Reference, LineChart
from openpyxl.chart.layout import Layout, ManualLayout

class DCharts(object):
    def __init__(self, ws):
        self.ws = ws
        self.ch1 = LineChart() # 创建图s
        
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

        s1 = self.ch1.series[0]
        s1.marker.symbol = "triangle"
        s1.marker.graphicalProperties.solidFill = "FF0000" # Marker filling
        s1.marker.graphicalProperties.line.solidFill = "FF0000" # Marker outline

        #s1.graphicalProperties.line.noFill = True

        #s2 = self.ch1.series[1]
        #s2.graphicalProperties.line.solidFill = "00AAAA"
        #s2.graphicalProperties.line.dashStyle = "sysDot"
        #s2.graphicalProperties.line.width = 100050 # width in EMUs

        #s2 = self.ch1.series[2]
        #s2.smooth = True # Make the line smooth

    def DrawCell(self, cell):
        self.ws.add_chart(self.ch1, cell)


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
    
    def SetChartsTitle(self, chart_title, x_axis_title, y_axis_title):
        self.ch1.title = chart_title
        self.ch1.style = 10
        self.ch1.x_axis.title = x_axis_title
        self.ch1.y_axis.title = y_axis_title
        self.ch1.legend.position = 'r' # 图例位置
    #    self.ch1.layout=Layout(
    #    manualLayout=ManualLayout(
    #        # x=0.25, y=0.25,
    #        h=0.5, w=0.5,
    #    )
    #)


    def DrawCell(self, cell):
        self.ws.add_chart(self.ch1, cell)