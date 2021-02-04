from openpyxl.chart import ScatterChart, Series, Reference, LineChart
from openpyxl.chart.layout import Layout, ManualLayout

#可画点散点图、三角形散点图、曲线图
#此处画散点图
class MyScatterChart(object):
    def __init__(self, ws):
        self.ws = ws
        self.ch = LineChart() # 创建图s

    def SetChartsAddData(self, s_x=None, s_y=None, e_x=None, e_y=None):
        values = Reference(self.ws, min_col=s_x, min_row=s_y, max_col=e_x, max_row=e_y)
        self.ch.add_data(values,from_rows=True,titles_from_data=False)
        # Style the lines
        s1 = self.ch.series[0]
        #s1.marker.symbol = "triangle"
        #s1.marker.graphicalProperties.solidFill = "FF0000" # Marker filling
        #s1.marker.graphicalProperties.line.solidFill = "FF0000" # Marker outline
        s1.graphicalProperties.line.solidFill = "6495ED"
        s1.graphicalProperties.line.dashStyle = "sysDot"
        s1.graphicalProperties.line.width = 100050 # width in EMUs
        s1.graphicalProperties.line.noFill = True

    def SetChartWH(self, width, height):
        self.ch.width = width
        self.ch.height = height
    
    def SetChartsTitle(self, chart_title, x_axis_title, y_axis_title):
        self.ch.title = chart_title
        self.ch.style = 13
        self.ch.x_axis.title = x_axis_title
        self.ch.y_axis.title = y_axis_title


    def DrawCell(self, cell):
        self.ws.add_chart(self.ch, cell)


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