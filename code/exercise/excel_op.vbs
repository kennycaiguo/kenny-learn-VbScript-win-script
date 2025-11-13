On Error Resume Next
Dim excel,wb,ws

Set excel = CreateObject("Excel.Application")
excel.Visible = True
excel.Caption = "VBS调用excel"
Set wb = excel.WorkBooks.Add
Set ws = wb.Sheets(1)

'写入数据
ws.Cells(1,1).Value = "月份"
ws.Cells(1, 2).Value = "销售额"
ws.Cells(2, 1).Value = "一月"
ws.Cells(2, 2).Value = 120
ws.Cells(3, 1).Value = "二月"
ws.Cells(3, 2).Value = 150
ws.Cells(4, 1).Value = "三月"
ws.Cells(4, 2).Value = 130

'创建图表对象
Set chart = ws.Shapes.AddChart2(201, 1, 300, 250, 300, 150).Chart

'set datasource for chart
chart.SetSourceData ws.Range("a1:b4")

' 6. 格式化图表
chart.HasTitle = True
chart.ChartTitle.Text = "月度销售额趋势图"
chart.Axes(1).HasTitle = True ' X轴
chart.Axes(1).AxisTitle.Text = "月份"
chart.Axes(2).HasTitle = True ' Y轴
chart.Axes(2).AxisTitle.Text = "销售额"

' 7. 保存或导出
' wb.SaveAs "D:\sales_chart.xlsx" ' 保存工作簿
wb.SaveAs ".\sales_chart.xlsx" ' 保存工作簿,这么写是会保存到文档文件夹里面
' chart.Export "D:\sales_chart.png", "PNG" ' 导出为PNG图片,api是Export
chart.Export ".\sales_chart.png", "PNG" ' 导出为PNG图片,api是Export,这么写是会保存到文档文件夹里面

' 释放对象
Set chart = Nothing
Set ws = Nothing
Set wb = Nothing
Set excel = Nothing