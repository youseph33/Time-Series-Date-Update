# Time-Series-Date-Update
Updates the X-Axis date range in all time series charts in an excel worksheet.


Sub updatecharts()
 
Dim ws As Worksheet
Dim wb As Workbook
 
Set wb = ActiveWorkbook
Set ws = wb.Worksheets("Exhibits")
 
For i = 1 To ws.ChartObjects.Count
    If ws.ChartObjects(i).Name <> "Chart 4" Then
        ws.ChartObjects(i).Activate
        minDate = ActiveChart.Axes(xlCategory).MinimumScale
        maxdate = ActiveChart.Axes(xlCategory).MaximumScale
        newMinDate = DateAdd("m", 1, minDate)
        newMaxDate = DateAdd("m", 1, maxdate)
        ActiveChart.Axes(xlCategory).MinimumScale = newMinDate
        ActiveChart.Axes(xlCategory).MaximumScale = newMaxDate
    End If
Next i
MsgBox ("Chart Axes Updated")
 
End Sub

Sub reversecharts()
 
Dim ws As Worksheet
Dim wb As Workbook
 
Set wb = ActiveWorkbook
Set ws = wb.Worksheets("Exhibits")
 
For i = 1 To ws.ChartObjects.Count
    If ws.ChartObjects(i).Name <> "Chart 4" Then
        ws.ChartObjects(i).Activate
        minDate = ActiveChart.Axes(xlCategory).MinimumScale
        maxdate = ActiveChart.Axes(xlCategory).MaximumScale
        newMinDate = DateAdd("m", -1, minDate)
        newMaxDate = DateAdd("m", -1, maxdate)
        ActiveChart.Axes(xlCategory).MinimumScale = newMinDate
        ActiveChart.Axes(xlCategory).MaximumScale = newMaxDate
    End If
Next i
MsgBox ("Chart Axes Updated")
 
End Sub
