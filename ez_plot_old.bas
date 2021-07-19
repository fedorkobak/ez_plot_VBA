Attribute VB_Name = "Module2"
Sub buil_special_plot_old_excel(range_str)
    range(range_str).Select
    
    Set Chart = ActiveSheet.Shapes.AddChart.Chart
    Chart.ChartType = xlLineMarkers
    Chart.SeriesCollection(1).ApplyDataLabels
    ' тут осторожно! может глючить
    Chart.SeriesCollection(1).DataLabels.NumberFormat = "# ##0,00"
    
    For i = Chart.SeriesCollection(1).Points.Count To Chart.SeriesCollection(1).Points.Count - 5 Step -1
        Chart.SeriesCollection(1).Points(i).MarkerSize = 10
    Next i
    
    Chart.Legend.Delete
    
End Sub


Sub ez_plot_full_old()
    j = 1
    Do While Selection.Cells(1, j) <> ""
        j = j + 1
    Loop
    
    data_end = j - 1
    
    buil_special_plot_old_excel (Selection.Cells(1, 1).Address & ":" & Selection.Cells(1, data_end).Address)
    
End Sub

Sub ez_plot_ten_old()
    buil_special_plot_old_excel (Selection.Cells(1, 1).Address & ":" & Selection.Cells(1, 10).Address)
End Sub
