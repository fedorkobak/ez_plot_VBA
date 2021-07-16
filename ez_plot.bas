Attribute VB_Name = "Module1"
Sub build_special_plot(range_str)

    range(range_str).Select
    
    Set my_chart = ActiveSheet.Shapes.AddChart2(227, xlLine).Chart
    
    my_chart.SetElement (msoElementChartTitleNone)
    my_chart.SetElement (msoElementDataLabelTop)
    Set my_series = my_chart.FullSeriesCollection(1)
    
    my_series.MarkerStyle = 1
    my_series.MarkerSize = 5
    
    For i = my_series.Points.Count To my_series.Points.Count - 5 Step -1
        my_series.Points(i).MarkerSize = 10
    Next i
    
    my_series.DataLabels.NumberFormat = "# ##0,00"
    
End Sub

Sub ez_plot_full()
    j = 1
    Do While Selection.Cells(1, j) <> ""
        j = j + 1
    Loop
    
    data_end = j - 1
    
    build_special_plot (Selection.Cells(1, 1).Address & ":" & Selection.Cells(1, data_end).Address)
    
End Sub

Sub ez_plot_ten()
    build_special_plot (Selection.Cells(1, 1).Address & ":" & Selection.Cells(1, 10).Address)
End Sub
