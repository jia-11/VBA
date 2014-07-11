Sub AddSubs()
    Worksheets("Summary (3)").Activate
	'http://msdn.microsoft.com/en-us/library/office/ff838166(v=office.15).aspx
    Selection.Subtotal GroupBy:=14, Function:=xlAverage, SummaryBelowData:=False, Replace:=False, PageBreaks:=True, TotalList:=Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39)
        Worksheets("Summary (3)").Activate
    Selection.Subtotal GroupBy:=14, Function:=xlStDev, SummaryBelowData:=False, Replace:=False, PageBreaks:=True, TotalList:=Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39)
        Worksheets("Summary (3)").Activate
    Selection.Subtotal GroupBy:=14, Function:=xlMin, SummaryBelowData:=False, Replace:=False, PageBreaks:=True, TotalList:=Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39)
        Worksheets("Summary (3)").Activate
    Selection.Subtotal GroupBy:=14, Function:=xlMax, SummaryBelowData:=False, Replace:=False, PageBreaks:=True, TotalList:=Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39)
     Worksheets("Summary (3)").Activate
    Selection.Subtotal GroupBy:=14, Function:=xlCount, SummaryBelowData:=False, Replace:=False, PageBreaks:=True, TotalList:=Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39)
End Sub
Sub AddSubs_2()
    Worksheets("Summary (3)").Activate
    Selection.Subtotal GroupBy:=14, Function:=xlAverage, TotalList:=Array(3)
    Worksheets("Summary (3)").Activate
    Selection.Subtotal GroupBy:=14, Function:=xlStDev, TotalList:=Array(3)
End Sub


Sub AddSubs()
    Worksheets("2012_pie").Activate
	'http://msdn.microsoft.com/en-us/library/office/ff838166(v=office.15).aspx
    Selection.Subtotal GroupBy:=16, Function:=xlAverage, SummaryBelowData:=False, Replace:=False, PageBreaks:=True, TotalList:=Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 17)
        Worksheets("2012_pie").Activate
    Selection.Subtotal GroupBy:=16, Function:=xlStDev, SummaryBelowData:=False, Replace:=False, PageBreaks:=True, TotalList:=Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 17)
        Worksheets("2012_pie").Activate
    Selection.Subtotal GroupBy:=16, Function:=xlMin, SummaryBelowData:=False, Replace:=False, PageBreaks:=True, TotalList:=Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 17)
        Worksheets("2012_pie").Activate
    Selection.Subtotal GroupBy:=16, Function:=xlMax, SummaryBelowData:=False, Replace:=False, PageBreaks:=True, TotalList:=Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 17)
     Worksheets("2012_pie").Activate
    Selection.Subtotal GroupBy:=16, Function:=xlCount, SummaryBelowData:=False, Replace:=False, PageBreaks:=True, TotalList:=Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 17)
End Sub

Sub RemoveSubs()
    Worksheets("2012_pie").Activate
    Selection.RemoveSubtotal
End Sub


'reset the size of the selected chart
Sub ChartSize24()
'
' '
    With ActiveChart.Parent
    .Height = 400 ' resize 2.5 pt at 72 ppi.
    .Width = 800 ' resize 4.0 pt at 72 ppi.
  
    End With


End Sub
'Set the size of plot area 
Sub plot_area()
    Dim sh As Shape
    For Each sh In ActiveSheet.Shapes
         
        ActiveSheet.ChartObjects(sh.Name).Activate
        ActiveChart.PlotArea.Width = 700
        ActiveChart.PlotArea.Height = 300
    Next sh
     
End Sub

Sub cust_error()
    ActiveChart.SeriesCollection(1).Select
     ActiveChart.SeriesCollection(1).ErrorBar Direction:=xlX, _
     Include:=xlErrorBarIncludeBoth, Type:=xlErrorBarTypeCustom, _
     Amount:="='RR'!$m$2:$p$9", MinusValues:="='RR'!$m$2:$p$9"
End Sub


Sub AddSubs()
    Worksheets("2013_pie_rainfall_eff").Activate
    
    Dim cnst
    'For Each cnst In Array(xlAverage, xlStDev, xlMin, xlMax, xlCount)
    For Each cnst In Array(xlAverage, xlCount)
        Selection.Subtotal GroupBy:=18, Function:=cnst, SummaryBelowData:=False, Replace:=False, PageBreaks:=True, TotalList:=Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15)
    Next

End Sub

Sub RemoveSubs()
    Worksheets("2012&2013").Activate
    Selection.RemoveSubtotal
End Sub
