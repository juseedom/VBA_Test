Sub Change_Excel_Chart_Data()
    Dim nChart As Long
    Dim nSeries As Long
    Dim nData As String
    nChart = ActiveSheet.ChartObjects.Count
    For i = 1 To nChart
        With ActiveSheet.ChartObjects(i)
            nSeries = .Chart.SeriesCollection.Count
            For j = 1 To nSeries
                sData = .Chart.SeriesCollection(j).Formula
                sData = Replace(sData, "CD", "O")
                .Chart.SeriesCollection(j).Formula = sData
            Next
        End With
    Next
End Sub
