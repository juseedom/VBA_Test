Sub DeleteNCell()
'RMV EUTRANINTRAFREQNCELL:LOCALCELLID=1,MCC="1",MNC="1",ENODEBID=1,CELLID=1; {1007_moctestbed}
    Dim time1 As Single
    time1 = Timer
    Dim localcellid As Integer, mcc As String, mnc As String, enodebid As Long, cellid As Integer, ne As String
    Dim nncellNum As Long, cfgNum As Long
    Dim arrNNCell As Variant, arrCFG As Variant
    Dim sh_nncell As Worksheet, sh_cfg As Worksheet
    Dim flag_range As String
    Application.ScreenUpdating = False
    Set sh_nncell = Workbooks("NNCell(3Month)(NCELL).xlsx").Sheets("NNCELL")
    Set sh_cfg = Workbooks("EUTRANINTRANCELL.csv").Sheets(1)
    nncellNum = sh_nncell.UsedRange.Rows.Count
    cfgNum = sh_cfg.UsedRange.Rows.Count
    arrNNCell = sh_nncell.Range("B2:C" & nncellNum, "F2:G" & nncellNum).Value ' cell, lcll, enb, lte
    arrCFG = sh_cfg.Range("A3:B" & cfgNum, "F3:G" & cfgNum).Value
    
    For i = 1 To cfgNum - 3
        For j = 1 To nncellNum - 2
            If arrCFG(i, 1) = arrNNCell(i, 4) Then
                If arrCFG(i, 2) = arrNNCell(i, 1) And arrCFG(i, 3) = arrNNCell(i, 3) And arrCFG(i, 4) = arrNNCell(i, 2) Then
                    flag_range = flag_range & "L" & j
                End If
            If Int(Mid(arrCFG(i, 1), 1, 6)) > Int(Mid(arrNNCell(i, 1), 1, 6)) Then
                Exit For
            End If
        Next j
    Next i
    Application.ScreenUpdating = True
    MsgBox "Total Time Used " & (Timer - time1) & "s"
End Sub
