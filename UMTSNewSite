Sub UMTSNewSite()
    Dim time1 As Single
    time1 = Timer
    Const MAX_NUM As Integer = 500
    Dim scell_id(MAX_NUM) As Long, tcell_id(MAX_NUM)
    Dim trnc_id(MAX_NUM) As Integer, tpsc_id(MAX_NUM) As Integer, tlac_id(MAX_NUM) As Integer, trac_id(MAX_NUM) As Integer, ecell_id(MAX_NUM) As Integer, enodeb_name(MAX_NUM) As String
    Dim ucellcount As Integer
    Dim wb_umts As Workbook, sh_umts As Worksheet, sh_lte As Worksheet
    Application.ScreenUpdating = False
    
    'Set wb_umts = Workbooks.Open("D:\UMTS New Site Script Generatorof P4565_Swap_Info.xls", ReadOnly = True)
    'Set sh_umts = wb_umts.Sheets("UMTS")
    Set sh_umts = ActiveWorkbook.Sheets("UMTS")
    Set sh_lte = ActiveWorkbook.Sheets("LTE")

    
    'Read UMTS info
    With sh_umts.Range("A:A")
        Set c = .Find("Swap ADJS")
        If Not c Is Nothing And c.Offset(1, 2).Value = "WCell_ID" And c.Offset(1, 4).Value = "ADJS_CI" _
            And c.Offset(1, 5).Value = "ADJS_LAC" And c.Offset(1, 7).Value = "ADJS_RNC_ID" _
            And c.Offset(1, 8).Value = "ADJS_SCR_Code" Then
            Set c = c.Offset(1, 0)
            ucellcount = 0
            Do
                ucellcount = ucellcount + 1
                scell_id(ucellcount) = c.Offset(ucellcount, 2).Value
                tcell_id(ucellcount) = c.Offset(ucellcount, 4).Value
                tlac_id(ucellcount) = c.Offset(ucellcount, 5).Value
                trac_id(ucellcount) = c.Offset(ucellcount, 6).Value
                trnc_id(ucellcount) = c.Offset(ucellcount, 7).Value
                tpsc_id(ucellcount) = c.Offset(ucellcount, 8).Value
            Loop While IsNumeric(c.Offset(ucellcount, 0).Value)
        Else
            MsgBox "Cannot find any keyword Swap ADJS" & "or format do not match!"
        End If
    End With
        
    'Map LTE Sites
    With sh_lte.Range("A:A")
        For i = 1 To ucellcount
            If scell_id(i) < 50000 And scell_id(i) > 30000 Then
                Set c = .Find(scell_id(i) \ 10)
                'enodeb_name(i) = Null
                If Not c Is Nothing Then
                    Dim firstaddr As String
                    firstaddr = c.Address
                    Do
                        If (c.Offset(0, 3).Value Mod 3 = (scell_id(i) Mod 10) Mod 3) Then
                            enodeb_name(i) = c.Offset(0, 15).Value
                            ecell_id(i) = c.Offset(0, 3).Value
                            Exit Do
                        Else
                            Set c = .FindNext(c)
                        End If
                    Loop While Not c Is Nothing And firstaddr <> c.Address
                End If
            End If
        Next
    End With
    
    'Create MML Command
    Open "Script for new UMTS ncell.txt" For Output As #1
    For i = 1 To ucellcount
        If enodeb_name(i) <> "" Then
            Dim mml As String
            mml = "ADD UTRANEXTERNALCELL: Mcc=" & Chr(34) & 525 & Chr(34) & ", Mnc=" & Chr(34) & "03" & Chr(34) & ", UtranCellId=xxx, UtranDlArfcn=xxx, UtranUlArfcnCfgInd=NOT_CFG, RncId=xxx, RacCfgInd=CFG, Rac=xxx, PScrambCode=xxx, Lac=xxx; {xxx}"
            mml = Replace(mml, "xxx", tcell_id(i) + 65536 * trnc_id(i), , 1)
            If tcell_id(i) Mod 10 > 3 Then
                mml = Replace(mml, "xxx", "10713", , 1)
            Else
                mml = Replace(mml, "xxx", "10737", , 1)
            End If
            mml = Replace(mml, "xxx", trnc_id(i), , 1)
            mml = Replace(mml, "xxx", trac_id(i), , 1)
            mml = Replace(mml, "xxx", tpsc_id(i), , 1)
            mml = Replace(mml, "xxx", tlac_id(i), , 1)
            mml = Replace(mml, "xxx", enodeb_name(i), , 1)
            Print #1, mml
            
            mml = "ADD UTRANNCELL: LocalCellId=xxx, Mcc=" & Chr(34) & 525 & Chr(34) & ", Mnc=" & Chr(34) & "03" & Chr(34) & ", UtranCellId=xxx; {xxx}"
            mml = Replace(mml, "xxx", ecell_id(i), , 1)
            mml = Replace(mml, "xxx", tcell_id(i) + 65536 * trnc_id(i), , 1)
            mml = Replace(mml, "xxx", enodeb_name(i), , 1)
            Print #1, mml
            
            tcell_id(i) = tcell_id(i) + 3
            mml = "ADD UTRANEXTERNALCELL: Mcc=" & Chr(34) & 525 & Chr(34) & ", Mnc=" & Chr(34) & "03" & Chr(34) & ", UtranCellId=xxx, UtranDlArfcn=xxx, UtranUlArfcnCfgInd=NOT_CFG, RncId=xxx, RacCfgInd=CFG, Rac=xxx, PScrambCode=xxx, Lac=xxx; {xxx}"
            mml = Replace(mml, "xxx", tcell_id(i) + 65536 * trnc_id(i), , 1)
            If tcell_id(i) Mod 10 > 3 Then
                mml = Replace(mml, "xxx", "10713", , 1)
            Else
                mml = Replace(mml, "xxx", "10737", , 1)
            End If
            mml = Replace(mml, "xxx", trnc_id(i), , 1)
            mml = Replace(mml, "xxx", trac_id(i), , 1)
            mml = Replace(mml, "xxx", tpsc_id(i), , 1)
            mml = Replace(mml, "xxx", tlac_id(i), , 1)
            mml = Replace(mml, "xxx", enodeb_name(i), , 1)
            Print #1, mml
            
            mml = "ADD UTRANNCELL: LocalCellId=xxx, Mcc=" & Chr(34) & 525 & Chr(34) & ", Mnc=" & Chr(34) & "03" & Chr(34) & ", UtranCellId=xxx; {xxx}"
            mml = Replace(mml, "xxx", ecell_id(i), , 1)
            mml = Replace(mml, "xxx", tcell_id(i) + 65536 * trnc_id(i), , 1)
            mml = Replace(mml, "xxx", enodeb_name(i), , 1)
            Print #1, mml
        End If
    Next
    Close #1
    
    Application.ScreenUpdating = True
    MsgBox "Total Time Used " & (Timer - time1) & "ms"
    'wb_umts.Close
End Sub
