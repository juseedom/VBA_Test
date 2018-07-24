Sub Sheet_Outlook_Body()
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Working in Excel 2000-2016
    Dim rng As Range
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strBody As String
    Dim count As Long
    

    Set rng = Nothing
    On Error Resume Next
    'Only the visible cells in the selection
    Set rng = Selection.SpecialCells(xlCellTypeVisible)
    'You can also use a fixed range if you want
    'Set rng = Sheets("YourSheet").Range("D4:D12").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If rng Is Nothing Then
        MsgBox "The selection is not a range or the sheet is protected" & _
               vbNewLine & "please correct and try again.", vbOKOnly
        Exit Sub
    End If

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    For count = 1 To 2

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    
    strBody = ActiveSheet.Cells(count, 1).Value

    On Error Resume Next
    With OutMail
        .To = "ron@debruin.nl"
        .CC = ""
        .BCC = ""
        .Subject = strBody
        '.HTMLBody = RangetoHTML(rng)
        .Body = strBody
        .Display
        '.Send   'or use .Display
    End With
    On Error GoTo 0
    
    Next count

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub
