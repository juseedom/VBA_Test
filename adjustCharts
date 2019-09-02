Sub adjustCharts()
'Adjust the size of charts each page

Dim PowerPointApp As Object

'Application.ScreenUpdating = False

'Check If PPT already Opened
Set PowerPointApp = CreateObject("PowerPoint.Application")
If PowerPointApp.Windows.Count = 0 Then
    If OpenPPT(PowerPointApp) = False Then
        Exit Sub
    End If
End If

'If Multiple PPT Opened

If PowerPointApp.Windows.Count > 1 Then
    Dim Ret_type As Integer
    Ret_type = 0
    Do Until Ret_type = 6
        Ret_type = MsgBox("Multiple PPT Windows Opened, Click OK to use:" & vbCrLf & _
                            PowerPointApp.Windows(1).Presentation.Name & vbCrLf & _
                            vbCrLf & _
                            "Click YES or close the others and then Click NO", _
                            vbYesNoCancel, _
                            "More Than 1 PPT Opened")
        '1   Specifies that OK button is clicked.
        '2   Specifies that Cancel button is clicked.
        '3   Specifies that Abort button is clicked.
        '4   Specifies that Retry button is clicked.
        '5   Specifies that Ignore button is clicked.
        '6   Specifies that Yes button is clicked.
        '7   Specifies that No button is clicked.
        If Ret_type = 2 Then
            Exit Sub
        End If
    Loop
End If


With PowerPointApp.Windows(1).Presentation
    Dim shape As Object
    Dim slide As Object
    Dim placehold As Object

    For Each slide In .Slides
        For Each placehold In slide.Shapes.Placeholders
            If placehold.Width > 8 * 72 And placehold.Height > 5 * 72 Then
                With placehold
                    .Height = 5.81 * 72
                    .Width = 8.07 * 72
                    .Left = 0.31 * 72
                    .Top = 1.27 * 72
                End With
            End If
        Next placehold
    Next slide
End With
    
'Application.ScreenUpdating = True
MsgBox "Done!"
End Sub
Private Function OpenPPT(PowerPointApp As Object) As Boolean
'Open/Create a PowerPoint to be pasted
Dim strPath As String
strPath = ""

On Error GoTo ErrHandle
    Err.Clear
    'Add File Filters
    Call Application.FileDialog(msoFileDialogOpen).Filters.Clear
    Call Application.FileDialog(msoFileDialogOpen).Filters.Add("PPT Files", "*.pptx;*.ppt;*.pptm")
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    Application.FileDialog(msoFileDialogOpen).Title = "Select the PowerPoint to be Pasted (Select Cancel to Create New)"
    'Application.FileDialog(msoFileDialogOpen).InitialFileName = "D:TempFolder to Start"
    If Application.FileDialog(msoFileDialogOpen).Show = -1 Then
        MsgBox Application.FileDialog(msoFileDialogOpen).Show
        'get the file path selected by the user
        strPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
    End If
    
    'Create an Instance of PowerPoint If Needed
    If PowerPointApp Is Nothing Then Set PowerPointApp = CreateObject("PowerPoint.Application")
    'Open the PPT File If Selected
    If strPath <> "" Then
        PowerPointApp.Presentations.Open FileName:=strPath
    Else
        PowerPointApp.Presentations.Add
    End If
    
OpenPPT = True
Exit Function

'Codes if Error happened
ErrHandle:
Const sMsg As String = "Please take a screenshot of this message and contact the developer for a resolution"
Const sTitle As String = "OpenPPT Failed"
MsgBox "The Error Happened on Line : " & Erl & vbNewLine & _
        "Error Message : " & Err.Description & vbNewLine & _
        "Error Number : " & Err.Number & vbNewLine & vbNewLine & _
        sMsg, vbCritical, sTitle
OpenPPT = False
End Function

