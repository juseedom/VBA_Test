Sub CopyToPPT()
'Copy the EXCEL table to PowerPoint

Dim rng As Range
Dim PowerPointApp As Object
Dim RowsPerSlide As Integer
RowsPerSlide = 18

'Copy Range from Excel
Application.ScreenUpdating = False
Set rng = Selection
If rng.Cells.Count < 2 Then
    MsgBox "PLS select the Table firstly (including headers)"
End If
    

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
    ' MsgBox .Slides.Count
    Dim copy_rng As Range
    Dim table_slide As Object
    Dim table_shape As Object
    Dim row_count As Integer
    row_count = 2
    
    Do Until row_count >= rng.Rows.Count
        If row_count + RowsPerSlide > rng.Rows.Count Then
            Set copy_rng = rng.Rows(CStr(row_count) & ":" & CStr(rng.Rows.Count))
        Else
            Set copy_rng = rng.Rows(CStr(row_count) & ":" & CStr(row_count + RowsPerSlide - 1))
        End If
        
        
        If row_count > 2 Then
            rng.Rows(1).EntireRow.Copy
            rng.Rows(row_count).EntireRow.Insert Shift:=xlDown
            'Application.CutCopyMode = False
            Application.Wait (Now + TimeValue("0:00:01"))
            Set copy_rng = Union(rng.Rows(row_count), copy_rng)
            copy_rng.Copy
            'copy_rng.CopyPicture
        Else
            Union(rng.Rows(1), copy_rng).Copy
            'Union(rng.Rows(1), copy_rng).CopyPicture
        End If
        

        Application.Wait (Now + TimeValue("0:00:01"))
        'ppLayoutBlank   12  Blank
        'ppLayoutChart   8   Chart
        'ppLayoutChartAndText    6   Chart and text
        'ppLayoutClipartAndText  10  Clipart and text
        'ppLayoutClipArtAndVerticalText  26  ClipArt and vertical text
        'ppLayoutCustom  32  Custom
        'ppLayoutFourObjects 24  Four objects
        'ppLayoutLargeObject 15  Large object
        'ppLayoutMediaClipAndText    18  MediaClip and text
        'ppLayoutMixed   -2  Mixed
        'ppLayoutObject  16  Object
        'ppLayoutObjectAndText   14  Object and text
        'ppLayoutObjectAndTwoObjects 30  Object and two objects
        'ppLayoutObjectOverText  19  Object over text
        'ppLayoutOrgchart    7   Organization chart
        'ppLayoutTable   4   Table
        'ppLayoutText    2   Text
        'ppLayoutTextAndChart    5   Text and chart
        'ppLayoutTextAndClipart  9   Text and clipart
        'ppLayoutTextAndMediaClip    17  Text and MediaClip
        'ppLayoutTextAndObject   13  Text and object
        'ppLayoutTextAndTwoObjects   21  Text and two objects
        'ppLayoutTextOverObject  20  Text over object
        'ppLayoutTitle   1   Title
        'ppLayoutTitleOnly   11  Title only
        'ppLayoutTwoColumnText   3   Two-column text
        'ppLayoutTwoObjects  29  Two objects
        'ppLayoutTwoObjectsAndObject 31  Two objects and object
        'ppLayoutTwoObjectsAndText   22  Two objects and text
        'ppLayoutTwoObjectsOverText  23  Two objects over text
        'ppLayoutVerticalText    25  Vertical text
        'ppLayoutVerticalTitleAndText    27  Vertical title and text
        'ppLayoutVerticalTitleAndTextOverChart   28  Vertical title and text over chart
        Set table_slide = .Slides.Add(.Slides.Count + 1, 12)
        'ppPasteBitmap
        'ppPasteDefault
        'ppPasteEnhancedMetafile
        'ppPasteHTML
        'ppPasteGIF
        'ppPasteJPG
        'ppPasteMetafilePicture
        'ppPastePNG
        'ppPasteShape
        table_slide.Shapes.PasteSpecial ppPasteHTML
        'DoEvents
        Application.Wait (Now + TimeValue("0:00:01"))
        Set table_shape = table_slide.Shapes(table_slide.Shapes.Count)
        table_shape.Left = 0.5 * 72
        table_shape.Top = 0.8 * 72
        table_shape.Width = 9 * 72
        
        If row_count > 2 Then
            rng.Rows(row_count).EntireRow.Delete
        End If
        
        If row_count + RowsPerSlide > rng.Rows.Count Then
            row_count = rng.Rows.Count
        Else
            row_count = row_count + RowsPerSlide
        End If
        Application.CutCopyMode = False
    Loop
    
    'PowerPointApp.Activate
End With
'MsgBox PowerPointApp.ActivePresentation.Slices().Count()
Application.ScreenUpdating = True
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
        PowerPointApp.Presentations.Open Filename:=strPath
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

