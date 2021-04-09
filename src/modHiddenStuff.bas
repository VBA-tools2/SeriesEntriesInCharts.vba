Attribute VB_Name = "modHiddenStuff"

'@Folder("ChartSeries")

Option Explicit
Option Private Module

'==============================================================================
'sheet name to store "hidden" worksheets, rows and columns
Public Const gcsHiddenSheetName As String = "SheetWithHiddenStuff"
'first cell where invisible sheets are stored
Public Const gcsInvisibleSheetsRange As String = "InvisibleSheets"
'first cell where hidden ranges are stored
Public Const gcsHiddenRangesRange As String = "HiddenRages"
'==============================================================================


'==============================================================================
Public Sub CollectAllHiddenStuffOnSheets( _
    ByVal wkb As Workbook _
)
    
    If DoesHiddenStuffSheetAlreadyExist(wkb) Then Exit Sub
    
    Dim wks As Worksheet
    Set wks = AddHiddenStuffSheetAndInitialize(wkb)
    
    Dim arrInvisibleSheets As Variant
    arrInvisibleSheets = CollectInvisibleSheets(wkb)
    
    If modArraySupport2.IsArrayAllocated(arrInvisibleSheets) Then
        wks.Cells(1, 1).Resize( _
                UBound(arrInvisibleSheets, 1), _
                UBound(arrInvisibleSheets, 2) _
        ) = arrInvisibleSheets
    End If
    
    Dim arrHiddenRanges As Variant
    arrHiddenRanges = CollectHiddenRanges(wkb)
    
    If modArraySupport2.IsArrayAllocated(arrHiddenRanges) Then
        wks.Cells(1, 4).Resize( _
                UBound(arrHiddenRanges, 1), _
                UBound(arrHiddenRanges, 2) _
        ) = arrHiddenRanges
    End If
    
End Sub


Private Function DoesHiddenStuffSheetAlreadyExist( _
    ByVal wkb As Workbook _
        ) As Boolean
    
    On Error Resume Next
    Dim wks As Worksheet
    Set wks = wkb.Worksheets(gcsHiddenSheetName)
    On Error GoTo 0
    
    DoesHiddenStuffSheetAlreadyExist = (Not wks Is Nothing)
    
End Function


Private Function AddHiddenStuffSheetAndInitialize( _
    ByVal wkb As Workbook _
        ) As Worksheet
    
    Dim sActiveSheetName As String
    sActiveSheetName = ActiveSheet.Name
    
    Dim wks As Worksheet
    Set wks = wkb.Worksheets.Add(, wkb.Worksheets(wkb.Worksheets.Count))
    wkb.Sheets(sActiveSheetName).Activate
    
    With wks
        .Name = gcsHiddenSheetName
        .Tab.ThemeColor = xlThemeColorLight1
        .Tab.TintAndShade = 0
        
        .Cells.NumberFormat = "@"
        
        .Names.Add Name:=gcsInvisibleSheetsRange, RefersTo:=.Cells(1, 1)
        .Names.Add Name:=gcsHiddenRangesRange, RefersTo:=.Cells(1, 4)
    End With
    
    Set AddHiddenStuffSheetAndInitialize = wks
    
End Function


Private Function CollectInvisibleSheets( _
    ByVal wkb As Workbook _
        ) As Variant
    
    Dim Arr As Variant
    ReDim Arr(1 To 2, 1 To wkb.Sheets.Count)
    
    Dim iHidden As Long
    iHidden = 0
    
    Dim i As Long
    For i = 1 To wkb.Sheets.Count
        With wkb.Sheets(i)
            If .Visible <> xlSheetVisible Then
                iHidden = iHidden + 1
                Arr(1, iHidden) = .Name
                Arr(2, iHidden) = .Visible
            End If
        End With
    Next
    
    If iHidden > 0 Then
        ReDim Preserve Arr(1 To 2, 1 To iHidden)
        
        Dim bOK As Boolean
        Dim arrTransposed() As Variant
        bOK = modArraySupport2.TransposeArray(Arr, arrTransposed)
    End If
    
    CollectInvisibleSheets = arrTransposed
    
End Function


Private Function CollectHiddenRanges( _
    ByVal wkb As Workbook _
        ) As Variant
    
    Dim Arr As Variant
    ReDim Arr(1 To 3, 1 To wkb.Worksheets.Count)
    
    Dim iHidden As Long
    iHidden = 0
    
    Dim ws As Worksheet
    For Each ws In wkb.Worksheets
        Dim sHiddenRows As String
        sHiddenRows = vbNullString
        
        Dim sHiddenColumns As String
        sHiddenColumns = vbNullString
        
        sHiddenRows = HiddenRowsInSheet(ws)
        sHiddenColumns = HiddenColumnsInSheet(ws)
        
        If Len(sHiddenRows) > 0 Or Len(sHiddenColumns) > 0 Then
            iHidden = iHidden + 1
            Arr(1, iHidden) = ws.Name
            Arr(2, iHidden) = sHiddenRows
            Arr(3, iHidden) = sHiddenColumns
        End If
    Next
    
    If iHidden > 0 Then
        ReDim Preserve Arr(1 To 3, 1 To iHidden)
        
        Dim bOK As Boolean
        Dim arrTransposed() As Variant
        bOK = modArraySupport2.TransposeArray(Arr, arrTransposed)
    End If
    
    CollectHiddenRanges = arrTransposed
    
End Function


'------------------------------------------------------------------------------
Public Sub MakeAllStuffVisibleHidden( _
    ByVal wkb As Workbook, _
    Optional ByVal bMakeHidden As Boolean = False _
)
    
    Dim wks As Worksheet
    Set wks = wkb.Worksheets(gcsHiddenSheetName)
    
    Call MakeSheetsVisibleHidden(wks, bMakeHidden)
    Call MakeRangesVisibleHidden(wks, bMakeHidden)
    
    If bMakeHidden Then
        Application.DisplayAlerts = False
        wks.Delete
        Application.DisplayAlerts = True
    End If
    
End Sub


Private Sub MakeSheetsVisibleHidden( _
    ByVal wks As Worksheet, _
    Optional ByVal bMakeHidden As Boolean = False _
)
    
    Dim rng As Range
    Set rng = wks.Range(gcsInvisibleSheetsRange)
    If rng.Value = vbNullString Then Exit Sub
    
    Dim Arr As Variant
    Arr = rng.CurrentRegion
    
    Dim wkb As Workbook
    Set wkb = wks.Parent
    
    With wkb
        If Not bMakeHidden Then
            Dim i As Long
            For i = 1 To UBound(Arr)
                .Sheets(Arr(i, 1)).Visible = xlSheetVisible
            Next
        Else
            For i = 1 To UBound(Arr)
                .Sheets(Arr(i, 1)).Visible = Arr(i, 2)
            Next
        End If
    End With
    
End Sub


Private Sub MakeRangesVisibleHidden( _
    ByVal wksHiddenStuff As Worksheet, _
    Optional ByVal bMakeHidden As Boolean = False _
)
    
    Dim rngHiddenRanges As Range
    Set rngHiddenRanges = wksHiddenStuff.Range(gcsHiddenRangesRange)
    If rngHiddenRanges.Value = vbNullString Then Exit Sub
    
    'resize to avoid error in case only rows are hidden
    Dim arrHiddenRanges As Variant
    arrHiddenRanges = rngHiddenRanges.CurrentRegion.Resize(, 3)
    
    Dim wkb As Workbook
    Set wkb = wksHiddenStuff.Parent
    
    Dim i As Long
    For i = 1 To UBound(arrHiddenRanges)
        Dim wks As Worksheet
        Set wks = wkb.Worksheets(arrHiddenRanges(i, 1))
        
        If Len(arrHiddenRanges(i, 2)) > 0 Then
            Dim Arr As Variant
            Arr = Split(arrHiddenRanges(i, 2), ",")
            
            Dim j As Long
            For j = LBound(Arr) To UBound(Arr)
                wks.Rows(Arr(j)).Hidden = bMakeHidden
            Next
        End If
        
        If Len(arrHiddenRanges(i, 3)) > 0 Then
            Arr = Split(arrHiddenRanges(i, 3), ",")
            For j = LBound(Arr) To UBound(Arr)
                wks.Columns(Arr(j)).Hidden = bMakeHidden
            Next
        End If
    Next
    
End Sub
