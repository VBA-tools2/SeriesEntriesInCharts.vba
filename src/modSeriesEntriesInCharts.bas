Attribute VB_Name = "modSeriesEntriesInCharts"

Option Explicit
Option Base 1


Public Enum eSD
    [_First] = 1
    ChartNumber = eSD.[_First]
    SheetName
    ChartName
    ChartTitle
    XLabel
    YLabel
    Y2Label
    SeriesName
    SeriesDataSheet
    SeriesXValues
    SeriesYValues
    AxisGroup
    PlotOrder
    PlotOrderTotal
    XYDataSheetEqual
    [_Last] = eSD.XYDataSheetEqual
End Enum

'==============================================================================
'sheet name to store "hidden" worksheets, rows and columns
Private Const gcsHiddenSheetName As String = "SheetWithHiddenStuff"
'first cell where invisible sheets are stored
Private Const gcsInvisibleSheetsRange As String = "InvisibleSheets"
'first cell where hidden ranges are stored
Private Const gcsHiddenRangesRange As String = "HiddenRages"
'------------------------------------------------------------------------------
'sheet name of pasted Series data?
Public Const gcsLegendSheetName As String = "SeriesEntriesInCharts"
'title row on 'gcsLegendSheetName'
Public Const gciTitleRow As Long = 2
'==============================================================================


Public Sub ListAllSCEntriesInAllCharts()

    Dim wkb As Workbook
    Dim wksSeriesLegend As Worksheet
    Dim arrData As Variant
    Dim bAreSCsFound As Boolean
    Dim bNewSeriesSheet As Boolean
    
    
    With Application
        .ScreenUpdating = False
        .StatusBar = "Running 'ListAllSCEntriesInAllCharts' ..."
    End With
    
    Set wkb = ActiveWorkbook
    
    Call CollectAllHiddenStuffOnSheets(wkb)
    Call MakeAllStuffVisibleHidden(wkb, False)
    
    bAreSCsFound = CollectSCData(wkb, arrData)
    
    'store if the sheet was newly added/created
    bNewSeriesSheet = WasSeriesEntriesInChartsWorksheetCreatedAndInitialized(wkb)
    Set wksSeriesLegend = wkb.Worksheets(gcsLegendSheetName)
    
    Call PasteDataToCollectionSheet(wksSeriesLegend, arrData)
    
    If bAreSCsFound Then
        Call MarkEachOddChartNumber(wksSeriesLegend)
        
        Call MarkSheetNameOrSeriesDataSheetIfSourceIsInvisible(wkb, arrData, False)
        Call MarkSheetNameOrSeriesDataSheetIfSourceIsInvisible(wkb, arrData, True)
        Call MarkSeriesDataIfSourceIsHidden(wkb, arrData)
    End If
    
    Call MakeAllStuffVisibleHidden(wkb, True)
    
    If bAreSCsFound Then
        Call AddHyperlinksToChartName(wksSeriesLegend)
        Call AddButtonsThatHyperlinkToCharts(wksSeriesLegend)
        Call AddHyperlinksToSeriesData(wksSeriesLegend)
        
        Call MarkSeriesNameIfXYValuesAreOnDifferentSheets(wksSeriesLegend, arrData)
        Call MarkYValuesIfRowsOrColsDoNotCorrespond(wksSeriesLegend, arrData)
        
        Call ApplyExtensions(wksSeriesLegend, bNewSeriesSheet, arrData)
        
        Call StuffToBeDoneLast(wksSeriesLegend, bNewSeriesSheet)
    End If
    
    Application.StatusBar = False
    
End Sub


'==============================================================================
Private Sub CollectAllHiddenStuffOnSheets( _
    ByVal wkb As Workbook _
)

    Dim wks As Worksheet
    Dim arrInvisibleSheets As Variant
    Dim arrHiddenRanges As Variant
    
    
    If DoesHiddenStuffSheetAlreadyExist(wkb) Then Exit Sub
    Set wks = AddHiddenStuffSheetAndInitialize(wkb)
    
    arrInvisibleSheets = CollectInvisibleSheets(wkb)
    If modArraySupport2.IsArrayAllocated(arrInvisibleSheets) Then
        wks.Cells(1, 1).Resize( _
                UBound(arrInvisibleSheets, 1), _
                UBound(arrInvisibleSheets, 2) _
        ) = arrInvisibleSheets
    End If
    
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

    Dim wks As Worksheet
    
    
    On Error Resume Next
    Set wks = wkb.Worksheets(gcsHiddenSheetName)
    On Error GoTo 0
    DoesHiddenStuffSheetAlreadyExist = (Not wks Is Nothing)
    
End Function


Private Function AddHiddenStuffSheetAndInitialize( _
    ByVal wkb As Workbook _
        ) As Worksheet

    Dim wks As Worksheet
    Dim sActiveSheetName As String
    
    
    sActiveSheetName = ActiveSheet.Name
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

    Dim i As Long
    Dim iHidden As Long
    Dim Arr As Variant
    Dim arrTransposed() As Variant
    Dim bOK As Boolean
    
    
    ReDim Arr(1 To 2, 1 To wkb.Sheets.Count)
    iHidden = 0
    For i = 1 To wkb.Sheets.Count
        With wkb.Sheets(i)
            If Not .Name = gcsLegendSheetName Then
                If .Visible <> xlSheetVisible Then
                    iHidden = iHidden + 1
                    Arr(1, iHidden) = .Name
                    Arr(2, iHidden) = .Visible
                End If
            End If
        End With
    Next
    
    If iHidden > 0 Then
        ReDim Preserve Arr(1 To 2, 1 To iHidden)
        bOK = modArraySupport2.TransposeArray(Arr, arrTransposed)
    End If
    
    CollectInvisibleSheets = arrTransposed

End Function


Private Function CollectHiddenRanges( _
    ByVal wkb As Workbook _
        ) As Variant

    Dim ws As Worksheet
    Dim Arr As Variant
    Dim arrTransposed() As Variant
    Dim sHiddenRows As String
    Dim sHiddenColumns As String
    Dim iHidden As Long
    Dim bOK As Boolean
    
    
    ReDim Arr(1 To 3, 1 To wkb.Worksheets.Count)
    iHidden = 0
    For Each ws In wkb.Worksheets
        sHiddenRows = vbNullString
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
        bOK = modArraySupport2.TransposeArray(Arr, arrTransposed)
    End If
    
    CollectHiddenRanges = arrTransposed

End Function


'------------------------------------------------------------------------------
Private Sub MakeAllStuffVisibleHidden( _
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

    Dim wkb As Workbook
    Dim rng As Range
    Dim Arr As Variant
    Dim i As Long
    
    
    Set rng = wks.Range(gcsInvisibleSheetsRange)
    If rng.Value = vbNullString Then Exit Sub
    
    Arr = rng.CurrentRegion
    Set wkb = wks.Parent
    
    With wkb
        If Not bMakeHidden Then
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

    Dim wkb As Workbook
    Dim wks As Worksheet
    Dim rngHiddenRanges As Range
    Dim arrHiddenRanges As Variant
    Dim Arr As Variant
    Dim i As Long
    Dim j As Long
    
    
    Set rngHiddenRanges = wksHiddenStuff.Range(gcsHiddenRangesRange)
    If rngHiddenRanges.Value = vbNullString Then Exit Sub
    
    'resize to avoid error in case only rows are hidden
    arrHiddenRanges = rngHiddenRanges.CurrentRegion.Resize(, 3)
    Set wkb = wksHiddenStuff.Parent
    
    For i = 1 To UBound(arrHiddenRanges)
        Set wks = wkb.Worksheets(arrHiddenRanges(i, 1))
        If Len(arrHiddenRanges(i, 2)) > 0 Then
            Arr = Split(arrHiddenRanges(i, 2), ",")
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


'==============================================================================

Private Function CollectSCData( _
    ByVal wkb As Workbook, _
    ByRef arrData As Variant _
        ) As Boolean
    
    Dim cha As Chart
    Dim crt As ChartObject
    Dim NoOfAllSCsInAllChartsInWorkbook As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long     'counts the charts
    Dim arrHeading As Variant
    
    
    'Set the default return value
    CollectSCData = False
    
    NoOfAllSCsInAllChartsInWorkbook = CountSCsInAllChartsInWorkbook(wkb)
    If NoOfAllSCsInAllChartsInWorkbook = 0 Then
        MsgBox ("There are no Charts (with readable SeriesCollections) in " & _
                "this Workbook.")
        Exit Function
    End If
    
    'declare bounds of array
        'for that the number of columns is needed which can be extracted from
        ''arrHeading' (+2 because 'arrHeading' is zero based and we need an
        'additional column to store, if 'XValues' and 'Values' are from the
        'same Worksheet)
        arrHeading = TransferHeadingNamesToArray
    ReDim arrData(NoOfAllSCsInAllChartsInWorkbook, UBound(arrHeading) + 2)
    
    With wkb
        'fill the array
        j = 1
        For i = 1 To .Sheets.Count
            If IsChart(wkb, i) Then
                k = k + 1
                Set cha = .Sheets(i)
                Call FillArrayWithSCData(wkb, cha, arrData, i, j, k)
            Else
                For Each crt In .Sheets(i).ChartObjects
                    k = k + 1
                    Set cha = crt.Chart
                    Call FillArrayWithSCData(wkb, cha, arrData, i, j, k)
                Next
            End If
        Next
    End With
    
    CollectSCData = True

End Function


Private Function WasSeriesEntriesInChartsWorksheetCreatedAndInitialized( _
    ByVal wkb As Workbook _
        ) As Boolean

    Dim wks As Worksheet
    Dim bNewSheet As Boolean
    
    
    On Error Resume Next
    Set wks = wkb.Worksheets(gcsLegendSheetName)
    On Error GoTo 0
    If wks Is Nothing Then
        Call CreateAndInitializeSeriesEntriesInChartsWorksheet(wkb)
        bNewSheet = True
    Else
        bNewSheet = False
    End If
    
    'store that the sheet was newly added
    WasSeriesEntriesInChartsWorksheetCreatedAndInitialized = bNewSheet
    
End Function


Private Sub PasteDataToCollectionSheet( _
    ByVal wks As Worksheet, _
    ByVal arrData As Variant _
)

    Dim rng As Range
    Dim iFirstRow As Long
    Dim iLastRow As Long
    Dim i As Long
    Dim iCol As Long
    Dim arrStatisticFormulae As Variant
    Dim sRange As String
    
    
    With wks
        'first clear all entries
        .UsedRange.Offset(gciTitleRow).EntireRow.Delete Shift:=xlShiftUp
        
        Set rng = .Cells(gciTitleRow + 1, 1)
    End With
    
    If Not IsArray(arrData) Then Exit Sub
    
    With rng
        'paste (needed/wanted) array to target range
        .Resize( _
                UBound(arrData, 1), _
                eSD.PlotOrderTotal _
        ).Value = arrData
        
        'find first and last used row
        iFirstRow = .Row
        iLastRow = .Offset(-1, 0).End(xlDown).Row
    End With
    
    'add some "statistic" formulae to the first row showing the number of
    'unique entries in a number of total entries
    arrStatisticFormulae = Array( _
            eSD.SeriesName, _
            eSD.SeriesXValues, _
            eSD.SeriesYValues _
    )
    
    With wks
        For i = LBound(arrStatisticFormulae) To UBound(arrStatisticFormulae)
            iCol = arrStatisticFormulae(i) - rng.Column + 1
            sRange = .Range( _
                    .Cells(iFirstRow, iCol), _
                    .Cells(iLastRow, iCol) _
            ).Address(False, False)
            .Cells(gciTitleRow - 1, iCol).Formula = _
                    "=CONCATENATE(CountUnique(" & sRange & ")," & _
                    Chr(34) & " / " & Chr(34) & ",COUNTA(" & sRange & "))"
        Next
    End With
    
End Sub


Private Sub AddHyperlinksToChartName( _
    ByVal wks As Worksheet _
)

    Dim wkb As Workbook
    Dim iFirstRow As Long
    Dim iLastRow As Long
    Dim iNoOfEntries As Long
    Dim i As Long
    Dim rngSeriesData As Range
    Dim rng As Range
    Dim sChartName As String
    Dim sDataSheet As String
    Dim sTopLeftCell As String
    Dim sHyperlinkTarget As String
    
    
    With wks
        Set wkb = .Parent
        Set rngSeriesData = .Cells(gciTitleRow + 1, 1)
    End With
    
    'find number of entries in list
    iFirstRow = rngSeriesData.Row
    iLastRow = rngSeriesData.Offset(-1, 0).End(xlDown).Row
    iNoOfEntries = iLastRow - iFirstRow + 1
    
    For i = 0 To iNoOfEntries - 1
        sChartName = rngSeriesData.Offset(i, eSD.ChartName - 1).Value
        If Len(sChartName) > 0 Then
            Set rng = rngSeriesData.Offset(i, eSD.ChartName - 1)
            sDataSheet = rngSeriesData.Offset(i, eSD.SheetName - 1).Value
            sTopLeftCell = wkb.Worksheets(sDataSheet). _
                    ChartObjects(sChartName).TopLeftCell.Address(False, False)
            sHyperlinkTarget = "'" & sDataSheet & "'!" & sTopLeftCell
            Call AddHyperlinkToCurrentCell(wks, rng, sHyperlinkTarget)
        End If
    Next
    
    'format the columns
    Set rng = rngSeriesData.Offset( _
            0, _
            eSD.SheetName - 1 _
    ).Resize( _
            iNoOfEntries, _
            eSD.ChartName - eSD.SheetName + 1 _
    )
    With rng.Font
        .ColorIndex = xlColorIndexAutomatic
        .Underline = xlUnderlineStyleNone
'        .Size = wks.Cells(2, l).Font.Size
    End With
    
End Sub


Private Sub AddButtonsThatHyperlinkToCharts( _
    ByVal wks As Worksheet _
)

    Dim iFirstRow As Long
    Dim iLastRow As Long
    Dim iNoOfEntries As Long
    Dim i As Long
    Dim rngSeriesData As Range
    Dim rng As Range
    Dim sSheetName As String
    Dim sChartName As String
    
    
    Call DeleteAllLabelShapesOnSheet(wks)
    
    Set rngSeriesData = wks.Cells(gciTitleRow + 1, 1)
    
    iFirstRow = rngSeriesData.Row
    iLastRow = rngSeriesData.Offset(-1, 0).End(xlDown).Row
    iNoOfEntries = iLastRow - iFirstRow + 1
    
    For i = 0 To iNoOfEntries - 1
        sSheetName = rngSeriesData.Offset(i, eSD.SheetName - 1).Value
        sChartName = rngSeriesData.Offset(i, eSD.ChartName - 1).Value
        If Len(sSheetName) > 0 And Len(sChartName) = 0 Then
            Set rng = rngSeriesData.Offset(i, eSD.SheetName - 1)
            Call AddLabelToCell(rng)
        End If
    Next

End Sub


Private Sub DeleteAllLabelShapesOnSheet( _
    ByVal wks As Worksheet _
)

    Dim shp As Shape
    
    
    For Each shp In wks.Shapes
        If shp.Type = msoTextBox Then
            shp.Delete
        End If
    Next

End Sub


Private Sub AddLabelToCell( _
    ByVal rng As Range _
)

    Dim wks As Worksheet
    Dim shp As Shape
    
    '==========================================================================
    Const csNamePrefix As String = "lblChart"
    '==========================================================================
    
    
    With rng
        Set wks = .Parent
        Set shp = wks.Shapes.AddLabel( _
                msoTextOrientationHorizontal, _
                .Left, .Top, .Width, .Height _
        )
    End With
    With shp
        .Name = csNamePrefix & wks.Cells(rng.Row, 1)
        .OnAction = "GotoChartWithCellName"
    End With

End Sub


'inspired by <https://excel.tips.net/T002539_Hyperlinks_to_Charts.html>
Public Sub GotoChartWithCellName()

    Dim wkb As Workbook
    Dim rng As Range
    
    
    Set rng = ActiveSheet.Shapes(Application.Caller).TopLeftCell
    Set wkb = ActiveSheet.Parent
    On Error GoTo errHandler
    wkb.Sheets(rng.Value).Select
    On Error GoTo 0
    
    Exit Sub
    
    
errHandler:
    MsgBox ("The chart '" & rng.Value & "' can't be found.")
    Exit Sub
    
End Sub


Private Sub AddHyperlinksToSeriesData( _
    ByVal wks As Worksheet _
)

    Dim i As Long
    Dim j As Long
    Dim iFirstRow As Long
    Dim iLastRow As Long
    Dim iNoOfEntries As Long
    Dim rngSeriesData As Range
    Dim rng As Range
    Dim rngTest As Range
    Dim sDataSheet As String
    Dim sRngValue As String
    Dim sFirstCell As String
    Dim sHyperlinkTarget As String
    
    Dim sListSeparator As String
    
    
    sListSeparator = Application.International(xlListSeparator)
    
    Set rngSeriesData = wks.Cells(gciTitleRow + 1, 1)
    
    'find number of entries in list
    With rngSeriesData
        iFirstRow = .Row
        iLastRow = .Offset(-1, 0).End(xlDown).Row
        iNoOfEntries = iLastRow - iFirstRow + 1
    End With
    
    For i = 0 To iNoOfEntries - 1
        sDataSheet = rngSeriesData.Offset(i, eSD.SeriesDataSheet - 1).Value
        For j = eSD.SeriesXValues - 1 To eSD.SeriesYValues - 1
            Set rng = rngSeriesData.Offset(i, j)
            sRngValue = rng.Value
            sRngValue = Replace(sRngValue, ",", sListSeparator)
'---
'because currently commata in worksheet names are not supported, the entries
'in the cells could be wrong and thus the following line can cause an error
'--> for now, skip them then
On Error GoTo SkipAddingHyperlink
'interestingly this 'GoTo' also doesn't work ...
'---
            On Error Resume Next
            Set rngTest = wks.Range(sRngValue)
            If Not rngTest Is Nothing Then
                sFirstCell = rngTest.Areas(1).Cells(1).Address(False, False)
                sHyperlinkTarget = "'" & sDataSheet & "'!" & sFirstCell
                Call AddHyperlinkToCurrentCell(wks, rng, sHyperlinkTarget)
            End If
            Set rngTest = Nothing
'---
SkipAddingHyperlink:
On Error GoTo 0
'---
        Next
    Next
    
    'format the columns
    Set rng = rngSeriesData.Offset( _
            0, _
            eSD.SeriesXValues - 1 _
    ).Resize( _
            iNoOfEntries, _
            eSD.SeriesYValues - eSD.SeriesXValues + 1 _
    )
    With rng.Font
        .ColorIndex = xlColorIndexAutomatic
        .Underline = xlUnderlineStyleNone
        .Size = wks.Cells(gciTitleRow + 1, 1).Font.Size
    End With
    
End Sub


Private Sub MarkEachOddChartNumber( _
    ByVal wks As Worksheet _
)

    Dim i As Long
    Dim iFirstRow As Long
    Dim iLastRow As Long
    Dim iNoOfEntries As Long
    Dim rngSeriesData As Range
    
    '==========================================================================
    'color for odd chart numbers
    Const ccOddChartNumbers As Long = 15853276   'R=220 G=230 B=241
    '==========================================================================
    
    
    Set rngSeriesData = wks.Cells(gciTitleRow + 1, 1)
    
    'find number of entries in list
    With rngSeriesData
        iFirstRow = .Row
        iLastRow = .Offset(-1, 0).End(xlDown).Row
        iNoOfEntries = iLastRow - iFirstRow + 1
    End With
    
    With rngSeriesData
        For i = 0 To iNoOfEntries - 1
            If .Offset(i, eSD.ChartNumber - 1).Value Mod 2 = 1 Then
                .Offset(i).EntireRow.Interior.Color = ccOddChartNumbers
            Else
                .Offset(i).EntireRow.Interior.Color = xlColorIndexNone
            End If
        Next
    End With

End Sub


'stuff that has to be done last
Private Sub StuffToBeDoneLast( _
    ByVal wks As Worksheet, _
    ByVal bNewSheet As Boolean _
)
    
    'do it only if the sheet was newly created
    If bNewSheet = True Then
        With wks
            'set 'AutoFilter' and 'AutoFit'
            .Rows(gciTitleRow).AutoFilter
            
            ''AutoFit' the 'UsedRange'
            .UsedRange.EntireColumn.AutoFit
        End With
        
        With ActiveWindow
            'freeze the panes
            .SplitRow = 2
            .SplitColumn = 3
            .FreezePanes = True
            'set the zoom factor
            .Zoom = 70
        End With
        
'        'optionally the group level can be changed that all are closed
'        wks.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    End If
End Sub


Private Function CountSCsInAllChartsInWorkbook( _
    ByVal wkb As Workbook _
        ) As Long

    Dim i As Long
    Dim j As Long
    Dim sc As Series
    Dim crt As ChartObject
    
    
    'just count the number of entries for the array first for all
    j = 0
    With wkb
        'it gets a bit complicated, because charts can occur as
        '- dedicated "Chart" sheets and
        '- as "ChartObjects" on normal WorkSheets
        For i = 1 To .Sheets.Count
            If IsChart(wkb, i) Then
                For Each sc In .Sheets(i).SeriesCollection
                    j = j + 1
                Next
            Else
                For Each crt In .Sheets(i).ChartObjects
                    For Each sc In crt.Chart.SeriesCollection
                        j = j + 1
                    Next
                Next
            End If
        Next
    End With
    
    CountSCsInAllChartsInWorkbook = j
    
End Function


Private Sub CreateAndInitializeSeriesEntriesInChartsWorksheet( _
    ByVal wkb As Workbook _
)

    Dim wks As Worksheet
    Dim arrHeading As Variant
    
    
    arrHeading = TransferHeadingNamesToArray
    
    Set wks = wkb.Worksheets.Add(, wkb.Worksheets(wkb.Worksheets.Count))
    '"configure" the new sheet
    With wks
        .Name = gcsLegendSheetName
        .Tab.ThemeColor = xlThemeColorLight1
        .Tab.TintAndShade = 0
        
        'set column titles
        .Cells(gciTitleRow, 1).Resize(, UBound(arrHeading) + 1) = arrHeading
        
        'AxisGroup
        With .Cells(gciTitleRow, eSD.AxisGroup)
            .AddComment ( _
                    "AxisGroup" & Chr(10) & _
                    "1 = Primary Axis" & Chr(10) & _
                    "2 = Secondary Axis" _
            )
            .Comment.Shape.TextFrame.AutoSize = True
        End With
        'PlotOrder per AxisGroup
        With .Cells(gciTitleRow, eSD.PlotOrder)
            .AddComment ("PlotOrder per AxisGroup")
            .Comment.Shape.TextFrame.AutoSize = True
        End With
        'PlotOrder total
        With .Cells(gciTitleRow, eSD.PlotOrderTotal)
            .AddComment ("PlotOrder total")
            .Comment.Shape.TextFrame.AutoSize = True
        End With
        
        'add groups to some columns
        .Columns(eSD.ChartName).Group
        .Columns(eSD.Y2Label).Group
        .Columns(eSD.SeriesDataSheet).Group
        '----------------------------------------------------------------------
        
        'change style of title row
        With .Rows(gciTitleRow)
            With .Interior
                'background color
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Font
                'font color
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End With
        
        'change page setup
        With .PageSetup
            'set header
            .LeftHeader = "&Z&F" & Chr(10) & "&A"
            .RightHeader = vbNullString & Chr(10) & "&D"
            'set page orientation
            .Orientation = xlLandscape
            'print everything on one page
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            'print these cells on each page
            .PrintTitleRows = "$2:$2"
            .PrintTitleColumns = "$A:$C"
        End With
    End With
End Sub


Private Function TransferHeadingNamesToArray() As Variant

    Dim arrHeading As Variant
    
    '==========================================================================
    'which string separator should be used
    Const csStringSep As String = ";"
    'which columns should be written to the table?
    Const csHeadingNames As String = _
            "No." & csStringSep & _
            "Sheet.Name" & csStringSep & _
            "Chart.Name" & csStringSep & _
            "Chart.Title" & csStringSep & _
            "x axis" & csStringSep & _
            "y axis" & csStringSep & _
            "y axis 2" & csStringSep & _
            "Series.Name" & csStringSep & _
            "Series.DataSheet" & csStringSep & _
            "Series.XValues" & csStringSep & _
            "Series.Values" & csStringSep & _
            "AG" & csStringSep & _
            "PO" & csStringSep & _
            "POt"
    '==========================================================================
    
    
    arrHeading = Split(csHeadingNames, csStringSep)
    
    TransferHeadingNamesToArray = arrHeading
    
End Function


Private Sub AddHyperlinkToCurrentCell( _
    ByVal wks As Worksheet, _
    ByRef rng As Range, _
    ByVal sHyperlinkTarget As String _
)

    'first test, if 'sHyperlinkTarget' is a (valid) range
    'if not, there is no hyperlink to add
    If Not RangeExists(wks.Parent, sHyperlinkTarget) Then Exit Sub
    
    With rng
        'if hyperlink already set, delete them first
        Do Until .Hyperlinks.Count = 0
            .Hyperlinks(1).Delete
        Loop
    End With
    
    With wks
        .Hyperlinks.Add _
                Anchor:=rng, _
                Address:=vbNullString, _
                SubAddress:=sHyperlinkTarget, _
                ScreenTip:=sHyperlinkTarget
    End With
End Sub


Private Sub FillArrayWithSCData( _
    ByVal wkb As Workbook, _
    ByVal cha As Chart, _
    ByRef arrData As Variant, _
    ByRef i As Long, _
    ByRef j As Long, _
    ByRef k As Long _
)

    Dim myChart As clsChartSeries
    Dim l As Long
    
    
    With wkb
        arrData(j, eSD.ChartNumber) = k
        arrData(j, eSD.SheetName) = .Sheets(i).Name
        If Not .Name = cha.Parent.Name Then
            arrData(j, eSD.ChartName) = cha.Parent.Name
        End If
    End With
    
    With cha
        If .HasTitle Then arrData(j, eSD.ChartTitle) = .ChartTitle.Text
        If .Axes(xlCategory).HasTitle Then _
                arrData(j, eSD.XLabel) = .Axes(xlCategory).AxisTitle.Text
        If .Axes(xlValue).HasTitle Then _
                arrData(j, eSD.YLabel) = .Axes(xlValue, xlPrimary).AxisTitle.Text
        If .HasAxis(xlValue, xlSecondary) Then
            If .Axes(xlValue, xlSecondary).HasTitle Then
                arrData(j, eSD.Y2Label) = .Axes(xlValue, xlSecondary).AxisTitle.Text
            End If
        End If
        For l = 1 To .SeriesCollection.Count
            arrData(j, eSD.ChartNumber) = k
            Set myChart = New clsChartSeries
            With myChart
                .Chart = cha
                .ChartSeries = l
'                arrData(j, eSD.SeriesName) = .SeriesName
                arrData(j, eSD.SeriesName) = cha.SeriesCollection(l).Name
                
                Select Case .XValuesType
                    Case "Range"
                        arrData(j, eSD.SeriesXValues) = .XValues.Address(False, False)
                    Case "inaccessible"
                        arrData(j, eSD.SeriesXValues) = "#REF"
                    Case Else
                        arrData(j, eSD.SeriesXValues) = .XValues
                End Select
                
                Select Case .ValuesType
                    Case "Range"
                        arrData(j, eSD.SeriesDataSheet) = .DataSheet(3)
                        arrData(j, eSD.SeriesYValues) = .Values.Address(False, False)
                    Case "inaccessible"
                        arrData(j, eSD.SeriesYValues) = "#REF"
                    Case Else
                        arrData(j, eSD.SeriesYValues) = .Values
                End Select
                
                Select Case cha.SeriesCollection(l).AxisGroup
                    Case xlPrimary
                        arrData(j, eSD.AxisGroup) = 1
                    Case xlSecondary
                        arrData(j, eSD.AxisGroup) = 2
                End Select
                
                arrData(j, eSD.PlotOrder) = cha.SeriesCollection(l).PlotOrder
                arrData(j, eSD.PlotOrderTotal) = .PlotOrder
                'we don't want to have written a possible "inaccessible" in the
                'list, so we overwrite it in that case
                If arrData(j, eSD.PlotOrderTotal) = "inaccessible" Then
                    arrData(j, eSD.PlotOrderTotal) = vbNullString
                End If
                
                arrData(j, eSD.XYDataSheetEqual) = (arrData(j, eSD.SeriesDataSheet) = .DataSheet(2))
                
                j = j + 1
            End With
        Next
    End With
End Sub


'==============================================================================
Private Sub MarkSheetNameOrSeriesDataSheetIfSourceIsInvisible( _
    ByVal wkb As Workbook, _
    ByVal arrData As Variant, _
    Optional ByVal bSheetName_Or_SeriesDataSheet As Boolean = False _
)

    Dim wksHiddenSheet As Worksheet
    Dim wksSeriesLegend As Worksheet
    Dim rngInvisibleSheets As Range
    Dim rng As Range
    Dim arrInvisibleSheets As Variant
    Dim arrToMarkCells As Variant
    Dim iCol As Long
    Dim i As Long
    Dim j As Long
    
    '==========================================================================
    'color for stuff that is hidden/invisible (fully)
    Const ccHidden As Long = 12040422            'R=230 G=184 B=183
    'which cells should be marked/tested
    arrToMarkCells = Array( _
            eSD.SeriesDataSheet, _
            eSD.SeriesXValues, _
            eSD.SeriesYValues _
    )
    '==========================================================================
    
    
    Set wksHiddenSheet = wkb.Worksheets(gcsHiddenSheetName)
    Set rngInvisibleSheets = wksHiddenSheet.Range(gcsInvisibleSheetsRange)
    If rngInvisibleSheets.Value = vbNullString Then Exit Sub
    arrInvisibleSheets = rngInvisibleSheets.CurrentRegion.Value
    
    Set wksSeriesLegend = wkb.Worksheets(gcsLegendSheetName)
    Set rng = wksSeriesLegend.Cells(gciTitleRow, 1)
    
    For i = LBound(arrData) To UBound(arrData)
        If Not bSheetName_Or_SeriesDataSheet Then
            If IsInFirstColOf2DArray( _
                    arrData(i, eSD.SheetName), _
                    arrInvisibleSheets _
            ) Then
                rng.Offset(i, eSD.SheetName - 1).Interior.Color = ccHidden
            End If
        Else
            If IsInFirstColOf2DArray( _
                    arrData(i, eSD.SeriesDataSheet), _
                    arrInvisibleSheets _
            ) Then
                For j = LBound(arrToMarkCells) To UBound(arrToMarkCells)
                    iCol = arrToMarkCells(j)
                    rng.Offset(i, iCol - 1).Interior.Color = ccHidden
                Next
            End If
        End If
    Next

End Sub


Private Sub MarkSeriesDataIfSourceIsHidden( _
    ByVal wkb As Workbook, _
    ByVal arrData As Variant _
)

    Dim wksHiddenSheet As Worksheet
    Dim wksSeriesLegend As Worksheet
    Dim rngHiddenRanges As Range
    Dim rng As Range
    Dim arrHiddenRanges As Variant
    Dim arrToMarkCells As Variant
    Dim iCol As Long
    Dim i As Long
    Dim j As Long
    
    '==========================================================================
    'color for stuff that is hidden/invisible (fully)
    Const ccHidden As Long = 12040422            'R=230 G=184 B=183
    'color for stuff that is hidden partially
    Const ccHiddenPartly As Long = 14408946      'R=242 G=220 B=219
    'which cells should be marked/tested
    arrToMarkCells = Array( _
            eSD.SeriesXValues, _
            eSD.SeriesYValues _
    )
    '==========================================================================
    
    
    Set wksHiddenSheet = wkb.Worksheets(gcsHiddenSheetName)
    Set rngHiddenRanges = wksHiddenSheet.Range(gcsHiddenRangesRange)
    If rngHiddenRanges.Value = vbNullString Then Exit Sub
    arrHiddenRanges = rngHiddenRanges.CurrentRegion.Value
    
    Set wksSeriesLegend = wkb.Worksheets(gcsLegendSheetName)
    Set rng = wksSeriesLegend.Cells(gciTitleRow, 1)
    
    For i = LBound(arrData) To UBound(arrData)
        If rng.Offset(i, eSD.SeriesDataSheet - 1) _
                .Interior.Color <> ccHidden Then
            For j = LBound(arrToMarkCells) To UBound(arrToMarkCells)
                iCol = arrToMarkCells(j)
                Select Case IsRangeHidden( _
                        wkb, _
                        arrData(i, eSD.SeriesDataSheet), _
                        arrData(i, iCol), _
                        arrHiddenRanges _
                )
                    Case 1
                        rng.Offset(i, iCol - 1).Interior.Color = ccHiddenPartly
                    Case 2
                        rng.Offset(i, iCol - 1).Interior.Color = ccHidden
                End Select
            Next
        End If
    Next

End Sub


Private Function IsInFirstColOf2DArray( _
    ByVal sString As String, _
    ByVal vArr As Variant _
        ) As Boolean
    
    Dim i As Long
    
    
    IsInFirstColOf2DArray = False
    For i = LBound(vArr) To UBound(vArr)
        If vArr(i, 1) = sString Then
            IsInFirstColOf2DArray = True
            Exit Function
        End If
    Next
End Function


Private Function IsRangeHidden( _
    ByVal wkb As Workbook, _
    ByVal sWorksheet As String, _
    ByVal sRange As String, _
    ByVal arrHiddenRanges As Variant _
        ) As Long
    
    Dim i As Long
    Dim iHidden() As Long
    Dim arrToTest As Variant
    Dim arrHidden As Variant
    Dim sHidden As String
    Dim Element As Variant
    Dim rngToTest As Range
    Dim bWksFound As Boolean
    
    
    If Len(sRange) = 0 Then Exit Function
    
    'test if there is a hidden range on 'sWorksheet' and if yes,
    'create an array of hidden areas ('arrHidden')
    For i = LBound(arrHiddenRanges) To UBound(arrHiddenRanges)
        If sWorksheet = arrHiddenRanges(i, 1) Then
            arrHidden = CreateArrayOfHiddenAreas( _
                    arrHiddenRanges(i, 2), _
                    arrHiddenRanges(i, 3) _
            )
            bWksFound = True
            Exit For
        End If
    Next
    If Not bWksFound Then
        IsRangeHidden = 0
        Exit Function
    End If
    
    arrToTest = Split(sRange, ",")
    ReDim iHidden(LBound(arrToTest) To UBound(arrToTest))
    
    For i = LBound(arrToTest) To UBound(arrToTest)
        Set rngToTest = wkb.Worksheets(sWorksheet).Range(arrToTest(i))
        iHidden(i) = IsAreaHidden(rngToTest, arrHidden)
    Next
    
    IsRangeHidden = ReturnHiddenState(iHidden)
    
End Function


Private Function CreateArrayOfHiddenAreas( _
    ByVal sRows As String, _
    ByVal sColumns As String _
        ) As Variant
    
    Dim sRange As String
    Dim Arr As Variant
    
    
    If Len(sRows) > 0 And Len(sColumns) > 0 Then
        sRange = sRows & "," & sColumns
    ElseIf Len(sRows) > 0 Then
        sRange = sRows
    Else
        sRange = sColumns
    End If
    Arr = Split(sRange, ",")
    
    CreateArrayOfHiddenAreas = Arr

End Function


Private Function IsAreaHidden( _
    ByVal rngToTest As Range, _
    ByVal Arr As Variant _
        ) As Long

    Dim wks As Worksheet
    Dim rng As Range
    Dim rngIntersect As Range
    Dim iHidden As Long
    Dim i As Long
    
    
    Set wks = rngToTest.Parent
    For i = LBound(Arr) To UBound(Arr)
        Set rng = wks.Range(Arr(i))
        Set rngIntersect = Application.Intersect(rngToTest, rng)
        If rngIntersect Is Nothing Then
            iHidden = Application.WorksheetFunction.Max(0, iHidden)
        ElseIf rngToTest.Address = rngIntersect.Address Then
            iHidden = Application.WorksheetFunction.Max(2, iHidden)
        Else
            iHidden = Application.WorksheetFunction.Max(1, iHidden)
        End If
    Next
    
    IsAreaHidden = iHidden

End Function


Private Function ReturnHiddenState( _
    ByVal iHidden As Variant _
        ) As Long
    If Application.WorksheetFunction.Average(iHidden) = 2 Then
        ReturnHiddenState = 2
    ElseIf Application.WorksheetFunction.sum(iHidden) = 0 Then
        ReturnHiddenState = 0
    Else
        ReturnHiddenState = 1
    End If
End Function


Private Sub MarkSeriesNameIfXYValuesAreOnDifferentSheets( _
    ByVal wksSeriesLegend As Worksheet, _
    ByVal arrData As Variant _
)

    Dim i As Long
    Dim rng As Range
    
    '==========================================================================
    'color for stuff that is "wrong"
    Const ccWrong As Long = 255            'R=255 G=0 B=0
    '==========================================================================
    
    
    Set rng = wksSeriesLegend.Cells(gciTitleRow, 1)
    For i = LBound(arrData) To UBound(arrData)
        If Not arrData(i, eSD.XYDataSheetEqual) Then
            rng.Offset(i, eSD.SeriesDataSheet - 1).Font.Color = ccWrong
        End If
    Next

End Sub


Private Sub MarkYValuesIfRowsOrColsDoNotCorrespond( _
    ByVal wksSeriesLegend As Worksheet, _
    ByVal arrData As Variant _
)

    Dim i As Long
    Dim rng As Range
    
    '==========================================================================
    'color for stuff that is "wrong"
    Const ccWrong As Long = 255            'R=255 G=0 B=0
    '==========================================================================
    
    
    Set rng = wksSeriesLegend.Cells(gciTitleRow, 1)
    For i = LBound(arrData) To UBound(arrData)
        If Not AreRowsOrColsEqual( _
                arrData(i, eSD.SeriesXValues), _
                arrData(i, eSD.SeriesYValues) _
        ) Then
            rng.Offset(i, eSD.SeriesYValues - 1).Font.Color = ccWrong
        End If
    Next

End Sub


Private Function AreRowsOrColsEqual( _
    ByVal sXRange As String, _
    ByVal sYRange As String _
        ) As Boolean

    Dim bAreRowsEqual As Boolean
    Dim bAreColsEqual As Boolean
    
    
    'if one of the ranges is not a range, no test is needed
    If Left(sXRange, 1) = "{" Or Left(sYRange, 1) = "{" Then Exit Function
    
    bAreRowsEqual = (ExtractRowsRange(sXRange) = ExtractRowsRange(sYRange))
    bAreColsEqual = (ExtractColumnsRange(sXRange) = ExtractColumnsRange(sYRange))
    
    AreRowsOrColsEqual = (bAreRowsEqual Or bAreColsEqual)
    
End Function
