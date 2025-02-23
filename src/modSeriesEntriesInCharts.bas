Attribute VB_Name = "modSeriesEntriesInCharts"

'@Folder("ChartSeries")

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
    SeriesNamePath
    SeriesNameBook
    SeriesNameSheet
    SeriesNameRange
    SeriesNameValue
    SeriesXPath
    SeriesXBook
    SeriesXSheet
    SeriesXValues
    SeriesYPath
    SeriesYBook
    SeriesYSheet
    SeriesYValues
    AxisGroup
    PlotOrder
    PlotOrderTotal
    [_Last] = eSD.PlotOrderTotal
End Enum


'==============================================================================
'sheet name of pasted Series data?
Private Const pcsLegendSheetName As String = "SeriesEntriesInCharts"
'title row on 'pcsLegendSheetName'
Public Const gciTitleRow As Long = 2
'what is written to 'Series.Values' (etc.) when the series is not accessible?
Private Const pcsInaccessible = "#Inaccessible"
'==============================================================================


Public Sub ListAllSCEntriesInAllCharts()
    
    Dim SpeedUp As XLSpeedUp
    Set SpeedUp = New XLSpeedUp
    SpeedUp.TurnOn statusBarMessage:="Running 'ListAllSCEntriesInAllCharts' ..."
    
    Dim wkb As Workbook
    Set wkb = ActiveWorkbook
    
    Call CollectAllHiddenStuffOnSheets(wkb)
    Call MakeAllStuffVisibleHidden(wkb, False)
    
    Dim bAreSCsFound As Boolean
    Dim arrData As Variant
    bAreSCsFound = CollectSCData(wkb, arrData)
    
    If bAreSCsFound Then
        Call RepeatSheetAndChartNames(arrData)
        
        'store if the sheet was newly added/created
        Dim bNewSeriesSheet As Boolean
        bNewSeriesSheet = WasSeriesEntriesInChartsWorksheetCreatedAndInitialized(wkb)
        
        Dim wksSeriesLegend As Worksheet
        Set wksSeriesLegend = wkb.Worksheets(pcsLegendSheetName)
        
        If Not bNewSeriesSheet Then
            Dim AFStorage As IAutoFilterStorage
            Set AFStorage = AutoFilterStorage.Create(wksSeriesLegend)
            AFStorage.StoreFilters
            AFStorage.AutoFilterMode = False
        End If
        
        Call PasteDataToCollectionSheet(wksSeriesLegend, arrData)
        Call ShowStartingQuotesInBookNames(wksSeriesLegend)
        
        Call MarkEachOddChartNumberRow(wksSeriesLegend)
        
        Call MarkSheetEntriesIfSourceIsInvisible(wkb, arrData)
        Call MarkSeriesDataIfSourceIsHidden(wkb, arrData)
    End If
    
    Call MakeAllStuffVisibleHidden(wkb, True)
    
    If bAreSCsFound Then
        Call AddHyperlinksToChartName(wksSeriesLegend)
        'needs to be done after `AddHyperlinksToChartName`
        'otherwise colors will be overwritten (again)
        Call ChangeColorOfRepeatingSheetChartNames(wksSeriesLegend, arrData)
        Call AddButtonsThatHyperlinkToCharts(wksSeriesLegend)
        Call AddHyperlinksToSeriesData(wksSeriesLegend)
        
        Call MarkSeriesSheetIfXYValuesAreOnDifferentSheets(wksSeriesLegend, arrData)
        Call MarkYValuesIfRowsOrColsDoNotCorrespond(wksSeriesLegend, arrData)
        
        Call ApplyExtensions(wksSeriesLegend, bNewSeriesSheet, arrData)
        
        Call StuffToBeDoneLast(wksSeriesLegend, bNewSeriesSheet)
        Call HideSeriesPathBookIfEmpty(wksSeriesLegend, arrData)
        
        If Not bNewSeriesSheet Then
            AFStorage.RestoreFilters
        End If
    End If
    
tidyUp:
    SpeedUp.TurnOff
    Set SpeedUp = Nothing
    
End Sub


'==============================================================================

Private Function CollectSCData( _
    ByVal wkb As Workbook, _
    ByRef arrData As Variant _
        ) As Boolean
    
    'Set the default return value
    CollectSCData = False
    
    Dim NoOfAllSCsInAllChartsInWorkbook As Long
    NoOfAllSCsInAllChartsInWorkbook = CountSCsInAllChartsInWorkbook(wkb)
    
    Dim NoOfChartsWithoutSeriesInWorkbook As Long
    NoOfChartsWithoutSeriesInWorkbook = CountChartsWithoutSeriesInWorkbook(wkb)
    
    If NoOfAllSCsInAllChartsInWorkbook = 0 Then
        MsgBox ("There are no Charts (with readable SeriesCollections) in " & _
                "this Workbook.")
        Exit Function
    End If
    
    'declare bounds of array
    '(for that the number of columns is needed which can be extracted from
    ' 'arrHeading' (+1 because 'arrHeading' is zero based))
    Dim arrHeading As Variant
    arrHeading = TransferHeadingNamesToArray
    ReDim arrData(1 To NoOfAllSCsInAllChartsInWorkbook + NoOfChartsWithoutSeriesInWorkbook, 1 To UBound(arrHeading) + 1)
    Erase arrHeading
    
    'fill the array
    Dim iSCTotal As Long
    iSCTotal = 1
    
    Dim iSheetIndex As Long
    For iSheetIndex = 1 To wkb.Sheets.Count
        If IsChart(wkb, iSheetIndex) Then
            Dim iChartNumber As Long
            iChartNumber = iChartNumber + 1
            
            Dim cha As Chart
            Set cha = wkb.Sheets(iSheetIndex)
            
            Call FillArrayWithSCData( _
                    wkb, _
                    cha, _
                    arrData, _
                    iSheetIndex, _
                    iSCTotal, _
                    iChartNumber _
            )
        Else
            Dim iChartObjectIndex As Long
            For iChartObjectIndex = 1 To wkb.Sheets(iSheetIndex).ChartObjects.Count
                iChartNumber = iChartNumber + 1
                
                Dim crt As ChartObject
                Set crt = wkb.Sheets(iSheetIndex).ChartObjects(iChartObjectIndex)
                Set cha = crt.Chart
                
                Call FillArrayWithSCData( _
                        wkb, _
                        cha, _
                        arrData, _
                        iSheetIndex, _
                        iSCTotal, _
                        iChartNumber _
                )
            Next
        End If
    Next
    
    CollectSCData = True
    
End Function


Private Sub RepeatSheetAndChartNames( _
    ByRef arrData As Variant _
)
    
    Dim i As Long
    For i = LBound(arrData) To UBound(arrData)
        If Len(arrData(i, eSD.SheetName)) > 0 Then
            Dim SheetName As String
            SheetName = arrData(i, eSD.SheetName)
            
            Dim ChartName As String
            ChartName = arrData(i, eSD.ChartName)
        Else
            arrData(i, eSD.SheetName) = SheetName
            If Len(ChartName) > 0 Then
                arrData(i, eSD.ChartName) = ChartName
            End If
        End If
    Next
    
End Sub


Private Function WasSeriesEntriesInChartsWorksheetCreatedAndInitialized( _
    ByVal wkb As Workbook _
        ) As Boolean
    
    On Error Resume Next
    Dim wks As Worksheet
    Set wks = wkb.Worksheets(pcsLegendSheetName)
    On Error GoTo 0
    
    If wks Is Nothing Then
        Call CreateAndInitializeSeriesEntriesInChartsWorksheet(wkb)
        Dim bNewSheet As Boolean
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
    
    With wks
        'first clear all entries
        .UsedRange.Offset(gciTitleRow).EntireRow.Delete Shift:=xlShiftUp
        
        Dim rng As Range
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
        Dim iFirstRow As Long
        iFirstRow = .Row
        
        Dim iLastRow As Long
        iLastRow = .Offset(-1, 0).End(xlDown).Row
    End With
    
    'add some "statistic" formulae to the first row showing the number of
    'unique entries in a number of total entries
    Dim arrStatisticFormulae As Variant
    arrStatisticFormulae = Array( _
            eSD.SeriesNamePath, _
            eSD.SeriesNameBook, _
            eSD.SeriesNameSheet, _
            eSD.SeriesNameRange, _
            eSD.SeriesNameValue, _
            eSD.SeriesXPath, _
            eSD.SeriesXBook, _
            eSD.SeriesXSheet, _
            eSD.SeriesXValues, _
            eSD.SeriesYPath, _
            eSD.SeriesYBook, _
            eSD.SeriesYSheet, _
            eSD.SeriesYValues _
    )
    
    With wks
        Dim i As Long
        For i = LBound(arrStatisticFormulae) To UBound(arrStatisticFormulae)
            Dim iCol As Long
            iCol = arrStatisticFormulae(i) - rng.Column + 1
            
            Dim sRange As String
            sRange = .Range( _
                    .Cells(iFirstRow, iCol), _
                    .Cells(iLastRow, iCol) _
            ).Address(False, False)
            .Cells(gciTitleRow - 1, iCol).Formula = _
                    "=CONCATENATE(CountUnique(" & sRange & ")," & _
                    Chr$(34) & " / " & Chr$(34) & ",COUNTA(" & sRange & "))"
        Next
    End With
    
End Sub


Private Sub ShowStartingQuotesInBookNames( _
    ByVal wks As Worksheet _
)
    
    Dim rngSeriesData As Range
    Set rngSeriesData = wks.Cells(gciTitleRow + 1, 1)
    
    Dim iNoOfEntries As Long
    iNoOfEntries = NoOfEntriesInList(rngSeriesData)
    
    Dim arrBooks As Variant
    arrBooks = Array( _
            eSD.SeriesXBook - 1, _
            eSD.SeriesYBook - 1 _
    )
    
    Dim i As Long
    For i = 0 To iNoOfEntries - 1
        Dim j As Long
        For j = LBound(arrBooks) To UBound(arrBooks)
            Dim rng As Range
            Set rng = rngSeriesData.Offset(i, arrBooks(j))
            
            With rng
                If .PrefixCharacter = "'" Then
                    .Value2 = "''" & .Value2
                End If
            End With
        Next
    Next
    
End Sub


Private Sub ChangeColorOfRepeatingSheetChartNames( _
    ByVal wksSeriesLegend As Worksheet, _
    ByVal arrData As Variant _
)
    
    '==========================================================================
    'color for repeating entries
    Const ccNotEqual As Long = 255            'R=255 G=0 B=0
    '==========================================================================
    
    Dim rng As Range
    Set rng = wksSeriesLegend.Cells(gciTitleRow, 1)
    
    Dim ChartNumber As Long
    ChartNumber = 0
    
    Dim i As Long
    For i = LBound(arrData) To UBound(arrData)
        If arrData(i, eSD.ChartNumber) <> ChartNumber Then
            ChartNumber = arrData(i, eSD.ChartNumber)
            
            Dim SheetName As String
            SheetName = arrData(i, eSD.SheetName)
            
            Dim ChartName As String
            ChartName = arrData(i, eSD.ChartName)
        Else
            With rng.Offset(i, eSD.SheetName - 1).Font
                'font color
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.5
            End With
            
            With rng.Offset(i, eSD.ChartName - 1).Font
                If Len(ChartName) > 0 Then
                    'font color
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.5
                End If
            End With
        End If
    Next
    
End Sub


Private Sub AddHyperlinksToChartName( _
    ByVal wks As Worksheet _
)
    
    With wks
        Dim wkb As Workbook
        Set wkb = .Parent
        
        Dim rngSeriesData As Range
        Set rngSeriesData = .Cells(gciTitleRow + 1, 1)
    End With
    
    'find number of entries in list
    Dim iFirstRow As Long
    iFirstRow = rngSeriesData.Row
    
    Dim iLastRow As Long
    iLastRow = rngSeriesData.Offset(-1, 0).End(xlDown).Row
    
    Dim iNoOfEntries As Long
    iNoOfEntries = iLastRow - iFirstRow + 1
    
    Dim CurrentChart As Chart
    Set CurrentChart = RememberActiveChartAndActivateGivenWorksheet(wks)
    
    Dim iPreviousChartNumber As Long
    iPreviousChartNumber = 0
    
    Dim sPreviousSheetName As String
    sPreviousSheetName = vbNullString
    
    Dim i As Long
    For i = 0 To iNoOfEntries - 1
        Dim iActualChartNumber As Long
        iActualChartNumber = rngSeriesData.Offset(i, eSD.ChartNumber - 1).Value2
        
        If iActualChartNumber <> iPreviousChartNumber Then
            iPreviousChartNumber = iActualChartNumber
            
            Dim sDataSheet As String
            sDataSheet = rngSeriesData.Offset(i, eSD.SheetName - 1).Value
            
            Dim sChartName As String
            sChartName = rngSeriesData.Offset(i, eSD.ChartName - 1).Value
            
            If Len(sChartName) > 0 Then
                If sDataSheet <> sPreviousSheetName Then
                    sPreviousSheetName = sDataSheet
                    
                    Dim ChartObjectIndex As Long
                    ChartObjectIndex = 1
                Else
                    ChartObjectIndex = ChartObjectIndex + 1
                End If
            End If
        End If
        
        If Len(sChartName) > 0 Then
            Dim rng As Range
            Set rng = rngSeriesData.Offset(i, eSD.ChartName - 1)
            
            Dim sTopLeftCell As String
            sTopLeftCell = wkb.Worksheets(sDataSheet). _
                    ChartObjects(ChartObjectIndex).TopLeftCell.Address(False, False)
            
            Dim sHyperlinkTarget As String
            sHyperlinkTarget = "'" & sDataSheet & "'!" & sTopLeftCell
            
            AddHyperlinkToCurrentCell _
                    wks, _
                    rng, _
                    sHyperlinkTarget
            
            FormatHyperlinkCell _
                    wks, _
                    rng
        End If
    Next
    
    If Not CurrentChart Is Nothing Then CurrentChart.Activate
    
End Sub


Private Sub AddButtonsThatHyperlinkToCharts( _
    ByVal wks As Worksheet _
)
    
    Call DeleteAllLabelShapesOnSheet(wks)
    
    Dim rngSeriesData As Range
    Set rngSeriesData = wks.Cells(gciTitleRow + 1, 1)
    
    Dim iFirstRow As Long
    iFirstRow = rngSeriesData.Row
    
    Dim iLastRow As Long
    iLastRow = rngSeriesData.Offset(-1, 0).End(xlDown).Row
    
    Dim iNoOfEntries As Long
    iNoOfEntries = iLastRow - iFirstRow + 1
    
    Dim i As Long
    For i = 0 To iNoOfEntries - 1
        Dim sSheetName As String
        sSheetName = rngSeriesData.Offset(i, eSD.SheetName - 1).Value
        
        Dim sChartName As String
        sChartName = rngSeriesData.Offset(i, eSD.ChartName - 1).Value
        
        If Len(sSheetName) > 0 And Len(sChartName) = 0 Then
            Dim rng As Range
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
    
    '==========================================================================
    Const csNamePrefix As String = "lblChart"
    '==========================================================================
    
    With rng
        Dim wks As Worksheet
        Set wks = .Parent
        
        Dim shp As Shape
        Set shp = wks.Shapes.AddLabel( _
                msoTextOrientationHorizontal, _
                .Left, .Top, .Width, .Height _
        )
    End With
    
    With shp
        .Name = csNamePrefix & wks.Cells(rng.Row, 1)
        .Placement = xlMoveAndSize
        .OnAction = "GotoChartWithCellName"
    End With
    
End Sub


'inspired by <https://excel.tips.net/T002539_Hyperlinks_to_Charts.html>
Public Sub GotoChartWithCellName()
    
    Dim rng As Range
    Set rng = ActiveSheet.Shapes(Application.Caller).TopLeftCell
    
    Dim wkb As Workbook
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
    
    Dim sListSeparator As String
    sListSeparator = Application.International(xlListSeparator)
    
    Dim rngSeriesData As Range
    Set rngSeriesData = wks.Cells(gciTitleRow + 1, 1)
    
    Dim iNoOfEntries As Long
    iNoOfEntries = NoOfEntriesInList(rngSeriesData)
    
    Dim arrHyperlinks As Variant
    arrHyperlinks = Array( _
            eSD.SeriesXValues - 1, _
            eSD.SeriesYValues - 1 _
    )
    
    Dim CurrentChart As Chart
    Set CurrentChart = RememberActiveChartAndActivateGivenWorksheet(wks)
    
    Dim i As Long
    For i = 0 To iNoOfEntries - 1
        Dim j As Long
        For j = LBound(arrHyperlinks) To UBound(arrHyperlinks)
            Dim rng As Range
            Set rng = rngSeriesData.Offset(i, arrHyperlinks(j))
            
            AddHyperlinkToCurrentEntry _
                    wks, _
                    rng, _
                    sListSeparator
            
            FormatHyperlinkCell _
                    wks, _
                    rng
        Next
    Next
    
    If Not CurrentChart Is Nothing Then CurrentChart.Activate
    
End Sub


Private Function NoOfEntriesInList( _
    ByVal rng As Range _
        ) As Long
    
    With rng
        Dim iFirstRow As Long
        iFirstRow = .Row
        
        Dim iLastRow As Long
        iLastRow = .Offset(-1, 0).End(xlDown).Row
    End With
    
    NoOfEntriesInList = iLastRow - iFirstRow + 1
    
End Function


Private Sub AddHyperlinkToCurrentEntry( _
    ByVal wks As Worksheet, _
    ByVal rng As Range, _
    ByVal sListSeparator As String _
)
    
    Dim wkb As Workbook
    Set wkb = wks.Parent
    
    Dim sDataBook As String
    sDataBook = rng.Offset(0, -2).Value2
    
    'don't add external hyperlinks
    If Len(sDataBook) > 0 Then
        If sDataBook <> wkb.Name Then
            Exit Sub
        End If
    End If
    
'NOTE: what is this good for?
    Dim sRngValue As String
    sRngValue = rng.Value2
    sRngValue = Replace(sRngValue, ",", sListSeparator)
    
    Dim sDataSheet As String
    sDataSheet = rng.Offset(0, -1).Value2
    
    If Len(sDataSheet) > 0 Then
        Dim wksDataSheet As Worksheet
        Set wksDataSheet = wkb.Worksheets(sDataSheet)
        
        On Error Resume Next
        Dim rngTest As Range
        Set rngTest = wksDataSheet.Range(sRngValue)
        On Error GoTo 0
    'global defined name case
    ElseIf Len(sDataBook) > 0 Then
        Dim NM As Name
        Set NM = wkb.Names(sRngValue)
        
        Set rngTest = NM.RefersToRange
        sDataSheet = rngTest.Parent.Name
    End If
    
    If Not rngTest Is Nothing Then
        Dim sFirstCell As String
        sFirstCell = rngTest.Areas(1).Cells(1).Address(False, False)
        
        Dim sHyperlinkTarget As String
        sHyperlinkTarget = "'" & sDataSheet & "'!" & sFirstCell
        
        Call AddHyperlinkToCurrentCell(wks, rng, sHyperlinkTarget)
    End If
    
End Sub


Private Sub MarkEachOddChartNumberRow( _
    ByVal wks As Worksheet _
)
    
    '==========================================================================
    'color for odd chart numbers
    Const ccOddChartNumbers As Long = 15853276   'R=220 G=230 B=241
    '==========================================================================
    
    Dim rngSeriesData As Range
    Set rngSeriesData = wks.Cells(gciTitleRow + 1, 1)
    
    'find number of entries in list
    With rngSeriesData
        Dim iFirstRow As Long
        iFirstRow = .Row
        
        Dim iLastRow As Long
        iLastRow = .Offset(-1, 0).End(xlDown).Row
        
        Dim iNoOfEntries As Long
        iNoOfEntries = iLastRow - iFirstRow + 1
    End With
    
    With rngSeriesData
        Dim i As Long
        For i = 0 To iNoOfEntries - 1
            If .Offset(i, eSD.ChartNumber - 1).Value Mod 2 = 1 Then
                .Offset(i).EntireRow.Interior.Color = ccOddChartNumbers
            Else
                .Offset(i).EntireRow.Interior.Color = xlColorIndexNone
            End If
        Next
    End With
    
End Sub


Private Sub StuffToBeDoneLast( _
    ByVal wks As Worksheet, _
    ByVal bNewSheet As Boolean _
)
    
    'do it only if the sheet was newly created
    If bNewSheet = True Then
        With wks
            .Rows(gciTitleRow).AutoFilter
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
    
'    wks.UsedRange.WrapText = False
    
End Sub


Private Function CountSCsInAllChartsInWorkbook( _
    ByVal wkb As Workbook _
        ) As Long
    
    'just count the number of entries for the array first for all
    Dim j As Long
    j = 0
    
    With wkb
        'it gets a bit complicated, because charts can occur as
        '- dedicated "Chart" sheets and
        '- as "ChartObjects" on normal WorkSheets
        Dim i As Long
        For i = 1 To .Sheets.Count
            If IsChart(wkb, i) Then
                Dim sc As Series
                '(If you are on Excel 2010 or older, you will get a compile error.
                ' Replace `FullSeriesCollection` by `SeriesCollection` in this
                ' module.)
                For Each sc In .Sheets(i).FullSeriesCollection
                    j = j + 1
                Next
            Else
                Dim crt As ChartObject
                For Each crt In .Sheets(i).ChartObjects
                    For Each sc In crt.Chart.FullSeriesCollection
                        j = j + 1
                    Next
                Next
            End If
        Next
    End With
    
    CountSCsInAllChartsInWorkbook = j
    
End Function


Private Function CountChartsWithoutSeriesInWorkbook( _
    ByVal wkb As Workbook _
        ) As Long
    
    'just count the number of entries for the array first for all
    Dim j As Long
    j = 0
    
    With wkb
        'it gets a bit complicated, because charts can occur as
        '- dedicated "Chart" sheets and
        '- as "ChartObjects" on normal WorkSheets
        Dim i As Long
        For i = 1 To .Sheets.Count
            If IsChart(wkb, i) Then
                If .Sheets(i).FullSeriesCollection.Count = 0 Then
                    j = j + 1
                End If
            Else
                Dim crt As ChartObject
                For Each crt In .Sheets(i).ChartObjects
                    If crt.Chart.FullSeriesCollection.Count = 0 Then
                        j = j + 1
                    End If
                Next
            End If
        Next
    End With
    
    CountChartsWithoutSeriesInWorkbook = j
    
End Function


Private Sub HideSeriesPathBookIfEmpty( _
    ByVal wksSeriesLegend As Worksheet, _
    ByVal arrData As Variant _
)
    
    Dim NoOfRows As Long
    NoOfRows = UBound(arrData) - LBound(arrData) + 1
    
    Dim NoOfColumns As Long
    NoOfColumns = eSD.SeriesNameBook - eSD.SeriesNamePath + 1
    
    Dim rng As Range
    Set rng = wksSeriesLegend.Cells(gciTitleRow, 1).Offset(1).Resize(NoOfRows)
    
    Dim rng2 As Range
    Set rng2 = rng.Offset(, eSD.SeriesNamePath - 1).Resize(, NoOfColumns)
    If Application.WorksheetFunction.CountA(rng2) = 0 Then rng2.Columns.Hidden = True
    
    Set rng2 = rng.Offset(, eSD.SeriesXPath - 1).Resize(, NoOfColumns)
    If Application.WorksheetFunction.CountA(rng2) = 0 Then rng2.Columns.Hidden = True
    
    Set rng2 = rng.Offset(, eSD.SeriesYPath - 1).Resize(, NoOfColumns)
    If Application.WorksheetFunction.CountA(rng2) = 0 Then rng2.Columns.Hidden = True
    
End Sub


Private Sub CreateAndInitializeSeriesEntriesInChartsWorksheet( _
    ByVal wkb As Workbook _
)
    
    Dim arrHeading As Variant
    arrHeading = TransferHeadingNamesToArray
    
    Dim wks As Worksheet
    Set wks = wkb.Worksheets.Add(wkb.Worksheets(1))
    
    '"configure" the new sheet
    With wks
        .Name = pcsLegendSheetName
        .Tab.ThemeColor = xlThemeColorLight1
        .Tab.TintAndShade = 0
        
        'set column titles
        .Cells(gciTitleRow, 1).Resize(, UBound(arrHeading) + 1) = arrHeading
        
        'AxisGroup
        With .Cells(gciTitleRow, eSD.AxisGroup)
            .AddComment ( _
                    "AxisGroup" & Chr$(10) & _
                    "1 = Primary Axis" & Chr$(10) & _
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
        'first level
        .Columns(eSD.ChartName).Group
        .Columns(eSD.Y2Label).Group
        .Columns(eSD.SeriesNamePath).Group
        .Columns(eSD.SeriesNameBook).Group
        .Columns(eSD.SeriesNameSheet).Group
        .Columns(eSD.SeriesNameRange).Group
        .Columns(eSD.SeriesXPath).Group
        .Columns(eSD.SeriesXBook).Group
        .Columns(eSD.SeriesXSheet).Group
        .Columns(eSD.SeriesYPath).Group
        .Columns(eSD.SeriesYBook).Group
        .Columns(eSD.SeriesYSheet).Group
        'second level
        .Columns(eSD.SeriesNamePath).Group
        .Columns(eSD.SeriesNameBook).Group
        .Columns(eSD.SeriesXPath).Group
        .Columns(eSD.SeriesXBook).Group
        .Columns(eSD.SeriesYPath).Group
        .Columns(eSD.SeriesYBook).Group
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
            .LeftHeader = "&Z&F" & Chr$(10) & "&A"
            .RightHeader = vbNullString & Chr$(10) & "&D"
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
            "Series.Name.Path" & csStringSep & _
            "Series.Name.Book" & csStringSep & _
            "Series.Name.Sheet" & csStringSep & _
            "Series.Name.Range" & csStringSep & _
            "Series.Name" & csStringSep & _
            "Series.X.Path" & csStringSep & _
            "Series.X.Book" & csStringSep & _
            "Series.X.Sheet" & csStringSep & _
            "Series.X.Values" & csStringSep & _
            "Series.Y.Path" & csStringSep & _
            "Series.Y.Book" & csStringSep & _
            "Series.Y.Sheet" & csStringSep & _
            "Series.Y.Values" & csStringSep & _
            "AG" & csStringSep & _
            "PO" & csStringSep & _
            "POt"
    '==========================================================================
    
    Dim arrHeading As Variant
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
    
    DeleteAllHyperlinks rng
    
    wks.Hyperlinks.Add _
            Anchor:=rng, _
            Address:=vbNullString, _
            SubAddress:=sHyperlinkTarget, _
            ScreenTip:=sHyperlinkTarget
    
End Sub


Private Sub DeleteAllHyperlinks( _
    ByVal rng As Range _
)
    With rng
        Do Until .Hyperlinks.Count = 0
            .Hyperlinks(1).Delete
        Loop
    End With
End Sub


Private Sub FormatHyperlinkCell( _
    ByVal wks As Worksheet, _
    ByVal rng As Range _
)
    With rng.Font
        .ColorIndex = xlColorIndexAutomatic
        .Underline = xlUnderlineStyleNone
        .Size = wks.Cells(gciTitleRow + 1, 1).Font.Size
    End With
End Sub


Private Sub FillArrayWithSCData( _
    ByVal wkb As Workbook, _
    ByVal cha As Chart, _
    ByRef arrData As Variant, _
    ByRef iSheetIndex As Long, _
    ByRef iSCTotal As Long, _
    ByRef iChartNumber As Long _
)
    
    With wkb
        arrData(iSCTotal, eSD.SheetName) = .Sheets(iSheetIndex).Name
        If Not .Name = cha.Parent.Name Then
            arrData(iSCTotal, eSD.ChartName) = cha.Parent.Name
        End If
    End With
    
    With cha
        If .HasTitle Then arrData(iSCTotal, eSD.ChartTitle) = .ChartTitle.Text
        If .Axes(xlCategory).HasTitle Then _
                arrData(iSCTotal, eSD.XLabel) = .Axes(xlCategory).AxisTitle.Text
        If .Axes(xlValue).HasTitle Then _
                arrData(iSCTotal, eSD.YLabel) = .Axes(xlValue, xlPrimary).AxisTitle.Text
        If .HasAxis(xlValue, xlSecondary) Then
            If .Axes(xlValue, xlSecondary).HasTitle Then
                arrData(iSCTotal, eSD.Y2Label) = .Axes(xlValue, xlSecondary).AxisTitle.Text
            End If
        End If
        
        Dim iSeriesCount As Long
        iSeriesCount = .FullSeriesCollection.Count
        
        If iSeriesCount = 0 Then
            arrData(iSCTotal, eSD.ChartNumber) = iChartNumber
            iSCTotal = iSCTotal + 1
        Else
            Dim iSC As Long
            For iSC = 1 To iSeriesCount
                Dim srs As Series
                Set srs = .FullSeriesCollection(iSC)
                
                arrData(iSCTotal, eSD.ChartNumber) = iChartNumber
                arrData(iSCTotal, eSD.SeriesNameValue) = srs.Name
                
                Dim MySeries As IChartSeries
                Set MySeries = ChartSeries.Create(srs)
                
                FillArrayWithSCDataCurrentSeries _
                        wkb, _
                        iSCTotal, _
                        MySeries, _
                        arrData
                
                arrData(iSCTotal, eSD.AxisGroup) = srs.AxisGroup
                arrData(iSCTotal, eSD.PlotOrder) = srs.PlotOrder
                
                'if 'srs' is inaccessible
                If Len(arrData(iSCTotal, eSD.PlotOrderTotal)) = 0 Then
                    arrData(iSCTotal, eSD.PlotOrderTotal) = iSC
                End If
                
                iSCTotal = iSCTotal + 1
            Next
        End If
    End With
End Sub


Private Sub FillArrayWithSCDataCurrentSeries( _
    ByVal wkb As Workbook, _
    ByVal iSCTotal As Long, _
    ByVal MySeries As IChartSeries, _
    ByRef arrData As Variant _
)
    
    With MySeries
'        arrData(iSCTotal, eSD.SeriesNameValue) = .SeriesName
        If .IsSeriesAccessible Then
            With .SeriesName
                If .IsRange Then
                    If Len(.RangePath) > 0 Then
                        arrData(iSCTotal, eSD.SeriesNamePath) = .RangePath
                        arrData(iSCTotal, eSD.SeriesNameBook) = .RangeBook
                    ElseIf .RangeBook <> wkb.Name Or Len(.RangeSheet) = 0 Then
                        arrData(iSCTotal, eSD.SeriesNameBook) = .RangeBook
                    End If
                    arrData(iSCTotal, eSD.SeriesNameSheet) = .RangeSheet
                    arrData(iSCTotal, eSD.SeriesNameRange) = .RangeString
                End If
            End With
            
            With .XValues
                If Not .IsRange Then
                    arrData(iSCTotal, eSD.SeriesXValues) = .FormulaPart
                Else
                    If Len(.RangePath) > 0 Then
                        arrData(iSCTotal, eSD.SeriesXPath) = .RangePath
                        arrData(iSCTotal, eSD.SeriesXBook) = .RangeBook
                    ElseIf .RangeBook <> wkb.Name Or Len(.RangeSheet) = 0 Then
                        arrData(iSCTotal, eSD.SeriesXBook) = .RangeBook
                    End If
                    arrData(iSCTotal, eSD.SeriesXSheet) = .RangeSheet
                    arrData(iSCTotal, eSD.SeriesXValues) = .RangeString
                End If
            End With
            
            With .Values
                If Not .IsRange Then
                    arrData(iSCTotal, eSD.SeriesYValues) = .FormulaPart
                Else
                    If Len(.RangePath) > 0 Then
                        arrData(iSCTotal, eSD.SeriesYPath) = .RangePath
                        arrData(iSCTotal, eSD.SeriesYBook) = .RangeBook
                    ElseIf .RangeBook <> wkb.Name Or Len(.RangeSheet) = 0 Then
                        arrData(iSCTotal, eSD.SeriesYBook) = .RangeBook
                    End If
                    arrData(iSCTotal, eSD.SeriesYSheet) = .RangeSheet
                    arrData(iSCTotal, eSD.SeriesYValues) = .RangeString
                End If
            End With
            
            arrData(iSCTotal, eSD.PlotOrderTotal) = .PlotOrder.Value
        Else
            arrData(iSCTotal, eSD.SeriesXValues) = pcsInaccessible
            arrData(iSCTotal, eSD.SeriesYValues) = pcsInaccessible
        End If
    End With
    
End Sub


'==============================================================================
Private Sub MarkSheetEntriesIfSourceIsInvisible( _
    ByVal wkb As Workbook, _
    ByVal arrData As Variant _
)
    
    Dim arrInvisibleSheets As Variant
    arrInvisibleSheets = GetInvisibleSheetsArray(wkb)
    
    If IsEmpty(arrInvisibleSheets) Then Exit Sub
    
    Dim wksSeriesLegend As Worksheet
    Set wksSeriesLegend = wkb.Worksheets(pcsLegendSheetName)
    
    Dim rng As Range
    Set rng = wksSeriesLegend.Cells(gciTitleRow, 1)
    
    Dim iRow As Long
    For iRow = LBound(arrData) To UBound(arrData)
        MarkSheetEntriesInCurrentRowIfSourceIsInvisible _
                wkb, _
                arrData, _
                iRow, _
                arrInvisibleSheets, _
                rng
    Next
    
End Sub


Private Function GetInvisibleSheetsArray( _
    ByVal wkb As Workbook _
        ) As Variant
    
    Dim wksHiddenSheet As Worksheet
    Set wksHiddenSheet = wkb.Worksheets(gcsHiddenSheetName)
    
    Dim rngInvisibleSheets As Range
    Set rngInvisibleSheets = wksHiddenSheet.Range(gcsInvisibleSheetsRange)
    
    If Len(rngInvisibleSheets.Value2) = 0 Then Exit Function
    
    GetInvisibleSheetsArray = rngInvisibleSheets.CurrentRegion.Value2
    
End Function


Private Sub MarkSheetEntriesInCurrentRowIfSourceIsInvisible( _
    ByVal wkb As Workbook, _
    ByVal arrData As Variant, _
    ByVal iRow As Long, _
    ByVal arrInvisibleSheets As Variant, _
    ByVal rng As Range _
)
    
    '==========================================================================
    'which cells should be marked/tested
    Dim arrInvisibleItems As Variant
    arrInvisibleItems = Array( _
            Array( _
                eSD.SheetName _
            ), _
            Array( _
                eSD.SeriesXSheet, _
                eSD.SeriesXValues _
            ), _
            Array( _
                eSD.SeriesYSheet, _
                eSD.SeriesYValues _
            ) _
    )
    '==========================================================================
    
    If IsInFirstColOf2DArray( _
            arrData(iRow, arrInvisibleItems(1)(1)), _
            arrInvisibleSheets _
    ) Then
        MarkSheetEntriesAsInvisible _
                iRow, _
                rng, _
                arrInvisibleItems(1)
    End If
    
    Dim i As Long
    For i = LBound(arrInvisibleItems) + 1 To UBound(arrInvisibleItems)
        'only do this for ranges on 'wkb'
        If IsRangeInWkb(wkb, arrData, iRow, arrInvisibleItems(i)(1)) Then
            If IsInFirstColOf2DArray( _
                    arrData(iRow, arrInvisibleItems(i)(1)), _
                    arrInvisibleSheets _
            ) Then
                MarkSheetEntriesAsInvisible _
                        iRow, _
                        rng, _
                        arrInvisibleItems(i)
            End If
        End If
    Next
    
End Sub


Private Function IsInFirstColOf2DArray( _
    ByVal sString As String, _
    ByVal vArr As Variant _
        ) As Boolean
    
    IsInFirstColOf2DArray = False
    
    Dim i As Long
    For i = LBound(vArr) To UBound(vArr)
        If vArr(i, 1) = sString Then
            IsInFirstColOf2DArray = True
            Exit Function
        End If
    Next
    
End Function


Private Function IsRangeInWkb( _
    ByVal wkb As Workbook, _
    ByVal arrData As Variant, _
    ByVal iRow As Long, _
    ByVal eSheet As eSD _
        ) As Boolean
    
    IsRangeInWkb = False
    
    If Len(arrData(iRow, eSheet - 2)) > 0 Then Exit Function
    If Len(arrData(iRow, eSheet - 1)) > 0 Then
        If arrData(iRow, eSheet - 1) <> wkb.Name Then Exit Function
    Else
        If Len(arrData(iRow, eSheet)) = 0 Then Exit Function
    End If
    
    IsRangeInWkb = True
    
End Function


Private Sub MarkSheetEntriesAsInvisible( _
    ByVal iRow As Long, _
    ByVal rng As Range, _
    ByRef arrToMarkCells As Variant _
)
    
    '==========================================================================
    'color for stuff that is hidden/invisible (fully)
    Const ccHidden As Long = 12040422            'R=230 G=184 B=183
    '==========================================================================
    
    Dim j As Long
    For j = LBound(arrToMarkCells) To UBound(arrToMarkCells)
        Dim iCol As Long
        iCol = arrToMarkCells(j)
        
        rng.Offset(iRow, iCol - 1).Interior.Color = ccHidden
    Next
    
End Sub


Private Sub MarkSeriesDataIfSourceIsHidden( _
    ByVal wkb As Workbook, _
    ByVal arrData As Variant _
)
    
    '==========================================================================
    'color for stuff that is hidden/invisible (fully)
    Const ccHidden As Long = 12040422            'R=230 G=184 B=183
    'color for stuff that is hidden partially
    Const ccHiddenPartly As Long = 14408946      'R=242 G=220 B=219
    'which cells should be marked/tested
    Dim arrToMarkCells As Variant
    arrToMarkCells = Array( _
            eSD.SeriesXValues, _
            eSD.SeriesYValues _
    )
    '==========================================================================
    
    Dim arrHiddenRanges As Variant
    arrHiddenRanges = GetInvisibleRangesArray(wkb)
    
    If IsEmpty(arrHiddenRanges) Then Exit Sub
    
    Dim wksSeriesLegend As Worksheet
    Set wksSeriesLegend = wkb.Worksheets(pcsLegendSheetName)
    
    Dim rng As Range
    Set rng = wksSeriesLegend.Cells(gciTitleRow, 1)
    
    Dim iRow As Long
    For iRow = LBound(arrData) To UBound(arrData)
        Dim j As Long
        For j = LBound(arrToMarkCells) To UBound(arrToMarkCells)
            Dim iCol As Long
            iCol = arrToMarkCells(j)
            
            If rng.Offset(iRow, iCol - 1).Interior.Color <> ccHidden Then
                If IsRangeInWkb(wkb, arrData, iRow, iCol - 1) Then
                    Select Case IsRangeHidden( _
                            wkb, _
                            arrData(iRow, iCol - 1), _
                            arrData(iRow, iCol), _
                            arrHiddenRanges _
                    )
                        Case 1
                            rng.Offset(iRow, iCol - 1).Interior.Color = ccHiddenPartly
                        Case 2
                            rng.Offset(iRow, iCol - 1).Interior.Color = ccHidden
                    End Select
                End If
            End If
        Next
    Next
    
End Sub


Private Function GetInvisibleRangesArray( _
    ByVal wkb As Workbook _
        ) As Variant
    
    Dim wksHiddenSheet As Worksheet
    Set wksHiddenSheet = wkb.Worksheets(gcsHiddenSheetName)
    
    Dim rngHiddenRanges As Range
    Set rngHiddenRanges = wksHiddenSheet.Range(gcsHiddenRangesRange)
    
    If Len(rngHiddenRanges.Value2) = 0 Then Exit Function
    
    GetInvisibleRangesArray = rngHiddenRanges.CurrentRegion.Value2
    
End Function


Private Function IsRangeHidden( _
    ByVal wkb As Workbook, _
    ByVal sWorksheet As String, _
    ByVal sRange As String, _
    ByVal arrHiddenRanges As Variant _
        ) As Long
    
    If Len(sRange) = 0 Then Exit Function
    
    'handler for global defined name range
    If Len(sWorksheet) = 0 Then
        Dim sWks As String
        sWks = GetWorksheetOfGlobalDefinedNameRange(wkb, sRange)
    Else
        sWks = sWorksheet
    End If
    
    'test if there is a hidden range on 'sWks' and if yes,
    'create an array of hidden areas ('arrHidden')
    Dim i As Long
    For i = LBound(arrHiddenRanges) To UBound(arrHiddenRanges)
        If sWks = arrHiddenRanges(i, 1) Then
            Dim arrHidden As Variant
            arrHidden = CreateArrayOfHiddenAreas( _
                    arrHiddenRanges(i, 2), _
                    arrHiddenRanges(i, 3) _
            )
            
            Dim bWksFound As Boolean
            bWksFound = True
            
            Exit For
        End If
    Next
    
    If Not bWksFound Then
        IsRangeHidden = 0
        Exit Function
    End If
    
    Dim arrToTest As Variant
    arrToTest = Split(sRange, ",")
    
    Dim iHidden() As Long
    ReDim iHidden(LBound(arrToTest) To UBound(arrToTest))
    
    For i = LBound(arrToTest) To UBound(arrToTest)
        Dim rngToTest As Range
        Set rngToTest = wkb.Worksheets(sWks).Range(arrToTest(i))
        
        iHidden(i) = IsAreaHidden(rngToTest, arrHidden)
    Next
    
    IsRangeHidden = ReturnHiddenState(iHidden)
    
End Function


Private Function GetWorksheetOfGlobalDefinedNameRange( _
    ByVal wkb As Workbook, _
    ByVal sRange As String _
        ) As String
    
    Dim NM As Name
    Set NM = wkb.Names(sRange)
    
    Dim rng As Range
    Set rng = NM.RefersToRange
    
    Dim wks As Worksheet
    Set wks = rng.Parent
    
    GetWorksheetOfGlobalDefinedNameRange = wks.Name
    
End Function


Private Function CreateArrayOfHiddenAreas( _
    ByVal sRows As String, _
    ByVal sColumns As String _
        ) As Variant
    
    Const csSep As String = ","
    
    If Len(sRows) > 0 And Len(sColumns) > 0 Then
        Dim sRange As String
        sRange = sRows & csSep & sColumns
    ElseIf Len(sRows) > 0 Then
        sRange = sRows
    Else
        sRange = sColumns
    End If
    
    Dim Arr As Variant
    Arr = Split(sRange, csSep)
    
    CreateArrayOfHiddenAreas = Arr
    
End Function


Private Function IsAreaHidden( _
    ByVal rngToTest As Range, _
    ByVal Arr As Variant _
        ) As Long
    
    Dim wks As Worksheet
    Set wks = rngToTest.Parent
    
    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)
        Dim rng As Range
        Set rng = wks.Range(Arr(i))
        
        Dim rngIntersect As Range
        Set rngIntersect = Application.Intersect(rngToTest, rng)
        
        If rngIntersect Is Nothing Then
        Dim iHidden As Long
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


Private Sub MarkSeriesSheetIfXYValuesAreOnDifferentSheets( _
    ByVal wksSeriesLegend As Worksheet, _
    ByVal arrData As Variant _
)
    
    '==========================================================================
    'color for stuff that is "wrong"
    Const ccNotEqual As Long = 255            'R=255 G=0 B=0
    '==========================================================================
    
    Dim rng As Range
    Set rng = wksSeriesLegend.Cells(gciTitleRow, 1)
    
    Dim i As Long
    For i = LBound(arrData) To UBound(arrData)
        If Not IsEmpty(arrData(i, eSD.SeriesXSheet)) Then
            If Not IsEmpty(arrData(i, eSD.SeriesYSheet)) Then
                If arrData(i, eSD.SeriesXSheet) <> arrData(i, eSD.SeriesYSheet) Then
                    rng.Offset(i, eSD.SeriesXSheet - 1).Font.Color = ccNotEqual
                    rng.Offset(i, eSD.SeriesYSheet - 1).Font.Color = ccNotEqual
                End If
            End If
        End If
    Next
    
End Sub


Private Sub MarkYValuesIfRowsOrColsDoNotCorrespond( _
    ByVal wksSeriesLegend As Worksheet, _
    ByVal arrData As Variant _
)
    
    '==========================================================================
    'color for stuff that is "wrong"
    Const ccWrong As Long = 255            'R=255 G=0 B=0
    '==========================================================================
    
    Dim rng As Range
    Set rng = wksSeriesLegend.Cells(gciTitleRow, 1)
    
    Dim i As Long
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
    
    'if one of the ranges is not a range, no test is needed
    If Left$(sXRange, 1) = "{" Or Left$(sYRange, 1) = "{" Then
        AreRowsOrColsEqual = True
        Exit Function
    End If
    
    Dim bAreRowsEqual As Boolean
    bAreRowsEqual = (ExtractRowsRange(sXRange) = ExtractRowsRange(sYRange))
    
    Dim bAreColsEqual As Boolean
    bAreColsEqual = (ExtractColumnsRange(sXRange) = ExtractColumnsRange(sYRange))
    
    AreRowsOrColsEqual = (bAreRowsEqual Or bAreColsEqual)
    
End Function
