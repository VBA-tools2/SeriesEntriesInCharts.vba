Attribute VB_Name = "modSeriesFormatting"

'@Folder("ChartSeries.Extensions")

Option Explicit
Option Private Module


Private Enum eMS
    [_First] = 1
    Blank = eMS.[_First]
    Style
    ForegroundColor
    BackgroundColor
    [_Last] = eMS.BackgroundColor
End Enum


Private Type tAxisSettings
    srs As Series
    cha As Chart
    axGroup As XlAxisGroup
    
    IsMinScaleAuto As Boolean
    IsMaxScaleAuto As Boolean
    MinScale As Double
    MaxScale As Double
    
    NumberFormat As String
    NumberFormatLinked As Boolean
End Type


Private Const MarkerStyleNone As String = "none"


Public Function AddSeriesMarkerFormatting( _
    ByVal wksSeriesLegend As Worksheet _
        ) As Boolean
    
    'Set the default return value
    AddSeriesMarkerFormatting = False
    
    Dim iLastCol As Long
    iLastCol = ReturnLastColumnInRow( _
            wksSeriesLegend.Cells(gciTitleRow, 1), _
            IncludingHiddenCells:=True _
    )
    
    Dim LastHeadingCell As Range
    Set LastHeadingCell = wksSeriesLegend.Cells(gciTitleRow, iLastCol)
    
    PrepareSheetForMarkerData LastHeadingCell
    
    Dim iCorrectedLastCol As Long
    iCorrectedLastCol = LastHeadingCell.Column
    
    Dim wkb As Workbook
    Set wkb = wksSeriesLegend.Parent
    
    'fill the array
    Dim iSCTotal As Long
    iSCTotal = 1
    
    Dim iSheetIndex As Long
    For iSheetIndex = 1 To wkb.Sheets.Count
        If IsChart(wkb, iSheetIndex) Then
            Dim cha As Chart
            Set cha = wkb.Sheets(iSheetIndex)
            
            ShowSeriesFormattingCurrentChart _
                    wksSeriesLegend, _
                    cha, _
                    iCorrectedLastCol, _
                    iSCTotal
        Else
            Dim iChartObjectIndex As Long
            For iChartObjectIndex = 1 To wkb.Sheets(iSheetIndex).ChartObjects.Count
                Dim crt As ChartObject
                Set crt = wkb.Sheets(iSheetIndex).ChartObjects(iChartObjectIndex)
                Set cha = crt.Chart
                
                ShowSeriesFormattingCurrentChart _
                        wksSeriesLegend, _
                        cha, _
                        iCorrectedLastCol, _
                        iSCTotal
            Next
        End If
    Next
    
    AddSeriesMarkerFormatting = True
    
End Function


Private Sub PrepareSheetForMarkerData( _
    ByRef outLastHeadingCell As Range _
)
    
    '==========================================================================
    Const BackgroundColorHeadingTitle As String = "BC"
    '==========================================================================
    
    With outLastHeadingCell
        If .Value2 = BackgroundColorHeadingTitle Then
            Set outLastHeadingCell = .Offset(, eMS.[_First] - eMS.[_Last] - 1)
        Else
            With .Offset(, eMS.Style)
                .Value2 = "M"
                .AddComment ("Marker Style(/Symbol)")
                .Comment.Shape.TextFrame.AutoSize = True
            End With
            With .Offset(, eMS.ForegroundColor)
                .Value2 = "FC"
                .AddComment ("Foreground Color")
                .Comment.Shape.TextFrame.AutoSize = True
            End With
            With .Offset(, eMS.BackgroundColor)
                .Value2 = BackgroundColorHeadingTitle
                .AddComment ("BackgroundColor")
                .Comment.Shape.TextFrame.AutoSize = True
            End With
            
            .Offset(, eMS.[_First]).Resize(, eMS.[_Last] - eMS.[_First] + 1).ColumnWidth = 2.71
        End If
    End With
    
    outLastHeadingCell.Offset(, eMS.[_First]).EntireColumn.Interior.ColorIndex = xlColorIndexNone
    
End Sub


Private Sub ShowSeriesFormattingCurrentChart( _
    ByVal wksSeriesLegend As Worksheet, _
    ByVal cha As Chart, _
    ByVal iLastCol As Long, _
    ByRef outSCTotal As Long _
)
    
    Dim iSeriesCount As Long
    iSeriesCount = cha.FullSeriesCollection.Count
    
    If iSeriesCount = 0 Then
        outSCTotal = outSCTotal + 1
        Exit Sub
    End If
    
    Dim iSC As Long
    For iSC = 1 To iSeriesCount
        Dim srs As Series
        Set srs = cha.FullSeriesCollection(iSC)
        
        If CanSeriesHaveMarkers(srs) Then
            Dim MarkerStyle As String
            MarkerStyle = GetMarkerStyle(srs)
            
            If MarkerStyle <> MarkerStyleNone Then
                Dim MarkerForegroundColor As Long
                MarkerForegroundColor = GetMarkerForegroundColor(srs)
                
                Dim MarkerBackgroundColor As Long
                MarkerBackgroundColor = GetMarkerBackgroundColor(srs)
            End If
            
            WriteMarkerDataToSheet _
                    wksSeriesLegend.Cells(gciTitleRow + outSCTotal, iLastCol), _
                    MarkerStyle, _
                    MarkerForegroundColor, _
                    MarkerBackgroundColor, _
                    srs
        End If
        
        outSCTotal = outSCTotal + 1
    Next
    
End Sub


Private Function CanSeriesHaveMarkers( _
    ByVal srs As Series _
        ) As Boolean
    
    CanSeriesHaveMarkers = False
    
    Select Case srs.ChartType
        Case xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers, _
                xlXYScatterSmooth, xlXYScatterSmoothNoMarkers
        Case xlLine, xlLineMarkers
        Case xlLineMarkersStacked, xlLineMarkersStacked100, _
                xlLineStacked, xlLineStacked100
        Case xlRadar, xlRadarFilled, xlRadarMarkers
        Case Else
            Exit Function
    End Select
    
    CanSeriesHaveMarkers = True
    
End Function


Private Function GetMarkerStyle( _
    ByVal srs As Series _
        ) As String
    
    With srs
        Select Case .MarkerStyle
            Case xlMarkerStyleAutomatic
                GetMarkerStyle = "auto"
            Case xlMarkerStyleCircle
                GetMarkerStyle = ChrW$(9679)
            Case xlMarkerStyleDash
                GetMarkerStyle = ChrW$(9644)
            Case xlMarkerStyleDiamond
                GetMarkerStyle = ChrW$(9670)
            Case xlMarkerStyleDot
                GetMarkerStyle = "dot"
            Case xlMarkerStyleNone
                GetMarkerStyle = MarkerStyleNone
            Case xlMarkerStylePicture
                GetMarkerStyle = "picture"
            Case xlMarkerStylePlus
                GetMarkerStyle = "+"
            Case xlMarkerStyleSquare
                GetMarkerStyle = ChrW$(9632)
            Case xlMarkerStyleStar
                GetMarkerStyle = ChrW$(9733)
            Case xlMarkerStyleTriangle
                GetMarkerStyle = ChrW$(9650)
            Case xlMarkerStyleX
                GetMarkerStyle = ChrW$(215)
        End Select
    End With
    
End Function


Private Function GetMarkerForegroundColor( _
    ByVal srs As Series _
        ) As Long
    
    With srs
        If .MarkerForegroundColor >= 0 Then
            GetMarkerForegroundColor = .MarkerForegroundColor
        ElseIf .MarkerForegroundColorIndex = xlColorIndexNone Then
            GetMarkerForegroundColor = xlColorIndexNone
        'automatic color
        'REF: <https://stackoverflow.com/a/25826428>
        Else
            'check needed, otherwise resetting 'AxisGroup' will fail
            If Not IsSeriesFormulaAccessible(srs) Then
                GetMarkerForegroundColor = -1
            Else
                Dim chtType As XlChartType
                chtType = .ChartType
                
                Dim AxisSettings As tAxisSettings
                AxisSettings = StoreXAxisSettings(srs)
                
                .ChartType = xlColumnClustered
                GetMarkerForegroundColor = .Format.Fill.ForeColor.RGB
                
                .ChartType = chtType
                RestoreXAxisSettings AxisSettings
            End If
        End If
    End With
    
End Function


Private Function StoreXAxisSettings( _
    ByVal srs As Series _
        ) As tAxisSettings
    
    Dim cha As Chart
    Set cha = srs.Parent.Parent
    
    Dim AxisSettings As tAxisSettings
    Set AxisSettings.srs = srs
    Set AxisSettings.cha = cha
    
    Dim ax As Axis
    Set ax = cha.Axes(xlCategory, srs.AxisGroup)
    
    With ax
        AxisSettings.axGroup = srs.AxisGroup
        AxisSettings.IsMinScaleAuto = .MinimumScaleIsAuto
        AxisSettings.IsMaxScaleAuto = .MaximumScaleIsAuto
            
        If Not AxisSettings.IsMinScaleAuto Then
            AxisSettings.MinScale = .MinimumScale
        End If
        If Not AxisSettings.IsMaxScaleAuto Then
            AxisSettings.MaxScale = .MaximumScale
        End If
        
        'sometimes the `.TickLabels` property doesn't exist although the axis exists
        'this most likely is a bug ...
        On Error GoTo errNoTickLabelsOnSecondaryAxis
        AxisSettings.NumberFormat = .TickLabels.NumberFormat
        AxisSettings.NumberFormatLinked = .TickLabels.NumberFormatLinked
        On Error GoTo 0
    End With
    
    If Err.Number <> 0 Then
errNoTickLabelsOnSecondaryAxis:
        Set ax = cha.Axes(xlCategory, xlPrimary)
        With ax
            AxisSettings.NumberFormat = .TickLabels.NumberFormat
            AxisSettings.NumberFormatLinked = .TickLabels.NumberFormatLinked
        End With
    End If
    
    StoreXAxisSettings = AxisSettings
    
End Function


Private Sub RestoreXAxisSettings( _
    ByRef AxisSettings As tAxisSettings _
)
    
    AxisSettings.srs.AxisGroup = AxisSettings.axGroup
    
    Dim ax As Axis
    Set ax = AxisSettings.cha.Axes(xlCategory, AxisSettings.axGroup)
    
    With ax
        .MinimumScaleIsAuto = AxisSettings.IsMinScaleAuto
        .MaximumScaleIsAuto = AxisSettings.IsMaxScaleAuto
        
        If Not AxisSettings.IsMinScaleAuto Then
            .MinimumScale = AxisSettings.MinScale
        End If
        If Not AxisSettings.IsMaxScaleAuto Then
            .MaximumScale = AxisSettings.MaxScale
        End If
        
        'sometimes the `.TickLabels` property doesn't exist although the axis exists
        'this most likely is a bug ...
        On Error GoTo errNoTickLabelsOnSecondaryAxis
        If .TickLabels.NumberFormat <> AxisSettings.NumberFormat Then
            .TickLabels.NumberFormat = AxisSettings.NumberFormat
        End If
        If .TickLabels.NumberFormatLinked <> AxisSettings.NumberFormatLinked Then
            .TickLabels.NumberFormatLinked = AxisSettings.NumberFormatLinked
        End If
    End With
    
    If Err.Number <> 0 Then
errNoTickLabelsOnSecondaryAxis:
        Set ax = AxisSettings.cha.Axes(xlCategory, xlPrimary)
        With ax
            If .TickLabels.NumberFormat <> AxisSettings.NumberFormat Then
                .TickLabels.NumberFormat = AxisSettings.NumberFormat
            End If
            If .TickLabels.NumberFormatLinked <> AxisSettings.NumberFormatLinked Then
                .TickLabels.NumberFormatLinked = AxisSettings.NumberFormatLinked
            End If
       End With
    End If
    
End Sub


Private Function IsSeriesFormulaAccessible( _
    ByVal srs As Series _
        ) As Boolean
    
    On Error Resume Next
    Dim srsFormula As String
    srsFormula = srs.Formula
    On Error GoTo 0
    
    IsSeriesFormulaAccessible = (Len(srsFormula) > 0)
    
End Function


Private Function GetMarkerBackgroundColor( _
    ByVal srs As Series _
        ) As Long
    
    With srs
        If .MarkerBackgroundColorIndex = xlColorIndexNone Then
            GetMarkerBackgroundColor = xlColorIndexNone
        Else
            GetMarkerBackgroundColor = .MarkerBackgroundColor
        End If
    End With
    
End Function


Private Sub WriteMarkerDataToSheet( _
    ByVal rng As Range, _
    ByVal MarkerStyle As String, _
    ByVal MarkerForegroundColor As Long, _
    ByVal MarkerBackgroundColor As Long, _
    ByVal srs As Series _
)
    
    With rng
        .Offset(, eMS.Style).Value2 = MarkerStyle
        
        If MarkerStyle = MarkerStyleNone Then Exit Sub
        
        With .Offset(, eMS.ForegroundColor)
            Select Case MarkerForegroundColor
                Case Is >= 0
                    .Interior.Color = MarkerForegroundColor
                Case xlColorIndexAutomatic, -1
                    .Value2 = "auto"
                Case xlColorIndexNone
                    .Value2 = "none"
            End Select
        End With
        
        With .Offset(, eMS.BackgroundColor)
            Select Case MarkerBackgroundColor
                Case Is >= 0
                    .Interior.Color = MarkerBackgroundColor
                Case xlColorIndexAutomatic, -1
                    .Value2 = "auto"
                Case xlColorIndexNone
                    .Value2 = "none"
            End Select
        End With
    End With
    
End Sub
