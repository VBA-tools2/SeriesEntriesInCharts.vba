Attribute VB_Name = "modExtensions"

'@Folder("ChartSeries.Extensions")

Option Explicit
Option Private Module

Public Sub ApplyExtensions( _
    ByVal wksSeriesLegend As Worksheet, _
    ByVal bNewSeriesSheet As Boolean, _
    ByVal arrData As Variant _
)
    
    AddSeriesMarkerFormatting wksSeriesLegend
    
End Sub
