Attribute VB_Name = "modUsefulFunctions"

Option Explicit
Option Private Module
Option Base 1


'test, if a (normal or named) range exists
Public Function RangeExists( _
    ByVal wkb As Workbook, _
    ByVal RangeName As String _
        ) As Boolean
    
    RangeExists = False
    
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In wkb.Worksheets
        Dim rng As Range
        Set rng = ws.Range(RangeName)
        If Not rng Is Nothing Then
            RangeExists = True
            Exit Function
        End If
    Next
    On Error GoTo 0
    
End Function


'==============================================================================
'at least in Excel 2016 there seems to be a bug when hyperlinks and
'FormatConditions should be added to cells and a Chart is (currently) active.
'Therefore store the currently active sheet, switch to 'wks' and restore it
'afterwards.
Public Function RememberActiveChartAndActivateGivenWorksheet( _
    ByVal wks As Worksheet _
        ) As Chart
    If TypeName(ActiveSheet) = "Chart" Then
        Set RememberActiveChartAndActivateGivenWorksheet = ActiveChart
        wks.Activate
    Else
        Set RememberActiveChartAndActivateGivenWorksheet = Nothing
    End If
End Function


'originally inspired by
'<https://newtonexcelbach.wordpress.com/2013/12/07/extracting-numbers-from-text-strings/>
'but heavily modified in the meantime
'
'NumberIndex_Or_SeparationString
'  Because it doesn't make sense to ask for a specific numbers index *and* to
'  state a string to separate the numbers, this argument is used for both
'  purposes. If you state a number > 0 the corresponding number will be
'  returned; if you state the number = 0 the number of found numbers will be
'  returned; if you state a number < 0 the corresponding number from the end
'  of the array will be returned and otherwise the given string will be used to
'  separate the numbers
Public Function ExtractNumbers( _
    ByVal NumString As Variant, _
    Optional ByVal NumberIndex_Or_SeparationString As Variant, _
    Optional ByVal DecimalSign As Variant, _
    Optional ByVal TakeNegative As Variant _
        ) As Variant
    
    '==========================================================================
    'negative sign
    Const csNeg As String = "-"
    'internal number separator
    Const csSep As String = " "
    '==========================================================================
    
    'initialize optional arguments, if they are not given
    If IsMissing(DecimalSign) Then
        DecimalSign = vbNullString
    End If
    If IsMissing(TakeNegative) Then
        TakeNegative = False
    ElseIf TypeName(TakeNegative) <> "Boolean" Then
        ExtractNumbers = CVErr(xlErrNA)
    End If
    'if 'SeparationString' is missing it doesn't make sense to have no
    ''SeparationString' if either 'TakeNegative' or 'DecimalSign' is given
    If IsMissing(NumberIndex_Or_SeparationString) Then
        If TakeNegative = True And DecimalSign <> vbNullString Then
            NumberIndex_Or_SeparationString = vbNullString
        Else
            NumberIndex_Or_SeparationString = " "
        End If
    End If
    
    'initialize 'sNeg'
    If TakeNegative = True Then
        Dim sNeg As String
        sNeg = csNeg 'Negative Sign MUST be before 1st number.
    Else
        sNeg = vbNullString
    End If
    
    
    'collect all valid chars and separate them by 'csSep' if in-between valid
    'chars is a non-valid char
    Dim i As Long
    For i = 1 To Len(NumString)
        Dim sChar As String
        sChar = Mid$(NumString, i, 1)
        If IsNumeric(sChar) Or sChar = sNeg Or sChar = DecimalSign Then
            'special handler for the first found valid character
            Dim sTemp As String
            If Len(sTemp) <> 0 Then
                Dim bValidChar As Boolean
                If bValidChar = True Then
                    sTemp = sTemp & Mid$(NumString, i, 1)
                Else
                    sTemp = sTemp & csSep & Mid$(NumString, i, 1)
                End If
            Else
                '"initialize" 'sTemp'
                sTemp = Mid$(NumString, i, 1)
            End If
            bValidChar = True
        Else
            bValidChar = False
        End If
    Next
    
    'convert all numbers to numbers
    '(e.g. if the number in the text was given as "-.1" then it will be
    ' formatted as "-0.1")
    Dim arrNumbers As Variant
    arrNumbers = Split(sTemp, csSep)
    
    For i = LBound(arrNumbers) To UBound(arrNumbers)
        'this test is needed because a number *range* could be found
        '(e.g. "2.5-5.5")
        If IsNumeric(arrNumbers(i)) Then
            arrNumbers(i) = CDbl(arrNumbers(i))
        End If
    Next
    
    'if 'arrNumbers' is empty then exit function
    If UBound(arrNumbers) = -1 Then
        ExtractNumbers = vbNullString
    'if 'NumberIndex_Or_SeparationString' is a number return either the
    'corresponding number with the given index or the number of found numbers
    'depending on if the given number is >0 or not
    ElseIf IsNumeric(NumberIndex_Or_SeparationString) Then
        If NumberIndex_Or_SeparationString > 0 Then
            ExtractNumbers = CDbl(arrNumbers(CInt(NumberIndex_Or_SeparationString) - 1))
        ElseIf NumberIndex_Or_SeparationString < 0 Then
            ExtractNumbers = arrNumbers(UBound(arrNumbers) + NumberIndex_Or_SeparationString + 1)
        Else
            ExtractNumbers = UBound(arrNumbers) - LBound(arrNumbers) + 1
        End If
    'otherwise return a string with all found numbers that is separated by
    'the given string of 'NumberIndex_Or_SeparationString'
    '(This is not totally correct, but have a look at the comments at
    ' 'IsMissing(NumberIndex_Or_SeparationString)' for more details)
    Else
        sTemp = arrNumbers(LBound(arrNumbers))
        For i = LBound(arrNumbers) + 1 To UBound(arrNumbers)
            sTemp = sTemp & NumberIndex_Or_SeparationString & arrNumbers(i)
        Next
        ExtractNumbers = sTemp
    End If
    
End Function


Public Function ExtractRowsRange( _
    ByVal NumString As String _
        ) As String
    
    If Len(NumString) = 0 Then Exit Function
    
    Dim i As Long
    For i = 1 To Len(NumString)
        Dim sChar As String
        sChar = Mid$(NumString, i, 1)
        
        If sChar Like "[0-9,:]" Then
            Dim sTemp As String
            sTemp = sTemp & sChar
        End If
    Next
    
    ExtractRowsRange = sTemp
    
End Function


Public Function ExtractColumnsRange( _
    ByVal NumString As String _
        ) As String
    
    If Len(NumString) = 0 Then Exit Function
    
    Dim i As Long
    For i = 1 To Len(NumString)
        Dim sChar As String
        sChar = Mid$(NumString, i, 1)
        
        If sChar Like "[A-Z,:]" Then
            Dim sTemp As String
            sTemp = sTemp & sChar
        End If
    Next
    
    ExtractColumnsRange = sTemp
    
End Function


'inspired by <https://www.extendoffice.com/documents/excel/5215-excel-check-if-row-is-hidden.html>
Public Function HiddenRowsInSheet( _
    ByVal oWorksheet As Worksheet _
        ) As String
    
    If oWorksheet Is Nothing Then
        HiddenRowsInSheet = "-1"
        Exit Function
    End If
    
    Dim wks As Worksheet
    Set wks = oWorksheet
    
    Dim rng As Range
    Set rng = wks.UsedRange
    
    Dim rngVisible As Range
    Set rngVisible = rng.SpecialCells(xlCellTypeVisible)
    If rng.Count <> rngVisible.Count Then
        Dim i As Long
        For i = 1 To rngVisible.Areas.Count - 1
            Dim rngItem As Range
            Set rngItem = rngVisible.Areas.Item(i)
            
            Dim iOne As Long
            iOne = rngItem.Rows(rngItem.Rows.Count).Row
            
            Dim iTwo As Long
            iTwo = rngVisible.Areas.Item(i + 1).Rows(1).Row
            
            If iOne < iTwo Then
                Dim sString As String
                sString = CStr(iOne + 1) & ":" & CStr(iTwo - 1)
                
                Dim sTemp As String
                sTemp = sTemp & sString & ","
            End If
        Next
        sTemp = Left$(sTemp, Application.WorksheetFunction.Max(0, Len(sTemp) - 1))
    End If
    
    'remove (possible) duplicates
    If Len(sTemp) > 0 Then
        Dim Arr As Variant
        Arr = Split(sTemp, ",")
        
        Dim ArrResult As Variant
        ArrResult = RemoveDuplicatesFromVector(Arr)
        
        Dim sResult As String
        sResult = Join(ArrResult, ",")
    Else
        sResult = sTemp
    End If
    
    HiddenRowsInSheet = sResult
    
End Function


Public Function HiddenColumnsInSheet( _
    ByVal oWorksheet As Worksheet _
        ) As String
    
    If oWorksheet Is Nothing Then
        HiddenColumnsInSheet = "-1"
        Exit Function
    End If
    
    Dim wks As Worksheet
    Set wks = oWorksheet
    
    Dim rng As Range
    Set rng = wks.UsedRange
    
    Dim rngVisible As Range
    Set rngVisible = rng.SpecialCells(xlCellTypeVisible)
    
    If rng.Count <> rngVisible.Count Then
        Dim i As Long
        For i = 1 To rngVisible.Areas.Count - 1
            Dim rngItem As Range
            Set rngItem = rngVisible.Areas.Item(i)
            
            Dim iOne As Long
            iOne = rngItem.Columns(rngItem.Columns.Count).Column
            
            Dim iTwo As Long
            iTwo = rngVisible.Areas.Item(i + 1).Columns(1).Column
            
            If iOne < iTwo Then
                Dim sString As String
                sString = ColumnNumberToLetter(iOne + 1) & ":" & _
                        ColumnNumberToLetter(iTwo - 1)
                
                Dim sTemp As String
                sTemp = sTemp & sString & ","
            End If
        Next
        sTemp = Left$(sTemp, Application.WorksheetFunction.Max(0, Len(sTemp) - 1))
    End If
    
    'remove (possible) duplicates
    If Len(sTemp) > 0 Then
        Dim Arr As Variant
        Arr = Split(sTemp, ",")
        
        Dim ArrResult As Variant
        ArrResult = RemoveDuplicatesFromVector(Arr)
        
        Dim sResult As String
        sResult = Join(ArrResult, ",")
    Else
        sResult = sTemp
    End If
    
    HiddenColumnsInSheet = sResult
    
End Function


Private Function ColumnNumberToLetter( _
    ByVal lngNumber As Long, _
    Optional ByVal bAbsolute As Boolean = False _
        ) As String
    
    Dim sDummy As String
    sDummy = Split(ThisWorkbook.Worksheets(1).Columns(lngNumber).Address, ":")(0)
    
    If Not bAbsolute Then sDummy = Right$(sDummy, Len(sDummy) - 1)
    ColumnNumberToLetter = sDummy
    
End Function


'---
'TODO:
'Maybe move to 'modArraySupport2'
'---
'DESCRIPTION: Removes duplicates from your array using the collection method.
'NOTES: (1) This function returns unique elements in your array, but
'           it converts your array elements to strings.
'SOURCE: <https://wellsr.com/vba/2017/excel/vba-remove-duplicates-from-array/>
Private Function RemoveDuplicatesFromVector( _
    ByVal Arr As Variant _
        ) As Variant
    
    Dim arrDummy() As Variant
    ReDim arrDummy(LBound(Arr) To UBound(Arr))
    
    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)    'convert to string
        arrDummy(i) = CStr(Arr(i))
    Next
    
    Dim arrColl As Collection
    Set arrColl = New Collection
    
    On Error Resume Next
    Dim Element As Variant
    For Each Element In arrDummy
        arrColl.Add Element, Element
    Next
    Err.Clear
    On Error GoTo 0
    
    Dim arrUnique() As Variant
    ReDim arrUnique(LBound(Arr) To LBound(Arr) + arrColl.Count - 1)
    
    i = LBound(Arr)
    
    For Each Element In arrColl
        arrUnique(i) = Element
        i = i + 1
    Next
    
    RemoveDuplicatesFromVector = arrUnique
    
End Function


'==============================================================================
'test, if a given Sheet is of the type 'xlChart'
'(this has to be tested indirectly, because of a bug in Excel
' have a look at
'    <https://excel.tips.net/T002538_Detecting_Types_of_Sheets_in_VBA.html>
' for more details)
Public Function IsChart( _
    ByVal wkb As Workbook, _
    ByVal iSheetIndex As Long _
        ) As Boolean
    
    With wkb
        If TypeName(.Sheets(iSheetIndex)) = "Chart" _
                Or .Sheets(iSheetIndex).Type = xlChart _
                Or .Sheets(iSheetIndex).Type = xlExcel4MacroSheet Then
            Dim cht As Chart
            For Each cht In .Charts
                If cht.Name = .Sheets(iSheetIndex).Name Then
                    IsChart = True
                    Exit Function
                End If
            Next
        End If
    End With
End Function
