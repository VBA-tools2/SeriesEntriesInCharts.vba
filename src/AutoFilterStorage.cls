VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoFilterStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An object to store the 'AutoFilter's."

'@Folder("AutoFilter")
'@ModuleDescription("An object to store the 'AutoFilter's.")
'@PredeclaredId

'based on <https://stackoverflow.com/a/15379095/5776000>

'Public Sub ExampleUsage()
'    Dim wks As Worksheet
'    Set wks = ActiveSheet
'    Dim MyAFStorage As IAutoFilterStorage
'    Set MyAFStorage = AutoFilterStorage.Create(wks)
'    With MyAFStorage
'        .StoreFilters
'        .RestoreFilters
'    End With
'End Sub

Option Explicit

Implements IAutoFilterStorage


Private Type TAutoFilterStorage
    wks As Worksheet
    AutoFilterMode As Boolean
    FilterRange As Range
    FilterCache As Variant
    
    InitialAutoFilterMode As Boolean
    AreFiltersStored As Boolean
End Type
Private This As TAutoFilterStorage


'@Description("Factory method to create the object (itself)")
Public Function Create(ByVal Value As Worksheet) As IAutoFilterStorage
Attribute Create.VB_Description = "Factory method to create the object (itself)"
    With New AutoFilterStorage
        Set .wks = Value
        Set Create = .Self
    End With
End Function


'@Description("Sets the 'AutoFilterStorage' object (itself)")
Friend Property Get Self() As IAutoFilterStorage
Attribute Self.VB_Description = "Sets the 'AutoFilterStorage' object (itself)"
    Set Self = Me
End Property


'@Description("Stores the Worksheet")
Friend Property Set wks(ByVal Value As Worksheet)
Attribute wks.VB_Description = "Stores the Worksheet"
    If Value Is Nothing Then RaiseErrorIsNothing
    
    Set This.wks = Value
End Property


'@Description("Get the 'AutoFilterMode'")
Private Property Get IAutoFilterStorage_AutoFilterMode() As Boolean
Attribute IAutoFilterStorage_AutoFilterMode.VB_Description = "Get the 'AutoFilterMode'"
    With This
        .AutoFilterMode = .wks.AutoFilterMode
        IAutoFilterStorage_AutoFilterMode = .AutoFilterMode
    End With
End Property


'@Description("Set the 'AutoFilterMode'")
Private Property Let IAutoFilterStorage_AutoFilterMode(ByVal Value As Boolean)
Attribute IAutoFilterStorage_AutoFilterMode.VB_Description = "Set the 'AutoFilterMode'"
    With This
        .wks.AutoFilterMode = Value
        .AutoFilterMode = Value
    End With
End Property


'@Description("Store the 'AutoFilter' settings")
Public Sub IAutoFilterStorage_StoreFilters()
Attribute IAutoFilterStorage_StoreFilters.VB_Description = "Store the 'AutoFilter' settings"
    
    With This
        .InitialAutoFilterMode = .wks.AutoFilterMode
        If Not .InitialAutoFilterMode Then Exit Sub
        
        Set .FilterRange = .wks.AutoFilter.Range
        
        With .wks.AutoFilter.Filters
            ReDim This.FilterCache(1 To .Count, 1 To 3)
            
            Dim i As Long
            For i = 1 To .Count
                With .Item(i)
                    If .On Then
                        Select Case .Operator
                            Case 1, 2           'xlAnd, xlOr
                                This.FilterCache(i, 1) = .Criteria1
                                This.FilterCache(i, 2) = .Operator
                                This.FilterCache(i, 3) = .Criteria2
                            Case 0, 3 To 7      'no operator, xlTop10Items, xlBottom10Items, _
                                                 xlTop10Percent, xlBottom10Percent, _
                                                 xlFilterValues
                                This.FilterCache(i, 1) = .Criteria1
                                This.FilterCache(i, 2) = .Operator
                            Case Else           'These are not correctly restored; there's something in Criteria1 but can't save it.
                                This.FilterCache(i, 2) = .Operator
'                                This.FilterCache(i, 1) = .Criteria1   ' <-- Generates an error
'                                'No error in next statement, but couldn't do restore operation
'                                Set This.FilterCache(i, 1) = .Criteria1
                        End Select
                    End If
                End With
            Next
        End With
        
        .AreFiltersStored = True
    End With
    
End Sub


'@Description("Restore the 'AutoFilter' settings")
Public Sub IAutoFilterStorage_RestoreFilters()
Attribute IAutoFilterStorage_RestoreFilters.VB_Description = "Restore the 'AutoFilter' settings"
    With This
        If Not .InitialAutoFilterMode Then Exit Sub
        If Not .AreFiltersStored Then RaiseErrorNotStored
        If .FilterRange Is Nothing Then Exit Sub
        .wks.AutoFilterMode = False
        
        With .FilterRange
            .AutoFilter
            
            Dim col As Long
            For col = 1 To UBound(This.FilterCache, 1)
                If Not IsEmpty(This.FilterCache(col, 2)) Then
                    Select Case This.FilterCache(col, 2)
                        Case 0          'no operator
                            .AutoFilter _
                                    Field:=col, _
                                    Criteria1:=This.FilterCache(col, 1)
                                    'Do NOT reload 'Operator'
                        Case 1, 2       'xlAnd, xlOr
                            .AutoFilter _
                                    Field:=col, _
                                    Criteria1:=This.FilterCache(col, 1), _
                                    Operator:=This.FilterCache(col, 2), _
                                    Criteria2:=This.FilterCache(col, 3)
                        Case 3 To 6     'xlTop10Items, xlBottom10Items, xlTop10Percent, xlBottom10Percent
                            .AutoFilter _
                                    Field:=col, _
                                    Criteria1:=This.FilterCache(col, 1)
                                    'Do NOT reload 'Operator', it doesn't work
'                            This.wks.AutoFilter.Filters.Item(col).Operator = This.FilterCache(col, 2)
                        Case 7          'xlFilterValues
                            .AutoFilter _
                                    Field:=col, _
                                    Criteria1:=This.FilterCache(col, 1), _
                                    Operator:=This.FilterCache(col, 2)
                    End Select
                End If
            Next
        End With
    End With
End Sub


'=============================================================================
'@Description("Raise an error if 'wks' is 'Nothing'")
Private Sub RaiseErrorIsNothing()
Attribute RaiseErrorIsNothing.VB_Description = "Raise an error if 'wks' is 'Nothing'"
    Err.Raise _
            Source:=TypeName(Me), _
            Number:=eAutoFilterStorageError.ErrIsNothing, _
            Description:= _
                    "'wks' is 'Nothing' what is not allowed."
End Sub


'@Description("Raise an error if it is tried to restore the filters before storing them")
Private Sub RaiseErrorNotStored()
Attribute RaiseErrorNotStored.VB_Description = "Raise an error if it is tried to restore the filters before storing them"
    Err.Raise _
            Source:=TypeName(Me), _
            Number:=eAutoFilterStorageError.ErrNotStored, _
            Description:= _
                    "You try to restore the filters without storing them first." & _
                    vbCrLf & _
                    "Please do that first."
End Sub
