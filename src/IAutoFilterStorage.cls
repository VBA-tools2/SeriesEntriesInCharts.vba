VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAutoFilterStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Describes the 'AutoFilterStorage' object."

'@Folder("AutoFilter")
'@ModuleDescription("Describes the 'AutoFilterStorage' object.")
'@Interface

Option Explicit

Public Enum eAutoFilterStorageError
    [_First] = vbObjectError + 1
    ErrIsNothing = [_First]
    ErrNotStored
    [_Last] = ErrIsNothing
End Enum

'@Description("Store the 'AutoFilter's")
Public Sub StoreFilters()
Attribute StoreFilters.VB_Description = "Store the 'AutoFilter's"
End Sub

'@Description("Restore the 'AutoFilter's")
Public Sub RestoreFilters()
Attribute RestoreFilters.VB_Description = "Restore the 'AutoFilter's"
End Sub

'@Description("Get the 'AutoFilterMode'")
Public Property Get AutoFilterMode() As Boolean
End Property

'@Description("Set the 'AutoFilterMode'")
Public Property Let AutoFilterMode(ByVal Value As Boolean)
End Property
