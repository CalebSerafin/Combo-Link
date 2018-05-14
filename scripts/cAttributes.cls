Option Explicit
Private pAttributesVersion As String
Private pSheetName As String
Private pEnabledMethods() As String


Private Sub Class_Initialize(ByVal SheetName As String, ParamArray EnabledMethods() As Variant)
    pAttributesVersion = "0.2"  'REMEMBER TO UPDATE ACCORDINGLY
    pSheetName = Name
    pEnabledMethods = EnabledMethods
    
    If IsInArray(pEnabledMethods, "TableSlots") Then
        Public TableSlots As cTableSlots
    End If
    
    If IsInArray(pEnabledMethods, "DetailsAPI") Then
        Public DetailsAPI As cDetailsAPI
    End If
    
End Sub
Property Get AttributesVersion() As String
     AttributesVersion = pAttributesVersion
End Property
Property Get SheetName() As String
     SheetName = pSheetName
End Property
Public Function IsEnabledMethod(ByVal Method As String) As Boolean
     IsEnabledMethod = IsInArray(pEnabledMethods, Method)
End Function
Property Get EnabledMethods() As Variant
    EnabledMethods = pEnabledMethods
End Property