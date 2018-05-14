Option Explicit
Private pParentDataType As String
Private pRootName As String

Private Sub Class_Initialize(ByVal ParentType As String, ByVal RootName As String)
    pParentDataType = ParentType
    pRootName = RootName
End Sub
