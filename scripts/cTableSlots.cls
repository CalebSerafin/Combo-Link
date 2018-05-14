Option Explicit
Private pFields As Integer 'how many boxes in one slot
Private pDirection As String 'going in directions of right, left, down, up
Private pTopBuffer As Integer 'how many columns inbetween it to side
Private pSideBuffer As Integer 'how many row inbetween it to top
Private pMaximium As Integer 'what is the maximium, use -1 for infinate
Private pHeaderFields As Integer 'how many the tops used in calculations down columns

Property Get Fields() As String
     Fields = pFields
End Property
Property Let Fields(Value As String)
     pFields = Value
End Property
Property Get Direction() As String
     Direction = pDirection
End Property
Property Let Direction(Value As String)
     pDirection = Value
End Property
Property Get TopBuffer() As String
     TopBuffer = pTopBuffer
End Property
Property Let TopBuffer(Value As String)
     pTopBuffer = Value
End Property
Property Get SideBuffer() As String
     SideBuffer = pSideBuffer
End Property
Property Let SideBuffer(Value As String)
     pSideBuffer = Value
End Property
Property Get Maximium() As String
     Maximium = pMaximium
End Property
Property Let Maximium(Value As String)
     pMaximium = Value
End Property
Property Get HeaderFields() As String
     HeaderFields = pHeaderFields
End Property
Property Let HeaderFields(Value As String)
     pHeaderFields = Value
End Property

Public Function store()

End Function

Private Sub Class_Initialize()
    Fields = 1
    Direction = "left"
    TopBuffer = 0
    SideBuffer = 0
    Maximium = -1
    HeaderFields = 0
End Sub