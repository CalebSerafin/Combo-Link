Option Explicit
Public Score As cAttributes
Sub Score_Initialize()
    Score.create "Score"
    
    With Score.TableSlots
        .IsEnabled = True
        .SideBuffer = 2
        .TopBuffer = 2
        .HeaderFields = 0
    End With
End Sub
