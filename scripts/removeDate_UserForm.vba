Function getMemberNo() As Integer
    Dim i As Integer
    For i = 0 To Int(Worksheets("COMPUTING DON'T TOUCH").Range("F15").Value)
        If Worksheets("Attendance").Cells(i + 3, 1).Value = " " Then
           Exit For
        End If
    Next
    getMemberNo = i
End Function
Sub emptyAttendanceColomn(ByVal colomn As Integer)  '====Includes the date/heading!
    If colomn >= 1 Then
        Dim i As Integer
        For i = 0 To getMemberNo()
            Worksheets("Attendance").Cells(i + 2, colomn).Value = ""
        Next
    End If
End Sub

Sub UserForm_Initialize()
    Call LastColomn_Opt_Click
End Sub


Private Sub LastColomn_Opt_Click()
    With FromDay_TextBox
        .Enabled = False
        .BackColor = &H8000000F
    End With
    With ToDay_TextBox
        .Enabled = False
        .BackColor = &H8000000F
    End With
    With StayOpen_check
        .Enabled = True
    End With
    FromDay_Label.Enabled = False
    ToDay_Label.Enabled = False
End Sub
Private Sub Range_Opt_Click()
    With FromDay_TextBox
        .Enabled = True
        .BackColor = &H80000005
    End With
    With ToDay_TextBox
        .Enabled = True
        .BackColor = &H80000005
    End With
    With StayOpen_check
        .Enabled = False
    End With
    FromDay_Label.Enabled = True
    ToDay_Label.Enabled = True
End Sub
Private Sub Remove_Button_Click()
    AttendanceSaving = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If LastColomn_Opt.Value = True Then
        '///code
        '///previous row
        If Worksheets("Attendance").removeDate_Button.TopLeftCell.column >= 5 Then '============================add first colomn delete
            Call emptyAttendanceColomn(Worksheets("Attendance").removeDate_Button.TopLeftCell.column - 1)
            Worksheets("Attendance").Cells(2, Worksheets("Attendance").removeDate_Button.TopLeftCell.column - 1).Value = ""
        End If '============================add first colomn delete
        Worksheets("Attendance").Range("B1").Value = Worksheets("Attendance").Range("B1").Value - 1
        
        If StayOpen_check = False Then
            removeDate_UserForm.Hide
        End If
        '\\\code
    ElseIf Range_Opt.Value = True Then
        If IsNumeric(FromDay_TextBox.Value) And IsNumeric(ToDay_TextBox.Value) Then
            If (1 <= Int(FromDay_TextBox.Value)) And (Int(ToDay_TextBox.Value) <= Worksheets("Attendance").Range("B1").Value) Then
                If Int(FromDay_TextBox.Value) <= Int(ToDay_TextBox.Value) Then
                    '///code
                    '///clear range
                    Dim target As Integer
                    For target = (FromDay_TextBox.Value + 2) To (ToDay_TextBox.Value + 2)
                        Call emptyAttendanceColomn(target)
                    Next
                    
                    '///range filling
                    Dim proceeding As Integer
                    Dim row As Integer
                    Dim colomnData() As String
                    Dim colomnRaw As String
                    target = FromDay_TextBox.Value + 2
                    
                    For proceeding = (ToDay_TextBox.Value + 3) To (Worksheets("Attendance").Range("B1").Value + 2)
                        Erase colomnData()
                        colomnRaw = Worksheets("Attendance").Cells(2, proceeding).Value
                        For row = 3 To getMemberNo() + 2
                            colomnRaw = colomnRaw & "\'\" & Worksheets("Attendance").Cells(row, proceeding).Value
                        Next
                        colomnData() = Split(colomnRaw, "\'\")
                        For row = 2 To getMemberNo() + 2
                            Worksheets("Attendance").Cells(row, target).Value = colomnData(row - 2)
                        Next
                        target = target + 1
                        Call emptyAttendanceColomn(proceeding)
                    Next proceeding
                    
                    'If FromDay_TextBox.Value = 1 Then
                    'End If
                    Worksheets("Attendance").Range("B1").Value = Worksheets("Attendance").Range("B1").Value - (ToDay_TextBox.Value - FromDay_TextBox.Value + 1)
                    '\\\code
                    Unload removeDate_UserForm
                Else
                    MsgBox "You can't travel forward in time and then decide to travel back in time! (Your first date was larger than your second date)", vbExclamation, "Doctor Who?"
                End If
            Else
                MsgBox ("Your dates cannot go under day 1 or go above day " & Worksheets("Attendance").Range("B1").Value & "!"), vbOKOnly, "Remove Dates"
            End If
        Else
            MsgBox ("You are meant to put the index numbers into the two boxes(The italic numbers at the top)"), vbExclamation, "Remove Dates"
        End If
    End If
    Call PositionAttendanceColomnButtons
    Call UpdateAttendanceList
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    AttendanceSaving = False
    
End Sub

Private Sub Cancel_Button_Click()
     removeDate_UserForm.Hide
End Sub