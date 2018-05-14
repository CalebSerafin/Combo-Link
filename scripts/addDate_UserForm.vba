'\\\Initializer///'
Private Sub UserForm_Initialize()
    Dim LDate() As String
    LDate = Split(Date, "/")
    nameDate_textBox.Value = LDate(0) & "-" & GetMonth(Int(LDate(1)))
End Sub
'//////// \\\\\\\\\'

Private Sub Insert_check_Click()
    If Insert_check Then
        With InsertIndex_TextBox
            .Enabled = True
            .BackColor = &H80000005
        End With
    Else
        With InsertIndex_TextBox
            .Enabled = False
            .BackColor = &H8000000F
        End With
    End If
End Sub
Function getMemberNo() As Integer
    Dim i As Integer
    For i = 0 To Int(Worksheets("COMPUTING DON'T TOUCH").Range("F15").Value)
        If Worksheets("Attendance").Cells(i + 3, 1).Value = " " Then
           Exit For
        End If
    Next
    getMemberNo = i
End Function
Sub fillAttendanceColomn(ByVal colomn As Integer, ByVal char As String)
    If colomn >= 1 Then
        Dim i As Integer
        For i = 1 To getMemberNo()
            Worksheets("Attendance").Cells(i + 2, colomn).Value = char
        Next
    End If
End Sub
Sub emptyAttendanceColomn(ByVal colomn As Integer)  '====Includes the date/heading!
    If colomn >= 1 Then
        Dim i As Integer
        For i = 0 To getMemberNo()
            Worksheets("Attendance").Cells(i + 2, colomn).Value = ""
        Next
    End If
End Sub
Sub pushColomns(ByVal fromColomn As Integer, ByVal shiftAmount As Integer)
    Dim proceeding As Integer
    Dim row As Integer
    Dim colomnData() As String
    Dim colomnRaw As String
    
    For proceeding = (Worksheets("Attendance").Range("B1").Value + 2) To (fromColomn) Step -1
        Erase colomnData()
        colomnRaw = Worksheets("Attendance").Cells(2, proceeding).Value
        For row = 3 To getMemberNo() + 2
            colomnRaw = colomnRaw & "\'\" & Worksheets("Attendance").Cells(row, proceeding).Value
        Next
        
        colomnData() = Split(colomnRaw, "\'\")
        
        For row = 2 To getMemberNo() + 2
            Worksheets("Attendance").Cells(row, proceeding + shiftAmount).Value = colomnData(row - 2)
        Next
        Call emptyAttendanceColomn(proceeding)
    Next proceeding
End Sub



'Sub UserForm_Initialize()
'    With nameDate_textBox
'        .ForeColor = &HC0C0C0
'        .Text = "01-Apr                     ."
'    End With
'    blank_opt.SetFocus
'End Sub
'
'Private Sub nameDate_textBox_Enter()
'    With nameDate_textBox
'        If .Text = ".                     01-Apr" Or .Text = "01-Apr                     ." Then
'            .ForeColor = &H80000008
'            .Text = ""
'        End If
'    End With
'End Sub
'
'Private Sub nameDate_textBox_AfterUpdate()
'    With nameDate_textBox
'        If .Text = "" Then
'            .ForeColor = &HC0C0C0
'            .Text = ".                     01-Apr"
'        End If
'    End With
'End Sub

Private Sub submit_button_Click()
    AttendanceSaving = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If Insert_check = True Then
        If IsNumeric(InsertIndex_TextBox) Then
            If 1 <= Int(InsertIndex_TextBox) And Int(InsertIndex_TextBox) <= Worksheets("Attendance").Range("B1").Value Then
                Call pushColomns(Int(InsertIndex_TextBox) + 2, 1)
                Call PositionAttendanceColomnButtons(Int(InsertIndex_TextBox) + 2)
            Else
                MsgBox ("Your position cannot go under day 1 or go above day " & Worksheets("Attendance").Range("B1").Value & "!"), vbOKOnly, "Add Date"
            End If
        Else
            MsgBox ("You are meant to put the index number into the position box(The italic number at the top)"), vbExclamation, "Add Date"
        End If
    End If
    
    Dim symbol As String
    If yes_opt = True Then
     symbol = "Y"
    ElseIf no_opt = True Then
        symbol = "N"
    ElseIf questionMark_opt = True Then
        symbol = "?"
    ElseIf copy_opt = True Then
        '///==============
        Dim row As Integer
        Dim source As Integer
        Dim colomnData() As String
        Dim colomnRaw As String
        Erase colomnData()
        source = Worksheets("Attendance").Range("B1").Value + 2
        
        colomnRaw = Worksheets("Attendance").Cells(3, source).Value
        For row = 4 To getMemberNo() + 2
            colomnRaw = colomnRaw & "\'\" & Worksheets("Attendance").Cells(row, source).Value
        Next
        
        colomnData() = Split(colomnRaw, "\'\")
        
        For row = 3 To getMemberNo() + 2
            Worksheets("Attendance").Cells(row, Worksheets("Attendance").addDate_Button.TopLeftCell.colomn).Value = colomnData(row - 3)
        Next
        '///==============
    Else: symbol = ""
    End If
    
    If Not (blank_opt Or copy_opt) Then
        Call fillAttendanceColomn(Worksheets("Attendance").addDate_Button.TopLeftCell.column, symbol)
    End If
    
    Worksheets("Attendance").addDate_Button.TopLeftCell.Value = nameDate_textBox.Value
    Worksheets("Attendance").Range("B1").Value = Worksheets("Attendance").Range("B1").Value + 1
    Call PositionAttendanceColomnButtons
    Call UpdateAttendanceList
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    AttendanceSaving = False
    
    addDate_UserForm.Hide
End Sub

Private Sub Cancel_Button_Click()
    addDate_UserForm.Hide
End Sub