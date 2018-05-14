Private Sub Cancel_btn_Click()
    MemberAdd_UserForm.Hide
End Sub

Private Sub Submit_btn_Click()
    If FindMember(FirstName_txt.Value, LastName_txt.Value, False) = 0 Then
        Dim wasEnabled As Boolean
        wasEnabled = Application.EnableEvents
        Application.EnableEvents = False
        Application.StatusBar = "Adding Oak..."
        Application.ScreenUpdating = False
        
        Dim AttOption As String
        AttOption = ""
        
        If AttB_opt.Value Then
            AttOption = "0"
        ElseIf AttY_opt.Value Then
            AttOption = "1"
        ElseIf AttN_opt.Value Then
            AttOption = "2"
        ElseIf AttQ_opt.Value Then
            AttOption = "3"
        End If
        
        Call addCellData("row", "Details", 2, maxMembers + 1, FirstName_txt.Value & "\'\" & LastName_txt.Value & "\'\" & Grade_txt.Value & "\'\" & Group_txt.Value & "\'\" & Role_txt.Value & "\'\" & Phone_txt.Value & "\'\" & Email_txt.Value & "\'\v2_" & StringMult(AttOption, Int(Worksheets("Attendance").Range("B1").Value)), 1, True) '<---This is what requires v2 Load and Save functions (the v2_ part)
        Call AttendanceData_load
        Call UpdateAttendanceList
        
        Application.ScreenUpdating = True
        Application.StatusBar = False
        Application.EnableEvents = wasEnabled
        
        Unload MemberAdd_UserForm
    Else
        MsgBox (LCase(FirstName_txt.Value) & " " & LCase(LastName_txt.Value) & " is already a member! Check your list."), vbInformation, "Add New Oak"
    End If
End Sub

Private Sub UserForm_Click()

End Sub