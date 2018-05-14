'\\\Initializer///'
Private Sub UserForm_Initialize()
    Call IREC
    Members_list.Clear
    Dim row As Integer
    
    For row = 2 To maxMembers + 1
        If Worksheets("Details").Cells(row, 1).Value & Worksheets("Details").Cells(row, 2).Value <> "" Then
            Members_list.AddItem (Worksheets("Details").Cells(row, 1).Value & " " & Worksheets("Details").Cells(row, 2).Value), row - 2
        End If
    Next
End Sub
'//////// \\\\\\\\\'

Sub RemoveMemberRow(ByVal row As Integer)
    Dim wasEnabled As Boolean
    wasEnabled = Application.EnableEvents
    Application.EnableEvents = False
    Application.StatusBar = "Removing Valued Member :'("
    Application.ScreenUpdating = False

    Dim column As Integer
    For column = 1 To 7
        Worksheets("Details").Cells(row, column).Value = ""
    Next
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.EnableEvents = wasEnabled
End Sub
Private Sub Cancel_btn_Click()
    Unload MemberRemove_UserForm
End Sub
Private Sub Submit_btn_Click()
    Dim victim As String
    victim = ""
    Dim index As Integer
    
    For index = 0 To Members_list.ListCount - 1
        If Members_list.Selected(index) Then
            victim = Members_list.List(index)
            Exit For
        End If
    Next
    
    If victim = "" Then
    
    ElseIf LCase(victim) = "caleb serafin" And Worksheets("COMPUTING DON'T TOUCH").Cells(5, 12).Value = "Drama Club" Then
        MemberRemoveConfirm_UserForm.Show
    Else
        RemoveMemberRow (FindMember(Split(victim, " ")(0), Split(victim, " ")(1), True))
        Call ScanCommonError
        Unload MemberRemove_UserForm
    End If
End Sub
