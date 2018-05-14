Option Explicit
Private Function dataTableLoad() As String
    Dim rangeString As String
    rangeString = ""
    Dim atI As Integer
    atI = 1
    Call IREC
    For atI = 1 To maxMembers + 1
        rangeString = rangeString & Worksheets("Details").Cells(atI + 1, 2)
    Next atI
    dataTableLoad = rangeString
End Function
Private Function dataTableChanged(ByVal safe As Boolean) As Boolean
    Debug_msg ("Sheet1: dataTableChanged() Called")
    dataTableChanged = False
    Dim dataTableNew As String
    dataTableNew = dataTableLoad()
    If dataTableNew <> dataTableOld Then
        dataTableChanged = True
    End If
    If safe = False Then
        dataTableOld = dataTableLoad()
    End If
    
    Debug_msg ("Sheet1: dataTableChanged() End Function = " & dataTableChanged & ", safe = " & safe)
End Function
Sub DetailsCheck()
    Debug_msg ("Sheet1: DetailsCheck() Called")
    Dim wasEnabled As Boolean
    wasEnabled = Application.EnableEvents
    Application.EnableEvents = False
    If AttendanceSaving <> True And Worksheets("COMPUTING DON'T TOUCH").Cells(15, 2).Value = "Y" Then
        Dim hasTableChanged As Boolean
        hasTableChanged = dataTableChanged(False)
        
        Debug_msg ("Sheet1: DetailsCheck(): Table Change: " & hasTableChanged)
        If hasTableChanged Then
            Call AttendanceData_load
            Call ScanCommonError
        End If
    End If
    Application.EnableEvents = wasEnabled
End Sub
Private Sub MemberAdd_Button_Click()
    MemberAdd_UserForm.Show
End Sub
Private Sub MemberRemove_Button_Click()
    MemberRemove_UserForm.Show
End Sub
Private Sub Refresh_Button_Click()
    Debug_msg ("Sheet 1: Refresh_Button_Click pressed")
    Dim wasEnabled As Boolean
    wasEnabled = Application.EnableEvents
    Application.EnableEvents = False
    Call AttendanceData_load
    Application.EnableEvents = wasEnabled
End Sub

Private Sub Worksheet_Change(ByVal target As Excel.Range)
    Debug_msg ("Sheet1: Worksheet_Change() Invoked")
    Call DetailsCheck
End Sub

Private Sub Worksheet_Calculate()
    Debug_msg ("Sheet1: Worksheet_Calculate() Invoked")
    Call DetailsCheck
End Sub