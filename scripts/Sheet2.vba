Option Explicit
Private Sub addDate_Button_Click()
    Call PositionAttendanceColomnButtons
    addDate_UserForm.Show
End Sub
Private Sub removeDate_Button_Click()
    Call PositionAttendanceColomnButtons
    removeDate_UserForm.Show
End Sub

Private Sub worksheet_activate()
    If Worksheets("COMPUTING DON'T TOUCH").Cells(15, 2).Value = "N" Then
        Dim wasEnabled As Boolean
        wasEnabled = Application.EnableEvents
        Application.EnableEvents = False
        
        Call AttendanceData_load
        Call AttendanceData_save
        
        Application.EnableEvents = wasEnabled
    End If
End Sub
Private Sub Worksheet_Change(ByVal target As Excel.Range)
    
    Debug_msg ("Saving Started")
    AttendanceSaving = True
           
    Dim intersection As Range
    Set intersection = Intersect(target, Range("C3:BN" & maxMembers + 2))
    If Not intersection Is Nothing Then
        Dim wasEnabled1 As Boolean
        wasEnabled1 = Application.EnableEvents
        Application.EnableEvents = False
        
        For Each x In intersection
            x.Value = UCase(x.Value)
        Next
        
        Call UpdateAttendanceList
        Call ScanCommonError
        Application.EnableEvents = wasEnabled1
        Debug_msg ("Saving Done")
    End If
    
    
    Set intersection = Intersect(target, Range("B1"))
    If Not intersection Is Nothing Then
        Dim wasEnabled2 As Boolean
        wasEnabled2 = Application.EnableEvents
        Application.EnableEvents = False
        Call UpdateAttendanceList
        Call ScanCommonError
        Debug_msg ("Up-to-Day Saved")
        Application.EnableEvents = wasEnabled2
    End If
      
    AttendanceSaving = False
End Sub