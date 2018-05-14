'\\\Initializers///'
'//////// \\\\\\\\\'
Option Explicit

Private Sub CommandButton1_Click()
    Application.EnableEvents = True
End Sub

Private Sub CommandButton2_Click()
    Dim wasEnabled As Boolean
    wasEnabled = Application.EnableEvents
    Application.EnableEvents = False
    Call AttendanceData_save
    Application.EnableEvents = wasEnabled
End Sub

Private Sub CommandButton3_Click()
    Dim wasEnabled As Boolean
    wasEnabled = Application.EnableEvents
    Application.EnableEvents = False
    Call AttendanceData_load
    Application.EnableEvents = wasEnabled
End Sub

Private Sub CommandButton4_Click()
    Application.EnableEvents = False
End Sub

Private Sub CommandButton5_Click()
    AttendanceSaving = False
End Sub

Private Sub CommandButton6_Click()
    MsgBox (AttendanceSaving)
End Sub

Private Sub CommandButton7_Click()
    AttendanceSaving = True
End Sub

Private Sub CommandButton8_Click()
    MsgBox Application.EnableEvents
End Sub

Private Sub GitRead_btn_Click()
    If Worksheets("COMPUTING DON'T TOUCH").Range("J15").Value = "Y" Then
        On Error GoTo VbaGitBootStrapNotFound
        Application.Run ("'VbaGitBootStrap.xlsm'!GitRead")
        'Call Debug_msg("Sheet 3: GitRead_btn_Click: Git Read has been disabled due to insufficant testing. Sorry", , "Notify")
    Else
        Call Debug_msg("Sheet 3: GitRead_btn_Click: Git Enabled is not Enabled! at COMPUTING DON'T TOUCH Cell J15", , "Notify")
    End If
    Exit Sub
    
VbaGitBootStrapNotFound:
    Call Debug_msg("Sheet 3: GitRead_btn_Click: 'VbaGitBootStrap.xlsm' was not found. Most likely because it is closed or does not have macros enabled, or you are not a developer for Combo Link.", , "Notify")
End Sub

Private Sub GitWrite_btn_Click()
    If Worksheets("COMPUTING DON'T TOUCH").Range("J15").Value = "Y" Then
        On Error GoTo VbaGitBootStrapNotFound
        Worksheets("COMPUTING DON'T TOUCH").Range("J15").Value = "N" 'Disables Git Controls before export
        Application.Run ("'VbaGitBootStrap.xlsm'!GitWrite")
        Call Debug_msg("Sheet 3: GitRead_btn_Click: Exported Successfully to Staging Area!", , "Notify")
    Else
        Call Debug_msg("Sheet 3: GitRead_btn_Click: Git Enabled is not Enabled! at COMPUTING DON'T TOUCH Cell J15", , "Notify")
    End If
    Exit Sub
    
VbaGitBootStrapNotFound:
    Worksheets("COMPUTING DON'T TOUCH").Range("J15").Value = "Y" 'Enables Git Controls due to unsuccessful export
    Call Debug_msg("Sheet 3: GitRead_btn_Click: 'VbaGitBootStrap.xlsm' was not found. Most likely because it is closed or does not have macros enabled, or you are not a developer for Combo Link.", , "Notify")
End Sub