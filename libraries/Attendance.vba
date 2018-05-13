'===============================================================
'Global Functions and Subs
'
'
'===============================================================
Sub AttendanceData_save_v2() 'String version 'Uncompressed
    Debug_msg ("Module1: AttendanceData_save_v2(): Application.EnableEvents found at " & Application.EnableEvents)
    If Application.EnableEvents <> True Then
        Debug_msg ("Module1: AttendanceData_save_v2() started")
        Dim att_row As Integer
        att_row = 3
        Dim att_column As Integer
        att_column = 3
        
        Dim det_row As Integer
        det_row = 2
        Dim det_column As Integer
        det_column = 8
        
        Dim practiceNo As Integer
        practiceNo = Worksheets("Attendance").Cells(1, 2).Value
        Dim serial As String
        serial = ""
        
        For att_row = 3 To maxMembers + 2
            
            serial = ""
            For att_column = 3 To practiceNo + 2
                If isEmpty(Worksheets("Attendance").Cells(att_row, att_column).Value) Then serial = serial & 0
                If Worksheets("Attendance").Cells(att_row, att_column).Value = "Y" Then serial = serial & 1
                If Worksheets("Attendance").Cells(att_row, att_column).Value = "N" Then serial = serial & 2
                If Worksheets("Attendance").Cells(att_row, att_column).Value = "?" Then serial = serial & 3
            Next att_column
            
            Worksheets("Details").Cells(det_row, det_column).Value = "v2_" & serial
            det_row = det_row + 1
        Next att_row
    End If
End Sub
Sub UpdateAttendanceList_v1(Optional ByVal save As Boolean = True)
    Debug_msg ("Attendance: UpdateAttendanceList_v1 called")
    If Application.EnableEvents <> True Then
        Dim row As Integer
        row = 3
        Dim column As Integer
        column = 2
        Dim practiceNo As Integer
        practiceNo = Worksheets("Attendance").Cells(1, 2).Value
        
        For row = 3 To maxMembers + 2
        Dim sum As Integer
        sum = 0
            For column = 3 To practiceNo + 2
                If Worksheets("Attendance").Cells(row, column).Value = "Y" Then sum = sum + 1
            Next column
            Worksheets("Attendance").Cells(row, 2).Value = sum / practiceNo
            Worksheets("Details").Cells(row - 1, 9).Value = sum / practiceNo
        Next row
        
        If save = True Then
            Debug_msg ("Module1: UpdateAttendanceList_v1: proceeding with save function")
            Call AttendanceData_save
        End If
        
    End If
End Sub
Sub AttendanceData_load_v2() 'String version 'Uncompressed
    If AttendanceSaving <> True And Application.EnableEvents <> True Then
        Application.StatusBar = "Please Wait ... Syncing Attendance List: "
        Application.ScreenUpdating = False
        
        Dim practiceNo As Integer
        practiceNo = Worksheets("Attendance").Cells(1, 2).Value
        Dim serial As String
        serial = ""
    
        Dim att_row As Integer
        att_row = 3
        Dim att_column As Integer
        att_column = 3
        
        Dim det_row As Integer
        det_row = 2
        Dim det_column As Integer
        det_column = 8
        
        For att_row = 3 To maxMembers + 2
            Application.StatusBar = "Please Wait ... Syncing Attendance List: " & att_row - 3 & "/" & maxMembers
            
            serial = Mid(CStr(Worksheets("Details").Cells(det_row, det_column).Value), 4)
            att_column = practiceNo + 2
            
            For att_column = 3 To practiceNo + 2
                If Mid(serial, att_column - 2, 1) = "0" Then
                    Worksheets("Attendance").Cells(att_row, att_column).Value = ""
                ElseIf Mid(serial, att_column - 2, 1) = "1" Then
                    Worksheets("Attendance").Cells(att_row, att_column).Value = "Y"
                ElseIf Mid(serial, att_column - 2, 1) = "2" Then
                    Worksheets("Attendance").Cells(att_row, att_column).Value = "N"
                Else
                    Worksheets("Attendance").Cells(att_row, att_column).Value = "?"
                End If
                
            Next att_column
            
            det_row = det_row + 1
        Next att_row
    
        
        Application.StatusBar = False
        Application.ScreenUpdating = True
    
        Call UpdateAttendanceList(False)
    End If
End Sub
Sub PositionAttendanceColomnButtons_v1(Optional ByVal colomn As Integer = 0)
    If colomn < 1 Then
        With Worksheets("Attendance")
            .addDate_Button.Left = .Cells(2, .Cells(1, 2).Value + 4).Left - 15
            .addDate_Button.Top = .addDate_Button.TopLeftCell.Top
            .removeDate_Button.Left = .Cells(2, .Cells(1, 2).Value + 3).Left
            .removeDate_Button.Top = .removeDate_Button.TopLeftCell.Top
        End With
    Else
        With Worksheets("Attendance")
            .addDate_Button.Left = .Cells(2, colomn + 1).Left - 15
            .addDate_Button.Top = .addDate_Button.TopLeftCell.Top
            .removeDate_Button.Left = .Cells(2, colomn).Left
            .removeDate_Button.Top = .removeDate_Button.TopLeftCell.Top
        End With
    End If
End Sub



