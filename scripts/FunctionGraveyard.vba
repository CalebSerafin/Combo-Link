'===============================================================
'Outdated Functions and Subs
'Only used while updating infrastructure to provide 99% Up-Time.
'
'===============================================================
Option Explicit
Sub AttendanceData_save_v1() 'number version
    Debug_msg ("Module1 Ln11: Application.EnableEvents = " & Application.EnableEvents)
    If Application.EnableEvents <> True Then
        Debug_msg ("Module1 Ln13: AttendanceData_save()")
        Debug_msg ("Module1 Ln14: Application.EnableEvents = " & Application.EnableEvents)
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
        Dim serial As Long
        serial = 0
        
        For att_row = 3 To 37
            
            serial = 0
            For att_column = 3 To practiceNo + 2
                If IsEmpty(Worksheets("Attendance").Cells(att_row, att_column).Value) Then serial = serial + 0 * 4 ^ (att_column - 3)
                If Worksheets("Attendance").Cells(att_row, att_column).Value = "Y" Then serial = serial + 1 * 4 ^ (att_column - 3)
                If Worksheets("Attendance").Cells(att_row, att_column).Value = "N" Then serial = serial + 2 * 4 ^ (att_column - 3)
                If Worksheets("Attendance").Cells(att_row, att_column).Value = "?" Then serial = serial + 3 * 4 ^ (att_column - 3)
            Next att_column
            
            Worksheets("Details").Cells(det_row, det_column).Value = serial
            det_row = det_row + 1
        Next att_row
        
    End If
End Sub
Sub AttendanceData_load_v1() 'Number Version
    If AttendanceSaving <> True And Application.EnableEvents <> True Then
        Debug_msg ("Module1 Ln73: AttendanceSaving: " & AttendanceSaving)
        Dim practiceNo As Integer
        practiceNo = Worksheets("Attendance").Cells(1, 2).Value
        Dim serial As Long
        serial = 0
    
        Dim att_row As Integer
        att_row = 3
        Dim att_column As Integer
        att_column = 3
        
        Dim det_row As Integer
        det_row = 2
        Dim det_column As Integer
        det_column = 8
        
        For att_row = 3 To 37
            
            serial = Worksheets("Details").Cells(det_row, det_column).Value
            att_column = practiceNo + 2
            
            For att_column = practiceNo + 2 To 3 Step -1
                If serial - (3 * 4 ^ (att_column - 3)) >= 0 Then
                    Worksheets("Attendance").Cells(att_row, att_column).Value = "?"
                    serial = serial - (3 * 4 ^ (att_column - 3))
                ElseIf serial - (2 * 4 ^ (att_column - 3)) >= 0 Then
                    Worksheets("Attendance").Cells(att_row, att_column).Value = "N"
                    serial = serial - (2 * 4 ^ (att_column - 3))
                ElseIf serial - (1 * 4 ^ (att_column - 3)) >= 0 Then
                    Worksheets("Attendance").Cells(att_row, att_column).Value = "Y"
                    serial = serial - (1 * 4 ^ (att_column - 3))
                Else
                    Worksheets("Attendance").Cells(att_row, att_column).Value = ""
                End If
            Next att_column
            
            det_row = det_row + 1
        Next att_row
    
        Call UpdateAttendanceList
    End If
End Sub
Sub UpdateAttendanceList_v1(Optional ByVal save As Boolean = True)
'#############################################################
Dim lastCalcValue As Long: lastCalcValue = Calculations_Off '#
'Call Calculations_On(lastCalcValue):lastCalcValue = 0 ''''''#
'#############################################################
    Debug_msg ("Attendance: UpdateAttendanceList_v1 called")
        
    Dim Row As Integer
    Row = 3
    Dim column As Integer
    column = 2
    Dim practiceNo As Integer
    practiceNo = Worksheets("Attendance").Cells(1, 2).Value
    
    For Row = 3 To CountMembers + 2
        Dim sum As Integer
        sum = 0
        For column = 3 To practiceNo + 2
            If Worksheets("Attendance").Cells(Row, column).Value = "Y" Then sum = sum + 1
        Next column
        Worksheets("Attendance").Cells(Row, 2).Value = sum / practiceNo
        Worksheets("Details").Cells(Row - 1, 9).Value = sum / practiceNo
    Next Row
    
    If save = True Then
        Debug_msg ("Module1: UpdateAttendanceList_v1: proceeding with save function")
        Call AttendanceData_save
    End If
    
'#############################################################
'Dim lastCalcValue As Long:lastCalcValue = Calculations_Off '#
Call Calculations_On(lastCalcValue): lastCalcValue = 0 ''''''#
'#############################################################
End Sub