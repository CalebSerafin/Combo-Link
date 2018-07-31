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
        Dim Att_Row As Long
        Att_Row = 3
        Dim Att_Column As Long
        Att_Column = 3
        
        Dim Det_Row As Long
        Det_Row = 2
        Dim Det_Column As Integer
        Det_Column = 8
        
        Dim PracticeNo As Integer
        PracticeNo = Worksheets("Attendance").Cells(1, 2).Value
        Dim Serial As Long
        Serial = 0
        
        For Att_Row = 3 To 37
            
            Serial = 0
            For Att_Column = 3 To PracticeNo + 2
                If IsEmpty(Worksheets("Attendance").Cells(Att_Row, Att_Column).Value) Then Serial = Serial + 0 * 4 ^ (Att_Column - 3)
                If Worksheets("Attendance").Cells(Att_Row, Att_Column).Value = "Y" Then Serial = Serial + 1 * 4 ^ (Att_Column - 3)
                If Worksheets("Attendance").Cells(Att_Row, Att_Column).Value = "N" Then Serial = Serial + 2 * 4 ^ (Att_Column - 3)
                If Worksheets("Attendance").Cells(Att_Row, Att_Column).Value = "?" Then Serial = Serial + 3 * 4 ^ (Att_Column - 3)
            Next Att_Column
            
            Worksheets("Details").Cells(Det_Row, Det_Column).Value = Serial
            Det_Row = Det_Row + 1
        Next Att_Row
        
    End If
End Sub
Sub AttendanceData_load_v1() 'Number Version
    If AttendanceSaving <> True And Application.EnableEvents <> True Then
        Debug_msg ("Module1 Ln73: AttendanceSaving: " & AttendanceSaving)
        Dim PracticeNo As Integer
        PracticeNo = Worksheets("Attendance").Cells(1, 2).Value
        Dim Serial As Long
        Serial = 0
    
        Dim Att_Row As Integer
        Att_Row = 3
        Dim Att_Column As Integer
        Att_Column = 3
        
        Dim Det_Row As Integer
        Det_Row = 2
        Dim Det_Column As Integer
        Det_Column = 8
        
        For Att_Row = 3 To 37
            
            Serial = Worksheets("Details").Cells(Det_Row, Det_Column).Value
            Att_Column = PracticeNo + 2
            
            For Att_Column = PracticeNo + 2 To 3 Step -1
                If Serial - (3 * 4 ^ (Att_Column - 3)) >= 0 Then
                    Worksheets("Attendance").Cells(Att_Row, Att_Column).Value = "?"
                    Serial = Serial - (3 * 4 ^ (Att_Column - 3))
                ElseIf Serial - (2 * 4 ^ (Att_Column - 3)) >= 0 Then
                    Worksheets("Attendance").Cells(Att_Row, Att_Column).Value = "N"
                    Serial = Serial - (2 * 4 ^ (Att_Column - 3))
                ElseIf Serial - (1 * 4 ^ (Att_Column - 3)) >= 0 Then
                    Worksheets("Attendance").Cells(Att_Row, Att_Column).Value = "Y"
                    Serial = Serial - (1 * 4 ^ (Att_Column - 3))
                Else
                    Worksheets("Attendance").Cells(Att_Row, Att_Column).Value = ""
                End If
            Next Att_Column
            
            Det_Row = Det_Row + 1
        Next Att_Row
    
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
    Dim Column As Integer
    Column = 2
    Dim PracticeNo As Integer
    PracticeNo = Worksheets("Attendance").Cells(1, 2).Value
    
    For Row = 3 To CountMembers + 2
        Dim sum As Integer
        sum = 0
        For Column = 3 To PracticeNo + 2
            If Worksheets("Attendance").Cells(Row, Column).Value = "Y" Then sum = sum + 1
        Next Column
        Worksheets("Attendance").Cells(Row, 2).Value = sum / PracticeNo
        Worksheets("Details").Cells(Row - 1, 9).Value = sum / PracticeNo
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
Sub AttendanceData_save_v2() 'String version 'Uncompressed
'#############################################################
Dim lastCalcValue As Long: lastCalcValue = Calculations_Off '#
'Call Calculations_On(lastCalcValue):lastCalcValue = 0 ''''''#
'#############################################################
    Debug_msg ("Module1: AttendanceData_save_v2(): Application.EnableEvents found at " & Application.EnableEvents)
    
    Debug_msg ("Module1: AttendanceData_save_v2() started")
    Dim Att_Row As Integer
    Att_Row = 3
    Dim Att_Column As Integer
    Att_Column = 3
    
    Dim Det_Row As Integer
    Det_Row = 2
    Dim Det_Column As Integer
    Det_Column = 8
    
    Dim PracticeNo As Integer
    PracticeNo = Worksheets("Attendance").Cells(1, 2).Value
    Dim Serial As String
    Serial = ""
    
    For Att_Row = 3 To CountMembers + 2
        
        Serial = ""
        For Att_Column = 3 To PracticeNo + 2
            If IsEmpty(Worksheets("Attendance").Cells(Att_Row, Att_Column).Value) Then Serial = Serial & 0
            If Worksheets("Attendance").Cells(Att_Row, Att_Column).Value = "Y" Then Serial = Serial & 1
            If Worksheets("Attendance").Cells(Att_Row, Att_Column).Value = "N" Then Serial = Serial & 2
            If Worksheets("Attendance").Cells(Att_Row, Att_Column).Value = "?" Then Serial = Serial & 3
        Next Att_Column
        
        Worksheets("Details").Cells(Det_Row, Det_Column).Value = "v2_" & Serial
        Det_Row = Det_Row + 1
    Next Att_Row
    
'#############################################################
'Dim lastCalcValue As Long:lastCalcValue = Calculations_Off '#
Call Calculations_On(lastCalcValue): lastCalcValue = 0 ''''''#
'#############################################################
End Sub
Sub AttendanceData_load_v2() 'String version 'Uncompressed
'#############################################################
Dim lastCalcValue As Long: lastCalcValue = Calculations_Off '#
'Call Calculations_On(lastCalcValue):lastCalcValue = 0 ''''''#
'#############################################################
    Application.StatusBar = "Please Wait ... Syncing Attendance List: "
    Dim PracticeNo As Integer
    PracticeNo = Worksheets("Attendance").Cells(1, 2).Value
    Dim Serial As String
    Serial = ""

    Dim Att_Row As Integer
    Att_Row = 3
    Dim Att_Column As Integer
    Att_Column = 3
    
    Dim Det_Row As Integer
    Det_Row = 2
    Dim Det_Column As Integer
    Det_Column = 8
    
    For Att_Row = 3 To CountMembers + 2
        Application.StatusBar = "Please Wait ... Syncing Attendance List: " & Att_Row - 3 & "/" & maxMembers
        
        Serial = Mid(CStr(Worksheets("Details").Cells(Det_Row, Det_Column).Value), 4)
        Att_Column = PracticeNo + 2
        
        For Att_Column = 3 To PracticeNo + 2
            If Mid(Serial, Att_Column - 2, 1) = "0" Then
                Worksheets("Attendance").Cells(Att_Row, Att_Column).Value = ""
            ElseIf Mid(Serial, Att_Column - 2, 1) = "1" Then
                Worksheets("Attendance").Cells(Att_Row, Att_Column).Value = "Y"
            ElseIf Mid(Serial, Att_Column - 2, 1) = "2" Then
                Worksheets("Attendance").Cells(Att_Row, Att_Column).Value = "N"
            Else
                Worksheets("Attendance").Cells(Att_Row, Att_Column).Value = "?"
            End If
            
        Next Att_Column
        
        Det_Row = Det_Row + 1
    Next Att_Row

    
    Application.StatusBar = False
    Call UpdateAttendanceList(False)
'#############################################################
'Dim lastCalcValue As Long:lastCalcValue = Calculations_Off '#
Call Calculations_On(lastCalcValue): lastCalcValue = 0 ''''''#
'#############################################################
End Sub

Public Function CountMembers_v1() As Long
    Dim CachedMembers As Long
    Dim RowMin As Long
    Dim BottomAddr As String
    Dim initTest As Variant
    CachedMembers = CLng(Worksheets("COMPUTING DON'T TOUCH").Range("J20").Value)
    RowMin = CachedMembers + 3
    BottomAddr = "B" & RowMin - 2 & ":B" & RowMin
    initTest = Worksheets("Details").Range(BottomAddr)
    
    If Not ((initTest(1, 1) <> "" And initTest(2, 1) = "" And initTest(3, 1) = "")) Then
'#############################################################
Dim lastCalcValue As Long: lastCalcValue = Calculations_Off '#
'Call Calculations_On(lastCalcValue):lastCalcValue = 0 ''''''#
'#############################################################
        Dim Row As Long
        Dim WholeAddr As String
        Dim FullRange As Range
        
        Row = 1
        WholeAddr = "B2:B" & (RowMin + 1)
        Set FullRange = Worksheets("Details").Range(WholeAddr)

        Dim TempRng As Range
        CachedMembers = CachedMembers + 4
        For Each TempRng In FullRange
            If TempRng.Value = "" Then
                CachedMembers = TempRng.Row - 2
                Exit For
            End If
        Next TempRng
        Worksheets("COMPUTING DON'T TOUCH").Range("J20") = CachedMembers
'#############################################################
'Dim lastCalcValue As Long:lastCalcValue = Calculations_Off '#
Call Calculations_On(lastCalcValue): lastCalcValue = 0 ''''''#
'#############################################################
    End If
    CountMembers_v1 = CachedMembers
End Function