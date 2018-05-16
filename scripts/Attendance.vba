'===============================================================
'Global Functions and Subs
'
'
'===============================================================
Option Explicit
Sub AttendanceData_save_v2() 'String version 'Uncompressed
'#############################################################
Dim lastCalcValue As Long: lastCalcValue = Calculations_Off '#
'Call Calculations_On(lastCalcValue):lastCalcValue = 0 ''''''#
'#############################################################
    Debug_msg ("Module1: AttendanceData_save_v2(): Application.EnableEvents found at " & Application.EnableEvents)
    
    Debug_msg ("Module1: AttendanceData_save_v2() started")
    Dim att_row As Integer
    att_row = 3
    Dim att_column As Integer
    att_column = 3
    
    Dim det_row As Integer
    det_row = 2
    Dim det_column As Integer
    det_column = 8
    
    Dim PracticeNo As Integer
    PracticeNo = Worksheets("Attendance").Cells(1, 2).Value
    Dim serial As String
    serial = ""
    
    For att_row = 3 To CountMembers + 2
        
        serial = ""
        For att_column = 3 To PracticeNo + 2
            If IsEmpty(Worksheets("Attendance").Cells(att_row, att_column).Value) Then serial = serial & 0
            If Worksheets("Attendance").Cells(att_row, att_column).Value = "Y" Then serial = serial & 1
            If Worksheets("Attendance").Cells(att_row, att_column).Value = "N" Then serial = serial & 2
            If Worksheets("Attendance").Cells(att_row, att_column).Value = "?" Then serial = serial & 3
        Next att_column
        
        Worksheets("Details").Cells(det_row, det_column).Value = "v2_" & serial
        det_row = det_row + 1
    Next att_row
    
'#############################################################
'Dim lastCalcValue As Long:lastCalcValue = Calculations_Off '#
Call Calculations_On(lastCalcValue): lastCalcValue = 0 ''''''#
'#############################################################
End Sub
Function testIntMax(ByVal InputInt As Long) As Long '[!]max is 32767
    Dim MrInt As Integer
    MrInt = InputInt
    testIntMax = MrInt
End Function
Function testLongMax(ByVal InputInt As Long) As Long '[!]max is 2147483647
    Dim MrLong As Long
    MrLong = InputInt
    testLongMax = MrLong
End Function
Function TestforZero()
    Dim i As Integer
    Dim out As Integer
    out = 0
    For i = 1 To 0
        out = out + 1
    Next i
    TestforZero = out
End Function
Sub UpdateAttendanceList_v2(Optional ByVal save As Boolean = True)
'#############################################################
Dim lastCalcValue As Long: lastCalcValue = Calculations_Off '#
'Call Calculations_On(lastCalcValue):lastCalcValue = 0 ''''''#
'#############################################################
    Debug_msg ("Attendance: UpdateAttendanceList_v2 called")
        
                 'Complete and implement this function if overflow occurs on variant
'#############################################################################################################
'    Dim CurrentColumn As Long 'Colomn we are buzy with
'    Dim CurrentRow As Long 'Row we are buzy with
'    Dim CurrentColomnRepeat As Integer 'What group we are on            '[!]max is 32767
'    Dim CurrentRowRepeat As Integer 'What group we are on               '[!]max is 32767
'
'    Dim AttendanceRange As String 'Range for AttendanceCells
'    Dim AttendanceCells As Range 'Where we load our data                '[!]max is 30000 cells
'    Dim SummeryPercentRange As Range 'Where we paste out data            '[!]max is 30000 cells
'    Dim PracticeNo As Long 'How many entries we are working with
'    Dim RowSum As Integer 'Current sum for that row
'
'    Call SummeryPercentRange.Resize(CountMembers, 1)
'    PracticeNo = CInt(Worksheets("Attendance").Cells(1, 2).Value)
'
'    Dim ColumnBasic As Integer 'Normal amount of columns loaded         '[!]max is 32767
'    Dim ColumnRepeats As Integer 'How many extra 30000 columns to load  '[!]max is 32767
'    Dim ColumnExtra As Integer 'reminder amount of last columns loaded  '[!]max is 32767
'    Dim RowBasic As Integer 'Maximium amount of rows loaded             '[!]max is 32767
'    Dim RowRepeats As Integer 'How many extra 30000 rows to load        '[!]max is 32767
'    Dim RowExtra As Integer 'reminder amount of last rows loaded        '[!]max is 32767
'
'
'                                                                        '[!]Requires thinking
'    If (30000 / PracticeNo) >= 1 Then 'If we can fit the whole row into the 30000 cell range
'        If (30000 / (PracticeNo * CountMembers)) <= 1 Then 'If we can fit everything into the 30000 cell range
'            ColumnBasic = PracticeNo
'            ColumnRepeats = 0
'            RowExtra = CountMembers
'            RowBasic = CountMembers
'            RowRepeats = 0
'        Else 'If we need to split the rows into groups
'            ColumnBasic = PracticeNo
'            ColumnRepeats = 0
'            RowBasic = Floor(30000 / ColumnBasic)
'            RowRepeats = Floor(CountMembers / RowBasic)
'            RowExtra = CountMembers Mod RowBasic
'        End If
'    Else 'If we need to split columns and rows into groups
'        ColumnBasic = 30000
'        ColumnRepeats = Floor(PracticeNo / ColumnBasic)
'        ColumnExtra = PracticeNo Mod ColumnBasic
'        RowBasic = 1
'        RowRepeats = 1
'        RowExtra = 0
'    End If
'
'    CurrentRow = 1
'    For CurrentColomnRepeat = 1 To ColumnRepeats
'        AttendanceRange = Range(Cells(CurrentRow, 30000 * (CurrentColomnRepeat - 1) + 1), Cells(CurrentRow, 30000 * CurrentColomnRepeat)).Address
'        Set AttendanceCells = Worksheets("Attendance").Range(AttendanceRange)
'    Next CurrentColomnRepeat
'#############################################################################################################

    

    Dim Row As Long 'Internal row starting at 1 going to CountMembers
    Dim Column As Long 'Internal column starting at 1 going to PracticeNo
    Dim SummeryPercentRange As Variant 'Where we paste out data            '[!]max is IDK
    Dim PracticeNo As Long 'How many entries we are working with
    Dim AttendanceCells As Variant 'Where we load our data                                  '[!]max is IDK
    Dim RowSum As Long 'Current sum for that row
    
    ReDim SummeryPercentRange(CountMembers - 1, 0)
    PracticeNo = CInt(Worksheets("Attendance").Cells(1, 2).Value)
    With Worksheets("Attendance")
        AttendanceCells = .Range(.Cells(3, 3), .Cells(CountMembers + 2, PracticeNo + 2))
    End With
    RowSum = 0

    For Row = 1 To CountMembers
        RowSum = 0
        For Column = 1 To PracticeNo
            If AttendanceCells(Row, Column) = "Y" Then RowSum = RowSum + 1
        Next Column
        SummeryPercentRange(Row - 1, 0) = CStr(RowSum / PracticeNo)
    Next Row
    
    With Worksheets("Attendance")
        .Range(.Cells(3, 2), .Cells(CountMembers + 2, 2)) = SummeryPercentRange
    End With
    With Worksheets("Details")
        .Range(.Cells(2, 9), .Cells(CountMembers + 1, 9)) = SummeryPercentRange
    End With
    
    If save = True Then
        Debug_msg ("Module1: UpdateAttendanceList_v1: proceeding with save function")
        Call AttendanceData_save
    End If
    
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
    
    For att_row = 3 To CountMembers + 2
        Application.StatusBar = "Please Wait ... Syncing Attendance List: " & att_row - 3 & "/" & maxMembers
        
        serial = Mid(CStr(Worksheets("Details").Cells(det_row, det_column).Value), 4)
        att_column = PracticeNo + 2
        
        For att_column = 3 To PracticeNo + 2
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
    Call UpdateAttendanceList(False)
'#############################################################
'Dim lastCalcValue As Long:lastCalcValue = Calculations_Off '#
Call Calculations_On(lastCalcValue): lastCalcValue = 0 ''''''#
'#############################################################
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



