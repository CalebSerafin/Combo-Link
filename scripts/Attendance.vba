'===============================================================
'Global Functions and Subs
'
'
'===============================================================
Option Explicit
Sub AttendanceData_save_v3() 'String version 'Uncompressed
'#############################################################
Dim lastCalcValue As Long: lastCalcValue = Calculations_Off '#
'Call Calculations_On(lastCalcValue):lastCalcValue = 0 ''''''#
'#############################################################
    Debug_msg ("Module1: AttendanceData_save_v3() started")
    
    Dim PracticeNo As Long
    Dim AmountMembers As Long
    Dim Column As Long
    Dim Row As Long
    Dim CurrentAttData As String    'Data from individual string from AttendanceData
    Dim Serial As String    'Data stored in ATTENDANCE DATA COLOMN in details
    Dim SerialAmmend As String  'Whats added on to serail
    
    PracticeNo = Worksheets("Attendance").Cells(1, 2).Value
    AmountMembers = CountMembers
    Serial = ""
    
    Dim AttendanceData As Variant   'Range of cells from Attendance that represent the attendance marking
    Dim DetailsStorage As Variant   'Where the serialized data goes
    
    With Worksheets("Attendance")
        AttendanceData = .Range(.Cells(3, 3), .Cells(AmountMembers + 2, PracticeNo + 2))
    End With
    ReDim DetailsStorage(1 To AmountMembers, 1 To 1)
    
    For Row = 1 To AmountMembers
        
        Serial = ""
        For Column = 1 To PracticeNo
            CurrentAttData = AttendanceData(Row, Column)
            
            SerialAmmend = "0"  'Default if no other values are found
            If CurrentAttData = "Y" Then SerialAmmend = "1"
            If CurrentAttData = "N" Then SerialAmmend = "2"
            If CurrentAttData = "?" Then SerialAmmend = "3"
            
            Serial = Serial & SerialAmmend
        Next Column
        
        DetailsStorage(Row, 1) = "v2_" & Serial 'seraial version does not always have to match function version,
                                                          '^^^only if the the serialization protocall changes, ie: moving from plain to compressed data
    Next Row
    
    With Worksheets("Details")
        .Range(.Cells(2, 8), .Cells(AmountMembers + 1, 8)).Value = DetailsStorage
    End With
'#############################################################
'Dim lastCalcValue As Long:lastCalcValue = Calculations_Off '#
Call Calculations_On(lastCalcValue): lastCalcValue = 0 ''''''#
'#############################################################
End Sub
Sub FormatColor()
    Dim UpperRow As Long
    Dim UpperColumn As Long
    
    UpperRow = CountMembers + 2
    UpperColumn = CLng(Worksheets("Attendance").Range("B1").Value) + 2
    
    'Found at:
    'http://www.bluepecantraining.com/portfolio/excel-vba-macro-to-apply-conditional-formatting-based-on-value/
    Dim rg As Range
    With Worksheets("Attendance")
        Set rg = .Range(.Cells(3, 3), .Cells(UpperRow, UpperColumn))
    End With
    Dim i As Long
    Dim c As Long
    Dim testcell As Range
    c = rg.Cells.Count
     
    For i = 1 To c
    Set testcell = rg(i)
    Select Case testcell
        Case Is = ""
            With testcell
                .Interior.ColorIndex = 0
            End With
        Case Is = " "
            With testcell
                .Interior.ColorIndex = 0
                .Value = ""
            End With
        Case Is = "Y"
            With testcell
                .Interior.Color = RGB(112, 173, 71)
            End With
        Case Is = "N"
            With testcell
                .Interior.Color = RGB(237, 125, 49)
            End With
        Case Is = "?"
            With testcell
                .Interior.Color = RGB(255, 192, 0)
             End With
        End Select
    Next i
    
    
End Sub
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
    Dim MemberListNames As Variant 'List of names to post to the Attendance sheet       '[!]max is IDK
    Dim PracticeNo As Long 'How many entries we are working with
    Dim AmountMembers As Long   'how many members there are
    Dim AttendanceCells As Variant 'Where we load our data                                  '[!]max is IDK
    Dim RowSum As Long 'Current sum for that row
    
    AmountMembers = CountMembers 'a function
    ReDim SummeryPercentRange(1 To AmountMembers, 1 To 1)
    PracticeNo = CLng(Worksheets("Attendance").Cells(1, 2).Value)
    
    Call FormatColor
    
    If PracticeNo > 0 Then 'Checks if the practice number includes actual day(s)and not zero.
        With Worksheets("Attendance")
            AttendanceCells = .Range(.Cells(3, 3), .Cells(AmountMembers + 2, PracticeNo + 2))
        End With
        RowSum = 0
    
        For Row = 1 To AmountMembers
            RowSum = 0
            For Column = 1 To PracticeNo
                If AttendanceCells(Row, Column) = "Y" Then RowSum = RowSum + 1
            Next Column
            SummeryPercentRange(Row, 1) = CStr(Round(RowSum / PracticeNo, 5))
        Next Row
    Else    'if 0 (or somehow lower)
        SummeryPercentRange = "1"
    End If
    
    ReDim MemberListNames(1 To AmountMembers + 2, 1 To 1) As Variant
    MemberListNames = Application.Transpose(JoinDetailNames)
    
    With Worksheets("Details")
        'MemberListNames = .Range(.Cells(2, 2), .Cells(AmountMembers + 1, 2))
        .Range(.Cells(2, 9), .Cells(AmountMembers + 1, 9)) = SummeryPercentRange
    End With
    With Worksheets("Attendance")
        .Range(.Cells(3, 2), .Cells(AmountMembers + 2, 2)) = SummeryPercentRange
        .Range(.Cells(3, 1), .Cells(AmountMembers + 2, 1)) = MemberListNames
    End With
    
    If save = True Then
        Debug_msg ("Module1: UpdateAttendanceList_v2: proceeding with save function")
        Call AttendanceData_save
    End If
    
'#############################################################
'Dim lastCalcValue As Long:lastCalcValue = Calculations_Off '#
Call Calculations_On(lastCalcValue): lastCalcValue = 0 ''''''#
'#############################################################
End Sub
Sub AttendanceData_load_v3() 'String version 'Uncompressed
'#############################################################
Dim lastCalcValue As Long: lastCalcValue = Calculations_Off '#
'Call Calculations_On(lastCalcValue):lastCalcValue = 0 ''''''#
'#############################################################
    Debug_msg ("Module1: AttendanceData_load_v3() started")
    Application.StatusBar = "Please Wait ... Syncing Attendance List: "
    
    Dim PracticeNo As Long
    Dim AmountMembers As Long
    Dim Column As Long
    Dim Row As Long
    Dim CurrentSerialChar As String    'Data from individual string from AttendanceData
    Dim Serial As String    'Individual Data stored in ATTENDANCE DATA COLOMN in details
    Dim CurrentAttChar As String    'Data for individual char for AttendanceData
    Dim CurrentAttColour As Long  'Colour for viewing
    
    PracticeNo = Worksheets("Attendance").Cells(1, 2).Value
    AmountMembers = CountMembers
    
    If PracticeNo > 0 Then 'Checks if Practice Number is more than 0
        Dim DetailsStorage As Variant   'Where the serialized data is from
        Dim AttendanceData As Variant   'Range of cells from Attendance that represent the attendance marking
        
        With Worksheets("Details")
            DetailsStorage = .Range(.Cells(2, 8), .Cells(AmountMembers + 1, 8))
        End With
        ReDim AttendanceData(1 To AmountMembers, 1 To PracticeNo)
        
        If Mid(CStr(DetailsStorage(1, 1)), 2, 1) <> "2" Then  'checking serial version
            On Error GoTo IncompatibleVersion
            Err.Raise 1, "AttendanceData_load_v3", ("AttendanceData_load_v3 is incompatible with Version " & Mid(CStr(DetailsStorage(1, 1)), 2, 1) & " serails. Please update your Combo-Link to the latest version!")
        End If
        
        
        For Row = 1 To AmountMembers
            Serial = Mid(CStr(DetailsStorage(Row, 1)), 4)
            For Column = 1 To PracticeNo
                CurrentSerialChar = Mid(Serial, Column, 1)
                
                CurrentAttChar = " " 'Default if no other values found.
                If CurrentSerialChar = "1" Then CurrentAttChar = "Y"
                If CurrentSerialChar = "2" Then CurrentAttChar = "N"
                If CurrentSerialChar = "3" Then CurrentAttChar = "?"
                
                AttendanceData(Row, Column) = CurrentAttChar
            Next Column
        Next Row
        
        
        With Worksheets("Attendance")
            .Range(.Cells(3, 3), .Cells(AmountMembers + 2, PracticeNo + 2)).Value = AttendanceData
        End With
    End If
    
    Application.StatusBar = False
    Call UpdateAttendanceList(False)
'#############################################################
'Dim lastCalcValue As Long:lastCalcValue = Calculations_Off '#
Call Calculations_On(lastCalcValue): lastCalcValue = 0 ''''''#
'#############################################################
Exit Sub
IncompatibleVersion:
    Call Debug_msg(Err.source & ": " & Err.Description, "Incompatible", "Notify")
    Err.Clear
End Sub
Sub PositionAttendanceColomnButtons_v1(Optional ByVal colomn As Long = 0)
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



