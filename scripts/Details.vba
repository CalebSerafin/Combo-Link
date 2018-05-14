Option Explicit
Public Function CountMembers_v1()
    Dim CachedMembers As Integer
    Dim RowMin As Integer
    Dim BottomAddr As String
    Dim initTest As Range
    CachedMembers = Int(Worksheets("COMPUTING DON'T TOUCH").Range("J20"))
    RowMin = CachedMembers + 3
    BottomAddr = "B" & RowMin - 2 & ":B" & RowMin
    Set initTest = Worksheets("Details").Range(BottomAddr)
    
    If Not ((initTest(1, 1) <> "" And initTest(2, 1) = "" And initTest(3, 1) = "")) Then
'#############################################################
Dim lastCalcValue As Long: lastCalcValue = Calculations_Off '#
'Call Calculations_On(lastCalcValue):lastCalcValue = 0 ''''''#
'#############################################################
        Dim Row As Integer
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
Function FindMember_v1(ByVal firstName As String, ByVal lastName As String, Optional ByVal matchCase As Boolean = True) 'returns row of member. returns 0 if nor found
    FindMember_v1 = 0
    
    Dim notWorking As String
    notWorking = True
    Dim fullName As String
    Dim Row As Integer
    Dim checkName As String
    
    fullName = (firstName & lastName)
    If Not matchCase Then
        fullName = LCase(fullName)
    End If
    
    For Row = 2 To maxMembers + 1
        checkName = Worksheets("Details").Cells(Row, 1).Value & Worksheets("Details").Cells(Row, 2).Value
        If Not matchCase Then
            checkName = LCase(checkName)
        End If
        If (checkName = fullName) Then
            FindMember_v1 = Row
            notWorking = False
            Exit For
        ElseIf Row = maxMembers + 1 Then
            notWorking = False
            Exit For
        End If
    Next
    
    If notWorking Then
        Debug_msg ("WARNING. Module1: FindMember_v1 is not working!")
    End If
End Function