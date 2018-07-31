Option Explicit
Public Function CountMembers_v2() As Long
    Dim CachedMembers As Long
    Dim RowMin As Long
    Dim BottomAddr As String
    Dim initTest As Variant
    CachedMembers = CLng(Worksheets("COMPUTING DON'T TOUCH").Range("J20").Value)
    RowMin = CachedMembers + 3
    BottomAddr = "B" & RowMin - 2 & ":B" & RowMin
    initTest = Worksheets("Details").Range(BottomAddr)
    
'#############################################################
Dim lastCalcValue As Long: lastCalcValue = Calculations_Off '#
'Call Calculations_On(lastCalcValue):lastCalcValue = 0 ''''''#
'#############################################################
    Do Until ((initTest(2, 1) = "" And initTest(3, 1) = "")) 'initTest(1, 1) <> "" And
        RowMin = CachedMembers + 3
        BottomAddr = "B" & RowMin - 2 & ":B" & RowMin
        initTest = Worksheets("Details").Range(BottomAddr)
    
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
    Loop
'#############################################################
'Dim lastCalcValue As Long:lastCalcValue = Calculations_Off '#
Call Calculations_On(lastCalcValue): lastCalcValue = 0 ''''''#
'#############################################################
    CountMembers_v2 = CachedMembers
End Function

Function FindMember_v1(ByVal firstName As String, ByVal lastName As String, Optional ByVal matchCase As Boolean = True) 'returns row of member. returns 0 if nor found
    FindMember_v1 = 0
    
    Dim notWorking As String
    notWorking = True
    Dim fullName As String
    Dim Row As Long
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