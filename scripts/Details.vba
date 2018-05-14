Option Explicit
Function FindMember_v1(ByVal firstName As String, ByVal lastName As String, Optional ByVal matchCase As Boolean = True) 'returns row of member. returns 0 if nor found
    FindMember_v1 = 0
    
    Dim notWorking As String
    notWorking = True
    Dim fullName As String
    Dim row As Integer
    Dim checkName As String
    
    fullName = (firstName & lastName)
    If Not matchCase Then
        fullName = LCase(fullName)
    End If
    
    For row = 2 To maxMembers + 1
        checkName = Worksheets("Details").Cells(row, 1).Value & Worksheets("Details").Cells(row, 2).Value
        If Not matchCase Then
            checkName = LCase(checkName)
        End If
        If (checkName = fullName) Then
            FindMember_v1 = row
            notWorking = False
            Exit For
        ElseIf row = maxMembers + 1 Then
            notWorking = False
            Exit For
        End If
    Next
    
    If notWorking Then
        Debug_msg ("WARNING. Module1: FindMember_v1 is not working!")
    End If
End Function