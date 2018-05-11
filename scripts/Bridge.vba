'===============================================================
'Function and Sub Linker
'Used to switch between different versions of functions while
'not refactoring entire projects. Providing 99% Up-Time.
'===============================================================
'\\\Debug Function///'
Function Debug_msg(ByVal msg As String, Optional ByVal code As String = "null", Optional ByVal request As String = "null") As Boolean '====== DO NOT put 'Call IREC' in here!
    Debug_msg = Debug_msg_v1(msg, code, request)
End Function
'/////////  \\\\\\\\\'

'\\\Initializer Runtime Error Check///' 'Please call in your functions to allow checking
Public Sub IREC()
    If (maxMembers = 0 And Int(Worksheets("COMPUTING DON'T TOUCH").Cells(15, 6).Value)) <> 0 Or maxMembers = Null Then
        Call Debug_msg("WARNING. IREC detected maxMembers value defect, reloading from cell value", "IREC")
        maxMembers = Int(Worksheets("COMPUTING DON'T TOUCH").Cells(15, 6).Value)
        
        If (maxMembers = 0 And Int(Worksheets("COMPUTING DON'T TOUCH").Cells(15, 6).Value)) <> 0 Or maxMembers = Null Then
             Call Debug_msg("WARNING. IREC detected maxMembers was unable to load from 'Worksheets(""COMPUTING DON'T TOUCH"").Cells(15, 6).Value', using defualt value of '64'.", "IREC")
             maxMembers = 64
             
             If (maxMembers = 0 And Int(Worksheets("COMPUTING DON'T TOUCH").Cells(15, 6).Value)) <> 0 Or maxMembers = Null Then
            Call Debug_msg("CRITICAL ERROR. IREC was unable to fix maxMembers recurring value defect. Please alert developer to check through the code for any functions modifing maxMembers.", "IREC")
            Call Debug_msg("CRITICAL ERROR. IREC set Application.EnableEvents to false and set AttendanceSaving to true to halt all macros.", "IREC")
            MsgBox "CRITICAL ERROR. IREC was unable to fix maxMembers recurring value defect. Please alert developer to check through the code for any functions modifing maxMembers.", vbCritical, "IREC"
            Application.EnableEvents = False
            AttendanceSaving = True
            MsgBox "CRITICAL ERROR. IREC has automatically set 'Application.EnableEvents' to 'False' and automatically set 'AttendanceSaving' to 'True' to halt all macros. YOU DO THIS NOW: (1)Please save and exit as soon as possible to minimise data damage. (2)If your computer freezes unnessisarly press both 'Alt F4' to force close Excel but you will lose unsaved data. (3)Please try not to edit the worksheet.", vbExclamation, "IREC"
            End If
        End If
    End If
End Sub
'////////////////// \\\\\\\\\\\\\\\\\\'
'\\\Module1///'
Function StringMult(ByVal Word As String, ByVal Multiply As Integer) As String
    StringMult = StringMult_v1(Word, Int(Multiply))
End Function
Function addCellData(ByVal mode As String, ByVal sheet As String, ByVal min As Integer, ByVal max As Integer, ByVal rawData As String, ByVal topLeft As Integer, ByVal forceLast As Boolean)
    Call addCellData_v1(mode, sheet, min, max, rawData, topLeft, forceLast)
End Function
Sub AttendanceData_save()
    Call IREC
    Call AttendanceData_save_v2
End Sub
Sub UpdateAttendanceList(Optional ByVal save As Boolean = True)
    Call UpdateAttendanceList_v1(save)
End Sub
Sub AttendanceData_load()
    Call IREC
    Call AttendanceData_load_v2
End Sub
Sub ScanCommonError()
    Call IREC
    Call ScanCommonError_v1
End Sub
Sub PositionAttendanceColomnButtons(Optional ByVal colomn As Integer = 0)
    Call PositionAttendanceColomnButtons_v1(colomn)
End Sub
Function GetMonth(ByVal number As Integer) As String
    GetMonth = GetMonth_v1(number)
End Function
Function FindMember(ByVal firstName As String, ByVal lastName As String, Optional ByVal matchCase As Boolean = True)
    Call IREC
    FindMember = FindMember_v1(firstName, lastName, matchCase)
End Function
Sub References_RemoveMissing()
    Call References_RemoveMissing_v1
End Sub
'/////  \\\\\\'