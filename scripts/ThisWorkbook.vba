Option Explicit
'\\\Initializers///'
Private Sub Workbook_Open()
    '///removes enable Macro Rectangle
    If Worksheets("COMPUTING DON'T TOUCH").Range("F26").Value = "Y" Then
        '///removes broken references
        Call References_RemoveMissing '[!]''Undefined behavour in non-programmatically allowed macro
        '///removes enable Macro Rectangle
        On Error Resume Next
        Worksheets("Details").Shapes("Rectangle 1").Delete
    End If
    '///maxMembers
    dataTableOld = "nil"
    AttendanceSaving = False
    maxMembers = Int(Worksheets("COMPUTING DON'T TOUCH").Cells(15, 6).Value)

    Application.EnableEvents = False '===== Basically Refresh to get filter buttons to werk
    Call AttendanceData_load
    Application.EnableEvents = True
    '///Version Number
    Worksheets("COMPUTING DON'T TOUCH").Range("F20").Value = "1.1"
    
    
    
    '///First Time Opened   <--- Put last
    Worksheets("COMPUTING DON'T TOUCH").Range("F26").Value = "N"
End Sub
'//////// \\\\\\\\\'

