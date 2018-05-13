Function Debug_msg_v1(ByVal msg As String, Optional ByVal code As String = "null", Optional ByVal request As String = "null") As Boolean
        If Worksheets("COMPUTING DON'T TOUCH").Cells(20, 2).Value = "Window" Then
            Debug_msg_v1 = True
            Debug.Print (msg)
        ElseIf Worksheets("COMPUTING DON'T TOUCH").Cells(20, 2).Value = "Pop-Up" Then
            Debug_msg_v1 = True
            Debug.Print (msg)
            MsgBox (msg)
        ElseIf Worksheets("COMPUTING DON'T TOUCH").Cells(20, 2).Value = "IREC-only" Then
            If code = "IREC" Then
                Debug_msg_v1 = True
                Debug.Print (msg)
            Else
                Debug_msg_v1 = False
            End If
        End If
        
        'Special Requests
        If request = "Notify" And Not Worksheets("COMPUTING DON'T TOUCH").Cells(20, 2).Value = "Pop-Up" Then
            Debug_msg_v1 = True
            MsgBox (msg)
        End If
End Function
Function IsInArray_v1(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray_v1 = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function
Function GetMonth_v1(ByVal number As Integer, Optional ByVal longName As Boolean = False) As String
    If IsNumeric(number) Then
        number = Int(number)
        Dim monthList() As String
        If longName Then
            monthList() = Split("Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec", ",")
        Else
            monthList() = Split("January,February,March,April,May,June,July,August,September,October,November,December", ",")
        End If
        
        GetMonth_v1 = monthList(number - 1)
    Else
        GetMonth_v1 = "NAN"
    End If
End Function
Function StringMult_v1(ByVal Word As String, ByVal Multiply As Integer) As String
    Dim repeat As Integer
    StringMult_v1 = ""
    Multiply = Int(Multiply)
    
    If Multiply >= 1 Then
        For repeat = 1 To Multiply
            StringMult_v1 = StringMult_v1 + Word
        Next repeat
    End If
End Function
Function addCellData_v1(ByVal mode As String, ByVal sheet As String, ByVal min As Integer, ByVal max As Integer, ByVal rawData As String, ByVal topLeft As Integer, ByVal forceLast As Boolean)
    
    addCellData_v1 = False
    
    If mode = "row" Or mode = "column" Then
        Dim data1D() As String
        Dim data1DUBound As Integer
        
        data1D = Split(rawData, "\'\")
        data1DUBound = UBound(data1D, 1)
    End If
    
    If mode = "row" Then
        Dim index1D As Integer
        Dim atColumn As Integer
        Dim isFree As Boolean
        Dim howLong As Integer
        If forceLast = True Then
            howLong = data1DUBound - 1
        Else
            howLong = data1DUBound
        End If
        For index1D = min To max
            isFree = True
            For atColumn = topLeft To howLong + topLeft
                If Not isEmpty(Worksheets(sheet).Cells(index1D, atColumn).Value) Then
                    isFree = False
                    Exit For
                End If
            Next atColumn
            If isFree = True Then
                For atColumn = topLeft To topLeft + data1DUBound
                    Worksheets(sheet).Cells(index1D, atColumn).Value = data1D(atColumn - topLeft)
                Next atColumn
                addCellData_v1 = True
                Exit For
            End If
        Next index1D
    End If
    
    
End Function
Sub References_RemoveMissing_v1() 'Removes missing References from VBE
    Dim theRef As Variant, i As Long
    On Error Resume Next
    
    For i = ThisWorkbook.VBProject.References.Count To 1 Step -1
        Set theRef = ThisWorkbook.VBProject.References.Item(i)
        If theRef.isBroken = True Then
            ThisWkbook.VBProject.References.Remove (theRef)
        End If
    Next i
    
    If Err <> 0 Then
        MsgBox "A missing reference has been encountered!" & "You will need to remove the reference manually.", vbCritical, "Unable To Remove Missing Reference"
    End If
End Sub
Sub ScanCommonError_v1() ' Requires v2 Load and Save functions
    If AttendanceSaving <> True Then
        Dim wasEnabled As Boolean
        wasEnabled = Application.EnableEvents
        Application.EnableEvents = False
        Application.StatusBar = "Refreshing...."
        Application.ScreenUpdating = False
        
        Dim row As Integer
        
        'row = 2
        'For row = 2 To 65 Step 1
        '    If ((Worksheets("Details").Cells(row, 1) = "Brink") And (Worksheets("Details").Cells(row, 2) = "Nelson")) Or (Worksheets("Details").Cells(row, 6) = "072 223 2173") Then
        '        Worksheets("Details").Cells(row, 5) = "Boss"
        '    End If
        'Next row
        If Worksheets("COMPUTING DON'T TOUCH").Cells(5, 12).Value = "Drama Club" Then
            For row = 2 To maxMembers + 1 Step 1
                If ((Worksheets("Details").Cells(row, 1) = "Caleb") And (Worksheets("Details").Cells(row, 2) = "Serafin")) Or (Worksheets("Details").Cells(row, 6) = "076 318 9700") Then
                    Worksheets("Details").Cells(row, 5) = "Memoral of the lost Generation"
                End If
            Next row
        End If
        '///GapRemover
        Dim index1D As Integer
        Dim atColumn As Integer
        Dim isFree As Boolean
        For index1D = 2 To maxMembers + 1
            isFree = True
            For atColumn = 1 To 7
                If Not isEmpty(Worksheets("Details").Cells(index1D, atColumn).Value) Then
                    isFree = False
                    Exit For
                End If
            Next atColumn
            If isFree = True And Worksheets("Details").Cells(index1D, 8).Value = ("v2_" & StringMult("0", Int(Worksheets("Attendance").Range("B1").Value))) Then
                Debug_msg ("Module 1: ScanCommonError_v1: Terminating GapRemover at row: " & index1D)
                Exit For
            End If
            If isFree = True Then
                Debug_msg ("Module 1: ScanCommonError_v1: Found data without details at row: " & index1D)
                Worksheets("Details").Cells(index1D, 8).Value = "v2_" & StringMult("0", Int(Worksheets("Attendance").Range("B1").Value)) '<---This is what requires v2 Load and Save functions (the v2_ part)
                Call AttendanceData_load
            End If
        Next index1D
        '///End GapRemover
        
        If Worksheets("COMPUTING DON'T TOUCH").Cells(5, 12).Value = "Drama Club" Then
            Dim found As Boolean
            found = False
            For row = 2 To maxMembers + 1 + 1 And found = False Step 1
                If ((Worksheets("Details").Cells(row, 1) = "Caleb") And (Worksheets("Details").Cells(row, 2) = "Serafin")) Or (Worksheets("Details").Cells(row, 6) = "076 318 9700") Then
                    found = True
                End If
            Next row
            If found <> True Then
                Call addCellData("row", "Details", 2, maxMembers + 1 + 1, "Caleb\'\Serafin\'\10.5\'\Ex-Technical\'\Memoral of the lost Generation\'\076 318 9700\'\calebserafin@outlook.com\'\v2_" & StringMult("1", Int(Worksheets("Attendance").Range("B1").Value)), 1, True) '<---This is what requires v2 Load and Save functions (the v2_ part)
                Call AttendanceData_load
            End If
        End If
        
        Application.ScreenUpdating = True
        Application.StatusBar = False
        Application.EnableEvents = wasEnabled
    End If
End Sub
