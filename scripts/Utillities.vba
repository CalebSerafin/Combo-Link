Option Explicit
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
'#############################################################
'Dim lastCalcValue As Long:lastCalcValue = Calculations_Off '#
'Call Calculations_On(lastCalcValue):lastCalcValue = 0 ''''''#
'#############################################################
Function Calculations_Off_v1() As Long   'Save return for Calculations_On function
    Dim lastCalcValue As Long
    lastCalcValue = 1
    With Application
        If (.EnableEvents = False Or .Calculation = xlCalculationManual) Then lastCalcValue = 0: 'Butchered to favour Enable Events as well as than calculate
        
        'lastCalcValue = .Calculation
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    Calculations_Off_v1 = lastCalcValue
End Function
Sub Calculations_On_v1(ByVal lastCalcValue As Long)    'Take value from Calculations_Off function
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic 'lastCalcValue
        .EnableEvents = True
        
        If lastCalcValue = 0 Then .EnableEvents = False: .Calculation = xlCalculationManual 'Butchered to favour Enable Events as well as than calculate
    End With
End Sub

Function IsInArray_v1(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray_v1 = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function
Function GetMonth_v1(ByVal number As Long, Optional ByVal longName As Boolean = False) As String
    If IsNumeric(number) Then
        number = CLng(number)
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
Function StringMult_v1(ByVal Word As String, ByVal Multiply As Long) As String
    Dim repeat As Long
    StringMult_v1 = ""
    Multiply = CLng(Multiply)
    
    If Multiply >= 1 Then
        For repeat = 1 To Multiply
            StringMult_v1 = StringMult_v1 + Word
        Next repeat
    End If
End Function
Function addCellData_v1(ByVal mode As String, ByVal sheet As String, ByVal min As Long, ByVal max As Long, ByVal rawData As String, ByVal topLeft As Long, ByVal forceLast As Boolean)
'#############################################################
Dim lastCalcValue As Long: lastCalcValue = Calculations_Off '#
'Call Calculations_On(lastCalcValue):lastCalcValue = 0 ''''''#
'#############################################################
    addCellData_v1 = False
    
    If mode = "row" Or mode = "column" Then
        Dim data1D() As String
        Dim data1DUBound As Long
        
        data1D = Split(rawData, "\'\")
        data1DUBound = UBound(data1D, 1)
    End If
    
    If mode = "row" Then
        Dim index1D As Long
        Dim atColumn As Long
        Dim isFree As Boolean
        Dim howLong As Long
        If forceLast = True Then
            howLong = data1DUBound - 1
        Else
            howLong = data1DUBound
        End If
        For index1D = min To max
            isFree = True
            For atColumn = topLeft To howLong + topLeft
                If Not IsEmpty(Worksheets(sheet).Cells(index1D, atColumn).Value) Then
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
'#############################################################
'Dim lastCalcValue As Long:lastCalcValue = Calculations_Off '#
Call Calculations_On(lastCalcValue): lastCalcValue = 0 ''''''#
'#############################################################
End Function
Sub References_RemoveMissing_v1() 'Removes missing References from VBE
    If Not VBAIsTrusted() Then
        Call CheckTrustAccess
    End If
    
    Dim theRef As Variant, i As Long
    
    For i = ThisWorkbook.VBProject.References.Count To 1 Step -1
        Set theRef = ThisWorkbook.VBProject.References.Item(i)
        If theRef.isBroken = True Then
            ThisWorkbook.VBProject.References.Remove (theRef)
        End If
    Next i
    
    If Err <> 0 Then
        MsgBox "A missing reference has been encountered!" & "You will need to remove the reference manually.", vbCritical, "Unable To Remove Missing Reference"
    End If
End Sub
Sub ScanCommonError_v1() ' Requires v2 Load and Save functions
'#############################################################
Dim lastCalcValue As Long: lastCalcValue = Calculations_Off '#
'Call Calculations_On(lastCalcValue):lastCalcValue = 0 ''''''#
'#############################################################
    Application.StatusBar = "Refreshing...."
    
    Dim Row As Long
    
    'row = 2
    'For row = 2 To 65 Step 1
    '    If ((Worksheets("Details").Cells(row, 1) = "Brink") And (Worksheets("Details").Cells(row, 2) = "Nelson")) Or (Worksheets("Details").Cells(row, 6) = "072 223 2173") Then
    '        Worksheets("Details").Cells(row, 5) = "Boss"
    '    End If
    'Next row
    If Worksheets("COMPUTING DON'T TOUCH").Cells(5, 12).Value = "Drama Club" Then
        For Row = 2 To maxMembers + 1 Step 1
            If ((Worksheets("Details").Cells(Row, 1) = "Caleb") And (Worksheets("Details").Cells(Row, 2) = "Serafin")) Or (Worksheets("Details").Cells(Row, 6) = "076 318 9700") Then
                Worksheets("Details").Cells(Row, 5) = "Memoral of the lost Generation"
            End If
        Next Row
    End If
    '///GapRemover
    Dim index1D As Long
    Dim atColumn As Long
    Dim isFree As Boolean
    For index1D = 2 To maxMembers + 1
        isFree = True
        For atColumn = 1 To 7
            If Not IsEmpty(Worksheets("Details").Cells(index1D, atColumn).Value) Then
                isFree = False
                Exit For
            End If
        Next atColumn
        If isFree = True And IsEmpty(Worksheets("Details").Cells(index1D, 8).Value) Then    '("v2_" & StringMult("0", CLng(Worksheets("Attendance").Range("B1").Value))) Then
            Debug_msg ("Module 1: ScanCommonError_v1: Terminating GapRemover at row: " & index1D)
            Exit For
        End If
        If isFree = True Then
            Debug_msg ("Module 1: ScanCommonError_v1: Found data without details at row: " & index1D)
            Worksheets("Details").Cells(index1D, 8).Value = ""  '"v2_" & StringMult("0", CLng(Worksheets("Attendance").Range("B1").Value)) '<---This is what requires v2 Load and Save functions (the v2_ part)
            Call AttendanceData_load
        End If
    Next index1D
    '///End GapRemover
    
    If Worksheets("COMPUTING DON'T TOUCH").Cells(5, 12).Value = "LostMemory" Then
        Dim found As Boolean
        found = False
        For Row = 2 To maxMembers + 1 + 1 And found = False Step 1
            If ((Worksheets("Details").Cells(Row, 1) = "Caleb") And (Worksheets("Details").Cells(Row, 2) = "Serafin")) Or (Worksheets("Details").Cells(Row, 6) = "076 318 9700") Then
                found = True
            End If
        Next Row
        If found <> True Then
            Call addCellData("row", "Details", 2, maxMembers + 1 + 1, "Caleb\'\Serafin\'\10.5\'\Ex-Technical\'\Memoral of the lost Generation\'\076 318 9700\'\calebserafin@outlook.com\'\v2_" & StringMult("1", CLng(Worksheets("Attendance").Range("B1").Value)), 1, True) '<---This is what requires v2 Load and Save functions (the v2_ part)
            Call AttendanceData_load
        End If
    End If
    
    Application.StatusBar = False
'#############################################################
'Dim lastCalcValue As Long:lastCalcValue = Calculations_Off '#
Call Calculations_On(lastCalcValue): lastCalcValue = 0 ''''''#
'#############################################################
End Sub
Function JoinDetailNames_v1() As String()
    Dim AmountMembers As Long
    Dim NamesBoth As Variant
    Dim NamesJoint() As String
    Dim CurrentMember As Long
    
    AmountMembers = CountMembers
    
    ReDim NamesBoth(1 To AmountMembers, 1 To 2) As Variant
    ReDim NamesJoint(1 To AmountMembers) As String
    ReDim JoinDetailNames_v1(1 To AmountMembers) As String
    
    With Worksheets("Details")
        NamesBoth = .Range(.Cells(2, 1), .Cells(AmountMembers + 1, 2))
    End With
    
    For CurrentMember = 1 To AmountMembers Step 1
        NamesJoint(CurrentMember) = NamesBoth(CurrentMember, 1) & " " & NamesBoth(CurrentMember, 2)
    Next CurrentMember
    
    JoinDetailNames_v1 = NamesJoint()
End Function
