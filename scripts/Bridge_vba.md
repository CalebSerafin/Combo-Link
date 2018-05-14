# VBA Project: **Combo-Link**
## VBA Module: **[Bridge](/scripts/Bridge.vba "source is here")**
### Type: StdModule  

This procedure list for repo (Combo-Link) was automatically created on 14/05/2018 16:03:29 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in Bridge

---
VBA Procedure: **Debug_msg**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function Debug_msg(ByVal msg As String, Optional ByVal code As String = "null", Optional ByVal request As String = "null") As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||
ByVal|Variant|True||
ByVal|Variant|True||


---
VBA Procedure: **IREC**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub IREC()*  

**no arguments required for this procedure**


---
VBA Procedure: **StringMult**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function StringMult(ByVal Word As String, ByVal Multiply As Integer) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||
ByVal|Integer|False||


---
VBA Procedure: **addCellData**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function addCellData(ByVal mode As String, ByVal sheet As String, ByVal min As Integer, ByVal max As Integer, ByVal rawData As String, ByVal topLeft As Integer, ByVal forceLast As Boolean)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||
ByVal|String|False||
ByVal|Integer|False||
ByVal|Integer|False||
ByVal|String|False||
ByVal|Integer|False||
ByVal|Boolean|False||


---
VBA Procedure: **AttendanceData_save**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub AttendanceData_save()*  

**no arguments required for this procedure**


---
VBA Procedure: **UpdateAttendanceList**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub UpdateAttendanceList(Optional ByVal save As Boolean = True)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Variant|True||


---
VBA Procedure: **AttendanceData_load**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub AttendanceData_load()*  

**no arguments required for this procedure**


---
VBA Procedure: **ScanCommonError**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub ScanCommonError()*  

**no arguments required for this procedure**


---
VBA Procedure: **PositionAttendanceColomnButtons**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub PositionAttendanceColomnButtons(Optional ByVal colomn As Integer = 0)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Variant|True||


---
VBA Procedure: **GetMonth**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function GetMonth(ByVal number As Integer) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Integer|False||


---
VBA Procedure: **FindMember**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function FindMember(ByVal firstName As String, ByVal lastName As String, Optional ByVal matchCase As Boolean = True)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||
ByVal|String|False||
ByVal|Variant|True||


---
VBA Procedure: **References_RemoveMissing**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub References_RemoveMissing()*  

**no arguments required for this procedure**


---
VBA Procedure: **IsInArray**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
stringToBeFound|String|False||
arr|Variant|False||
