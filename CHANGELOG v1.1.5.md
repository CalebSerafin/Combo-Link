### VBA Project: Combo-Link ###
### Makes, links and syncs different data sheets into Summery sheets. ###

# Recent Change Log: v1.1.5 #
## Optimisation ##
Optimising all large range and 2 dimensional cell interactions to improve performance.

### Overall ###
```diff
+Added All Modules and Classes to VBAGit List.
+Added All Userforms to VBAGit List.
+Added All Worksheets and ThisWorkbook to VBAGit List.
-Removed Excution block for Git Read.
-Removed Worksheets and ThisWorkbook to VBAGit List.
-Removed Excution block for Git Read.
+Added Excution block for Git Read again.
-Removed Excution block for Git Read for the last time.
+Added Goal: Creat Own GitSync Macro.
+Fixed Spelling Error in Computing sheet.
+Added Calculation Toggle Function and Sub.
+Added CountMembers_v1 which will Greatly improve performance.
+Duplicated UpdateAttendanceList_v1 to allow working code.
-Removed use of Range object for optimisation since it's a pointer
+Added use of Variants instead of Range objects
+Tested out Optimisation on UpdateAttendanceList_v2, It worked.
+Changed Most Integers in range functions to longs
-Removed direct range calls
+Optimized all functions that work with ranges
```

## Nearby Goals ##
- [x] Fixing worksheet not updating.
- [x] Fixing References_RemoveMissing crashing.
- [x] Optimising all large range and 2 dimensional cell interactions to improve performance.
- [ ] Merge UpdateAttendanceList_v2 with AttendanceData_save_v2
- [ ] Completing Attributes Class System to improve combustibility and generic handling of different types of worksheets.
- [ ] Creating the first worksheet to utilise the Attributes Class System, the Score worksheet.
- [ ] Create Own GitSync Macro.
- [ ] Converting the Attendance worksheet and all its functions to use Attributes Class System.
- [ ] More to come...
