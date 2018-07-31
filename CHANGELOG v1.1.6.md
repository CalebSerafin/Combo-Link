### VBA Project: Combo-Link ###
### Makes, links and syncs different data sheets into Summery sheets. ###

# Recent Change Log: v1.1.6 #
## Attributes Class System ##
 Completing Attributes Class System to improve combustibility and generic handling of different types of worksheets.

### Overall ###
```diff
+ When inserting a Day the future date will no longer be auto formated(13/07/2018).
+ Fixed Array bug when inserting dates.
+ Percents are now truncated and will npt produce random numbers anymore.
- Removed the horrid getMembers sub found in the userForms.
- The word Integer has been removed from all source files.
+ All Integers are now Longs (Includes methods and CInt to CLng).
+ CountMembers_v2! runs faster and can handle any size change.
+ Calculations_On/_Off now fully disable all Excel UI features deep down the stack.
- Removed hard limit of 64 members from AddMember calls (Not the actual sub).
+ MaxMembers is now mostly outsourced to CountMembers Function(Good thing).
+ UpdateAttendanceList now focusing on Varients rather than a direct range.
```

## Nearby Goals ##
- [x] Fixing worksheet not updating.
- [x] Fixing References_RemoveMissing crashing.
- [x] Optimising all large range and 2 dimensional cell interactions to improve performance.
- [ ] Merge UpdateAttendanceList_v2 with AttendanceData_save_v2
- [x] Fix Load and save scripts going up to incorrect maxMembers instead of CountMembers
- [x] Fix Percent displays not always been a noramal sized number
- [ ] Completing Attributes Class System to improve combustibility and generic handling of different types of worksheets.
- [ ] Creating the first worksheet to utilise the Attributes Class System, the Score worksheet.
- [ ] Create Own GitSync Macro.
- [ ] Converting the Attendance worksheet and all its functions to use Attributes Class System.
- [ ] More to come...
