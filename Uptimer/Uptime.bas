Attribute VB_Name = "Uptime"
Public Enum TimeFormatType
DaysHoursMinutesSecondsMilliseconds = 0
DaysHoursMinutesSeconds = 1
DHMSMColonSeparated = 2
DaysHoursMinutes = 3
End Enum

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Function FormatCount(Count As Long, Optional FormatType As TimeFormatType = 0) As String
Dim Days As Long, Hours As Long, Minutes As Long, Seconds As Long, Miliseconds As Long
    
Miliseconds = Count Mod 1000
Count = Count \ 1000
Days = Count \ (24& * 3600&)
If Days > 0 Then Count = Count - (24& * 3600& * Days)
Hours = Count \ 3600&
If Hours > 0 Then Count = Count - (3600& * Hours)
Minutes = Count \ 60
Seconds = Count Mod 60

Select Case FormatType
Case 0

FormatCount = Days & " days, " & Hours & " hours, " & _
Minutes & " minutes, " & Seconds & " seconds, " & Miliseconds & _
" Miliseconds"
Case 1

FormatCount = Days & " days, " & Hours & " hours, " & _
Minutes & " minutes, " & Seconds & " seconds"
Case 2

FormatCount = Days & ":" & Hours & ":" & _
Minutes & ":" & Seconds & ":" & Miliseconds
Case 3
            
FormatCount = Days & " days, " & Hours & " hours, " & _
Minutes & " minutes"
End Select
End Function
