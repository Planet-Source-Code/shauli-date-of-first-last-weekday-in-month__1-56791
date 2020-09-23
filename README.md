<div align="center">

## Date of First/Last weekday in month


</div>

### Description

What date is Thanksgiving, or Labor Day? Get date of First or last weekday in month. i.e. Get first Monday of September, or last Thursday of November.
 
### More Info
 
Required day (Sunday-Saturday), month (1-12), year (any), first (0) or last (1).

Date


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Shauli](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/shauli.md)
**Level**          |Advanced
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/shauli-date-of-first-last-weekday-in-month__1-56791/archive/master.zip)





### Source Code

```
Public Function GetFirstLastDate(ByVal fnDay As String, fnMonth As Integer, fnYear As Integer, fnFirstLast As Byte) As Date
Dim tmpDate As Date, dLoop As Integer, addDate As Date, tmpLastDate As Date
addDate = DateSerial(fnYear, fnMonth, 1)
Select Case fnFirstLast
 Case 0
 If WeekdayName(Weekday(addDate)) = fnDay Then
  GetFirstLastDate = addDate
  Exit Function
 End If
 For dLoop = 1 To 7
  tmpDate = DateAdd("w", dLoop, addDate)
  If WeekdayName(Weekday(tmpDate)) = fnDay Then
   GetFirstLastDate = tmpDate
   Exit For
  End If
 Next dLoop
 Case 1
 tmpLastDate = DateAdd("d", -1, DateAdd("m", 1, addDate))
 If WeekdayName(Weekday(tmpLastDate)) = fnDay Then
  GetFirstLastDate = tmpLastDate
  Exit Function
 End If
 For dLoop = 7 To 1 Step -1
  tmpDate = DateAdd("w", -dLoop, tmpLastDate)
  If WeekdayName(Weekday(tmpDate)) = fnDay Then
   GetFirstLastDate = tmpDate
   Exit For
  End If
 Next dLoop
End Select
End Function
'Usage example:
Private Sub Command1_Click()
MsgBox GetFirstLastDate("Monday", 9, 2004, 0)
End Sub
```

