Sub Main()
    Dim d As Date
    d = DateSerial(2024, 1, 31)
    P "dateserial", d
    P "dateserial month 13", DateSerial(2023, 13, 1)
    P "dateserial day 0", DateSerial(2024, 3, 0)
    P "dateadd month end", DateAdd("m", 1, d)
    P "dateadd days", DateAdd("d", 30, d)
    P "dateadd year leap", DateAdd("yyyy", 1, DateSerial(2024, 2, 29))
    P "datediff days", DateDiff("d", DateSerial(2024, 1, 1), DateSerial(2024, 12, 31))
    P "datediff months", DateDiff("m", DateSerial(2024, 1, 31), DateSerial(2024, 2, 1))
    P "datediff weeks", DateDiff("ww", DateSerial(2024, 1, 1), DateSerial(2024, 3, 1))
    P "weekday", Weekday(d)
    P "weekday monday", Weekday(d, vbMonday)
    P "weekdayname", WeekdayName(Weekday(d))
    P "monthname", MonthName(2, True)
    P "datepart q", DatePart("q", d)
    P "datepart y", DatePart("y", DateSerial(2024, 3, 1))
    P "year month day", Year(d) & "/" & Month(d) & "/" & Day(d)
    P "timeserial", TimeSerial(13, 45, 30)
    P "hour", Hour(TimeSerial(13, 45, 30))
    P "cdate iso", CDate("2024-02-29")
    P "date literal", #3/15/2024#
    P "date plus", #3/15/2024# + 1.5
    P "date minus date", #3/15/2024# - #3/1/2024#
    P "format date", Format(d, "dddd, mmmm d, yyyy")
    P "format short", Format(d, "dd/mm/yy")
    P "format time", Format(TimeSerial(13, 5, 0), "h:mm AM/PM")
    P "isdate", IsDate("2024-13-01")
    P "isdate ok", IsDate("2024-12-01")
    P "datevalue", DateValue("February 3, 2024")
    P "cdbl date", CDbl(d)
    P "typename date", TypeName(d)
    P "date compare", d > #1/1/2024#
End Sub
