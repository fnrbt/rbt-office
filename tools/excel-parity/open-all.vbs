' cscript open-all.vbs DIR: opens DIR\docs\*.xlsx read-only in Excel.
' Appends the existing aggregate records to DIR\opened.tsv and writes one UTF-16
' tab-separated file per workbook in DIR\evidence\<workbook name>.tsv.
' Fields in evidence are escaped: backslash, tab, CR and LF become \\, \t, \r, \n.
' Records: WORKBOOK(name); SHEET(index,name,type,visible); CELL(sheet,address,
' value type,value,formula); TABLE(sheet,name,range); SHAPE(sheet,index,name,
' type,top-left,bottom-right); CHART(sheet,shape,chart type); SERIES(sheet,
' shape,index,formula); NOTE(sheet,address,author,text); THREAD(sheet,address,
' author,text); REPLY(sheet,address,index,author,text). Chart sheets use an
' empty shape name. Unsupported thread or chart-series APIs emit CAPABILITY_ERROR
' and a nonzero exit, but other workbook evidence is still captured. Values are Excel
' Value2; numbers use Excel's 15-significant-digit scientific formatting, and
' errors use ERROR.TYPE. Formula is Excel's invariant Formula property.
' For each *.original.xlsx, evidence also includes an *.expected-insert.xlsx
' snapshot after Excel inserts row 2 and column B in its first worksheet.
' This sidecar has the same records but no corresponding opened.tsv row.
' Run in the interactive desktop session; compare each save/edit/insert to original.
Option Explicit
Dim fso, base, x, wb, detail, exitCode
Set fso = CreateObject("Scripting.FileSystemObject")
Set x = Nothing : Set wb = Nothing : Set detail = Nothing
exitCode = 0
If WScript.Arguments.Count <> 1 Then
  WScript.StdErr.WriteLine "Usage: cscript open-all.vbs DIR"
  WScript.Quit 2
End If
base = WScript.Arguments(0)
Main
WScript.Quit exitCode

Sub Main()
  Dim f, n, description, source, closeNumber, closeDescription
  On Error Resume Next
  If Not fso.FolderExists(base & "\evidence") Then fso.CreateFolder base & "\evidence"
  If Err.Number = 0 Then Set x = CreateObject("Excel.Application")
  If Err.Number = 0 Then WScript.Echo "Excel " & x.Version & " build " & x.Build
  If Err.Number = 0 Then x.Visible = False
  If Err.Number = 0 Then x.DisplayAlerts = False
  If Err.Number = 0 Then x.AutomationSecurity = 3 ' msoAutomationSecurityForceDisable
  If Err.Number = 0 Then x.AskToUpdateLinks = False
  If Err.Number = 0 Then
    For Each f In fso.GetFolder(base & "\docs").Files
      If LCase(fso.GetExtensionName(f.Name)) = "xlsx" Then
        Set wb = Nothing
        Set wb = x.Workbooks.Open(f.Path, 0, True) ' UpdateLinks=0, ReadOnly=True; no repair mode
        If Err.Number <> 0 Then
          n = Err.Number : description = Err.Description : source = Err.Source
          Err.Clear
          Say f.Name & vbTab & "ERROR" & vbTab & n & vbTab & CleanMessage(description)
          If Err.Number = 0 Then Err.Raise n, source, description
        Else
          CaptureWorkbook f.Name, True
          If Err.Number = 0 And LCase(Right(f.Name, 14)) = ".original.xlsx" Then CaptureExpectedInsertion f.Name
        End If
        n = Err.Number : description = Err.Description : source = Err.Source
        Err.Clear
        If Not detail Is Nothing Then
          detail.Close
          If Err.Number <> 0 And n = 0 Then n = Err.Number : description = Err.Description : source = Err.Source
          Err.Clear
          Set detail = Nothing
        End If
        If Not wb Is Nothing Then
          wb.Close False
          If Err.Number <> 0 And n = 0 Then n = Err.Number : description = Err.Description : source = Err.Source
          Err.Clear
          Set wb = Nothing
        End If
        If n <> 0 Then
          Err.Raise n, source, description
          Exit For
        End If
      End If
    Next
  End If
  n = Err.Number : description = Err.Description : source = Err.Source
  Err.Clear
  If Not x Is Nothing Then
    x.Quit
    closeNumber = Err.Number : closeDescription = Err.Description
    Err.Clear
    If n = 0 And closeNumber <> 0 Then n = closeNumber : description = closeDescription
    Set x = Nothing
  End If
  If n <> 0 Then
    WScript.StdErr.WriteLine "Excel oracle ERROR " & n & ": " & CleanMessage(description)
    exitCode = 1
  Else
    Say "DONE"
    If Err.Number <> 0 Then
      WScript.StdErr.WriteLine "Excel oracle ERROR " & Err.Number & ": " & CleanMessage(Err.Description)
      exitCode = 1
    End If
  End If
  On Error GoTo 0
End Sub

Sub CaptureExpectedInsertion(name)
  Dim expectedName, operation, number, description
  expectedName = Left(name, Len(name) - 14) & ".expected-insert.xlsx"
  detail.Close
  Set detail = Nothing
  On Error Resume Next
  operation = "insert row 2"
  wb.Worksheets(1).Rows(2).Insert
  If Err.Number = 0 Then
    operation = "insert column B"
    wb.Worksheets(1).Columns(2).Insert
  End If
  number = Err.Number : description = Err.Description
  Err.Clear
  On Error GoTo 0
  If number <> 0 Then
    Set detail = fso.CreateTextFile(base & "\evidence\" & expectedName & ".tsv", True, True)
    Row Array("WORKBOOK", expectedName)
    Row Array("OPERATION_ERROR", operation, number, description)
    WScript.StdErr.WriteLine name & ": Excel rejected " & operation & ": " & number & " " & CleanMessage(description)
    exitCode = 1
  Else
    CaptureWorkbook expectedName, False
  End If
End Sub

Sub CaptureWorkbook(name, aggregate)
  Dim s, r, c, v, typ, value, formula, used, formulas, chars, total, state
  Dim sheetIndex, t, sh, i, chart, comment
  used = 0 : formulas = 0 : chars = 0 : total = 0
  state = "ok"
  If InStr(1, x.Caption & wb.Name, "Repaired", 1) > 0 Then state = "repaired"
  Set detail = fso.CreateTextFile(base & "\evidence\" & name & ".tsv", True, True)
  Row Array("WORKBOOK", name)
  sheetIndex = 0
  For Each s In wb.Sheets
    sheetIndex = sheetIndex + 1
    Row Array("SHEET", sheetIndex, s.Name, s.Type, s.Visible)
    If s.Type = -4167 Then ' xlWorksheet
      Set r = s.UsedRange
      used = used + r.Cells.Count
      For Each c In r.Cells
        v = c.Value2
        formula = ""
        If c.HasFormula Then
          formula = c.Formula
          formulas = formulas + 1
        End If
        If Not IsEmpty(v) Or c.HasFormula Then
          If VarType(v) = vbError Then
            typ = "error"
            value = CStr(x.Evaluate("ERROR.TYPE('" & Replace(s.Name, "'", "''") & "'!" & c.Address & ")"))
          ElseIf IsEmpty(v) Then
            typ = "empty" : value = ""
          ElseIf VarType(v) = vbString Then
            typ = "string" : value = v : chars = chars + Len(v)
          ElseIf VarType(v) = vbBoolean Then
            typ = "boolean" : value = CStr(v)
          Else
            typ = "number"
            value = x.WorksheetFunction.Text(v, "0.00000000000000E+00")
            total = total + v
          End If
          Row Array("CELL", s.Name, c.Address(False, False), typ, value, formula)
        End If
      Next
      For Each t In s.ListObjects
        Row Array("TABLE", s.Name, t.Name, t.Range.Address(False, False))
      Next
      For Each comment In s.Comments
        Row Array("NOTE", s.Name, comment.Parent.Address(False, False), comment.Author, comment.Text)
      Next
      CaptureThreads s
      i = 0
      For Each sh In s.Shapes
        i = i + 1
        Row Array("SHAPE", s.Name, i, sh.Name, sh.Type, sh.TopLeftCell.Address(False, False), sh.BottomRightCell.Address(False, False))
        If sh.Type = 3 Then ' msoChart
          Set chart = sh.Chart
          CaptureChart s.Name, sh.Name, chart
        End If
      Next
    ElseIf s.Type = -4109 Then ' xlChart
      CaptureChart s.Name, "", s
    End If
  Next
  If aggregate Then Say name & vbTab & state & vbTab & wb.Worksheets.Count & vbTab & used & vbTab & formulas & vbTab & chars & vbTab & Round(total, 2)
End Sub

Sub CaptureThreads(sheet)
  Dim comments, comment, reply, replyIndex, number, description, source
  On Error Resume Next
  Set comments = sheet.CommentsThreaded
  number = Err.Number : description = Err.Description : source = Err.Source
  On Error GoTo 0
  If number <> 0 Then
    If number = 1004 And InStr(1, description, "Object Model for Threaded Comments is not supported", vbTextCompare) > 0 Then
      Row Array("CAPABILITY_ERROR", sheet.Name, "CommentsThreaded", number, description)
      WScript.StdErr.WriteLine sheet.Name & ": " & CleanMessage(description)
      exitCode = 1
      Exit Sub
    End If
    Err.Raise number, source, description
  End If
  For Each comment In comments
    Row Array("THREAD", sheet.Name, comment.Parent.Address(False, False), comment.Author.Name, comment.Text)
    replyIndex = 0
    For Each reply In comment.Replies
      replyIndex = replyIndex + 1
      Row Array("REPLY", sheet.Name, comment.Parent.Address(False, False), replyIndex, reply.Author.Name, reply.Text)
    Next
  Next
End Sub

Sub CaptureChart(sheetName, shapeName, chart)
  Dim series, i, formula, number, description, source
  Row Array("CHART", sheetName, shapeName, chart.ChartType)
  i = 0
  For Each series In chart.SeriesCollection
    i = i + 1
    On Error Resume Next
    formula = series.Formula
    number = Err.Number : description = Err.Description : source = Err.Source
    Err.Clear
    On Error GoTo 0
    If number = 445 Then
      Row Array("CAPABILITY_ERROR", sheetName, "Series.Formula", number, description, shapeName, i)
      WScript.StdErr.WriteLine sheetName & "/" & shapeName & ": Series.Formula: " & CleanMessage(description)
      exitCode = 1
    ElseIf number <> 0 Then
      Err.Raise number, source, description
    Else
      Row Array("SERIES", sheetName, shapeName, i, formula)
    End If
  Next
End Sub

Sub Row(fields)
  Dim field, line
  line = ""
  For Each field In fields
    If Len(line) > 0 Then line = line & vbTab
    line = line & EscapeField(CStr(field))
  Next
  detail.WriteLine line
End Sub

Function EscapeField(value)
  value = Replace(value, "\", "\\")
  value = Replace(value, vbTab, "\t")
  value = Replace(value, vbCr, "\r")
  EscapeField = Replace(value, vbLf, "\n")
End Function

Function CleanMessage(value)
  CleanMessage = Replace(Replace(value, vbCr, " "), vbLf, " ")
End Function

Sub Say(text)
  Dim sink
  Set sink = fso.OpenTextFile(base & "\opened.tsv", 8, True)
  sink.WriteLine text
  sink.Close
End Sub
