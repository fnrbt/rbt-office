' cscript //nologo open-all.vbs DIR
' Open DIR\docs\*.docx in Word (desktop session only). Preserve the append-only
' opened.tsv rows: file, ok, pages, characters; or file, ERROR, number, message;
' followed by DONE. opened-detail.tsv is also append-only, with five columns:
' escaped file, kind, 1-based index (0 for document), field, escaped value.
' Escaping is ASCII: \\, \t, \r, \n and \uXXXX for every other non-ASCII
' UTF-16 code unit. Thus story/control text remains complete, including tabs,
' paragraph marks and surrogate pairs, without making extra TSV records.
' A file's final "status" row is ok or ERROR; prior rows may be partial on error.
' OpenAndRepair is explicitly False. Word's Document COM API does not expose
' an authoritative "this open was repaired" result: repair_status is unknown.
Option Explicit
Dim fso, base, w, sink, names(), count, f, i, j, tmp, failure, quitError, oldUpdateLinks, linksCaptured, wordVersion, wordBuild
Set fso = CreateObject("Scripting.FileSystemObject")
base = WScript.Arguments(0)
failure = False
count = 0
ReDim names(0)
For Each f In fso.GetFolder(base & "\docs").Files
  If LCase(fso.GetExtensionName(f.Name)) = "docx" Then
    ReDim Preserve names(count)
    names(count) = f.Name
    count = count + 1
  End If
Next
For i = 1 To count - 1
  tmp = names(i)
  j = i - 1
  Do While j >= 0
    If StrComp(names(j), tmp, vbBinaryCompare) <= 0 Then Exit Do
    names(j + 1) = names(j)
    j = j - 1
  Loop
  names(j + 1) = tmp
Next

On Error Resume Next
linksCaptured = False
Set w = CreateObject("Word.Application")
If Err.Number = 0 Then w.Visible = False
If Err.Number = 0 Then w.DisplayAlerts = 0
If Err.Number = 0 Then w.AutomationSecurity = 3 ' msoAutomationSecurityForceDisable
If Err.Number = 0 Then
  oldUpdateLinks = w.Options.UpdateLinksAtOpen
  If Err.Number = 0 Then linksCaptured = True
End If
If Err.Number = 0 Then w.Options.UpdateLinksAtOpen = False
If Err.Number = 0 Then wordVersion = w.Version
If Err.Number = 0 Then wordBuild = w.Build
If Err.Number <> 0 Then
  WScript.Echo "Word startup ERROR " & Err.Number & ": " & Err.Description
  Err.Clear
  If linksCaptured Then w.Options.UpdateLinksAtOpen = oldUpdateLinks
  If Not w Is Nothing Then w.Quit
  WScript.Quit 1
End If
WScript.Echo "Word Version=" & wordVersion & " Build=" & wordBuild
On Error GoTo 0

On Error Resume Next
Set sink = fso.OpenTextFile(base & "\opened-detail.tsv", 8, True, 0)
If Err.Number <> 0 Then
  WScript.Echo "Detail output ERROR " & Err.Number & ": " & Err.Description
  Err.Clear
  w.Options.UpdateLinksAtOpen = oldUpdateLinks
  w.Quit
  WScript.Quit 1
End If
On Error GoTo 0
For i = 0 To count - 1
  ProcessFile names(i)
Next
On Error Resume Next
Say "DONE"
If Err.Number <> 0 Then
  WScript.Echo "Aggregate output ERROR " & Err.Number & ": " & Err.Description
  failure = True
  Err.Clear
End If
sink.Close
If Err.Number <> 0 Then
  WScript.Echo "Detail close ERROR " & Err.Number & ": " & Err.Description
  failure = True
  Err.Clear
End If
w.Options.UpdateLinksAtOpen = oldUpdateLinks
If Err.Number <> 0 Then
  WScript.Echo "Restore link setting ERROR " & Err.Number & ": " & Err.Description
  failure = True
  Err.Clear
End If
w.Quit
quitError = Err.Number
If quitError <> 0 Then WScript.Echo "Word.Quit ERROR " & quitError & ": " & Err.Description
On Error GoTo 0
If failure Or quitError <> 0 Then WScript.Quit 1

Sub ProcessFile(name)
  Dim d, stats, code, message, closeCode, closeMessage
  Set d = Nothing
  code = 0
  message = ""
  On Error Resume Next
  ' Documents.Open: ConfirmConversions=False, ReadOnly=True,
  ' AddToRecentFiles=False, Visible=True (12th), OpenAndRepair=False (14th).
  ' No retry or repair fallback; default passwords are never supplied.
  ' Word returns E_FAIL for Shape.Type when the document has no visible window.
  Set d = w.Documents.Open(fso.BuildPath(base & "\docs", name), False, True, False, , , , , , , , True, , False)
  If Err.Number <> 0 Then
    code = Err.Number
    message = "open: " & Err.Description
    Err.Clear
  ElseIf d Is Nothing Then
    code = -1
    message = "open returned no document"
  Else
    stats = Evidence(name, d)
    If Err.Number <> 0 Then
      code = Err.Number
      message = "evidence: " & Err.Description
      Err.Clear
    End If
  End If
  If Not d Is Nothing Then
    d.Close 0 ' wdDoNotSaveChanges
    If Err.Number <> 0 Then
      closeCode = Err.Number
      closeMessage = "close: " & Err.Description
      Err.Clear
      If code = 0 Then
        code = closeCode
        message = closeMessage
      Else
        message = message & "; " & closeMessage & " (" & closeCode & ")"
      End If
    End If
  End If
  If code = 0 Then
    Say name & vbTab & "ok" & vbTab & stats(1) & vbTab & stats(2)
    WriteField name, "status", 0, "value", "ok"
  Else
    failure = True
    message = Replace(Replace(Replace(message, vbCr, " "), vbLf, " "), vbTab, " ")
    Say name & vbTab & "ERROR" & vbTab & code & vbTab & message
    WriteField name, "status", 0, "value", "ERROR " & code & ": " & message
  End If
  If Err.Number <> 0 Then
    failure = True
    WScript.Echo "Output ERROR for " & name & ": " & Err.Number & ": " & Err.Description
    Err.Clear
  End If
  On Error GoTo 0
End Sub

Function Evidence(name, d)
  Dim body, sr, current, nextStory, tbl, cell, controls, cc, mapping
  Dim shape, inlineShape, hyperlink, field, para, sec, hf, slot, sectionIndex, hfIndex, hfShapeIndex
  Dim n, pages, chars, ordinal, cellIndex
  Set body = d.Content
  n = body.End
  chars = Len(body.Text)
  d.Repaginate
  pages = d.ComputeStatistics(2) ' wdStatisticPages
  WriteField name, "document", 0, "repair_status", "unknown (no authoritative Word Document COM property; OpenAndRepair=False)"
  WriteField name, "document", 0, "body_end", n
  WriteField name, "document", 0, "pages", pages
  WriteField name, "document", 0, "characters", chars
  WriteField name, "document", 0, "body_text", body.Text
  WriteField name, "document", 0, "fields", d.Fields.Count
  WriteField name, "document", 0, "tables_of_contents", d.TablesOfContents.Count
  WriteField name, "document", 0, "hyperlinks", d.Hyperlinks.Count
  For ordinal = 1 To d.Fields.Count
    Set field = d.Fields.Item(ordinal)
    WriteField name, "field", ordinal, "type", field.Type
    WriteField name, "field", ordinal, "code", field.Code.Text
    WriteField name, "field", ordinal, "result", field.Result.Text
  Next
  For ordinal = 1 To d.Hyperlinks.Count
    Set hyperlink = d.Hyperlinks.Item(ordinal)
    WriteField name, "hyperlink", ordinal, "address", hyperlink.Address
    WriteField name, "hyperlink", ordinal, "sub_address", hyperlink.SubAddress
  Next

  ordinal = 0
  For Each sr In d.StoryRanges
    Set current = sr
    Do While Not current Is Nothing
      ordinal = ordinal + 1
      WriteField name, "story", ordinal, "type", current.StoryType
      WriteField name, "story", ordinal, "start", current.Start
      WriteField name, "story", ordinal, "end", current.End
      WriteField name, "story", ordinal, "text", current.Text
      WriteField name, "story", ordinal, "inline_shapes", current.InlineShapes.Count
      Set nextStory = current.NextStoryRange
      Set current = nextStory
    Loop
  Next
  WriteField name, "document", 0, "stories", ordinal

  ordinal = 0
  For Each para In d.Paragraphs
    ordinal = ordinal + 1
    WriteField name, "paragraph", ordinal, "start", para.Range.Start
    WriteField name, "paragraph", ordinal, "end", para.Range.End
    WriteField name, "paragraph", ordinal, "page", para.Range.Information(3) ' wdActiveEndPageNumber
  Next
  WriteField name, "document", 0, "body_paragraphs", ordinal

  WriteField name, "document", 0, "tables", d.Tables.Count
  ordinal = 0
  For Each tbl In d.Tables
    ordinal = ordinal + 1
    WriteField name, "table", ordinal, "start", tbl.Range.Start
    WriteField name, "table", ordinal, "end", tbl.Range.End
    WriteField name, "table", ordinal, "text", tbl.Range.Text
    WriteField name, "table", ordinal, "cells", tbl.Range.Cells.Count
    cellIndex = 0
    For Each cell In tbl.Range.Cells
      cellIndex = cellIndex + 1
      WriteField name, "cell", ordinal & "." & cellIndex, "start", cell.Range.Start
      WriteField name, "cell", ordinal & "." & cellIndex, "end", cell.Range.End
      WriteField name, "cell", ordinal & "." & cellIndex, "text", cell.Range.Text
    Next
  Next

  Set controls = d.ContentControls
  WriteField name, "document", 0, "content_controls", controls.Count
  ordinal = 0
  For Each cc In controls
    ordinal = ordinal + 1
    WriteField name, "control", ordinal, "type", cc.Type
    WriteField name, "control", ordinal, "id", cc.ID
    WriteField name, "control", ordinal, "title", cc.Title
    WriteField name, "control", ordinal, "tag", cc.Tag
    WriteField name, "control", ordinal, "start", cc.Range.Start
    WriteField name, "control", ordinal, "end", cc.Range.End
    WriteField name, "control", ordinal, "text", cc.Range.Text
    Set mapping = cc.XMLMapping
    WriteField name, "control", ordinal, "is_mapped", mapping.IsMapped
    If mapping.IsMapped Then
      WriteField name, "control", ordinal, "xpath", mapping.XPath
      WriteField name, "control", ordinal, "prefix_mappings", mapping.PrefixMappings
      WriteField name, "control", ordinal, "custom_xml_part_id", mapping.CustomXMLPart.Id
    End If
  Next

  WriteField name, "document", 0, "inline_shapes", d.InlineShapes.Count
  ordinal = 0
  For Each inlineShape In d.InlineShapes
    ordinal = ordinal + 1
    WriteField name, "inline_shape", ordinal, "type", inlineShape.Type
    WriteField name, "inline_shape", ordinal, "start", inlineShape.Range.Start
    WriteField name, "inline_shape", ordinal, "end", inlineShape.Range.End
  Next
  WriteField name, "document", 0, "floating_shapes", d.Shapes.Count
  For ordinal = 1 To d.Shapes.Count
    Set shape = IndexedShape(d.Shapes, ordinal, "document")
    WriteField name, "floating_shape", ordinal, "type", ShapeValue(shape, "type", "document", ordinal)
    WriteField name, "floating_shape", ordinal, "anchor_start", ShapeValue(shape, "anchor_start", "document", ordinal)
    WriteField name, "floating_shape", ordinal, "anchor_end", ShapeValue(shape, "anchor_end", "document", ordinal)
  Next
  ' Document.Shapes excludes header/footer stories; report their Shapes
  ' separately by section and primary/first/even slot (including linked
  ' sections as Word exposes them). Body floating_shapes remains unchanged.
  sectionIndex = 0
  For Each sec In d.Sections
    sectionIndex = sectionIndex + 1
    For slot = 1 To 3 ' wdHeaderFooterPrimary, FirstPage, EvenPages
      Set hf = sec.Headers(slot)
      hfIndex = sectionIndex & ".header." & slot
      WriteField name, "header_footer", hfIndex, "exists", hf.Exists
      If hf.Exists Then
        WriteField name, "header_footer", hfIndex, "floating_shapes", hf.Shapes.Count
        For hfShapeIndex = 1 To hf.Shapes.Count
          Set shape = IndexedShape(hf.Shapes, hfShapeIndex, "header_footer " & hfIndex)
          WriteField name, "header_footer_shape", hfIndex & "." & hfShapeIndex, "type", ShapeValue(shape, "type", "header_footer " & hfIndex, hfShapeIndex)
          WriteField name, "header_footer_shape", hfIndex & "." & hfShapeIndex, "anchor_start", ShapeValue(shape, "anchor_start", "header_footer " & hfIndex, hfShapeIndex)
        Next
      End If
      Set hf = sec.Footers(slot)
      hfIndex = sectionIndex & ".footer." & slot
      WriteField name, "header_footer", hfIndex, "exists", hf.Exists
      If hf.Exists Then
        WriteField name, "header_footer", hfIndex, "floating_shapes", hf.Shapes.Count
        For hfShapeIndex = 1 To hf.Shapes.Count
          Set shape = IndexedShape(hf.Shapes, hfShapeIndex, "header_footer " & hfIndex)
          WriteField name, "header_footer_shape", hfIndex & "." & hfShapeIndex, "type", ShapeValue(shape, "type", "header_footer " & hfIndex, hfShapeIndex)
          WriteField name, "header_footer_shape", hfIndex & "." & hfShapeIndex, "anchor_start", ShapeValue(shape, "anchor_start", "header_footer " & hfIndex, hfShapeIndex)
        Next
      End If
    Next
  Next
  WriteField name, "document", 0, "sections", sectionIndex
  Evidence = Array(n, pages, chars)
End Function

' Preserve COM errors while identifying the shape collection, index and property.
Function IndexedShape(shapes, index, context)
  Dim result, number, description, source
  On Error Resume Next
  Set result = shapes.Item(index)
  number = Err.Number : description = Err.Description : source = Err.Source
  On Error GoTo 0
  If number <> 0 Then Err.Raise number, source, context & " Shapes.Item(" & index & "): " & description
  Set IndexedShape = result
End Function

Function ShapeValue(shape, field, context, index)
  Dim value, number, description, source
  On Error Resume Next
  Select Case field
    Case "type": value = shape.Type
    Case "anchor_start": value = shape.Anchor.Start
    Case "anchor_end": value = shape.Anchor.End
  End Select
  number = Err.Number : description = Err.Description : source = Err.Source
  On Error GoTo 0
  If number <> 0 Then Err.Raise number, source, context & " Shapes.Item(" & index & ")." & field & ": " & description
  ShapeValue = value
End Function

Sub WriteField(name, kind, ordinal, field, value)
  EmitEscaped name
  sink.Write vbTab & kind & vbTab & CStr(ordinal) & vbTab & field & vbTab
  EmitEscaped CStr(value)
  sink.WriteLine
End Sub

Sub EmitEscaped(value)
  Dim k, u, fragment, chunk
  chunk = ""
  For k = 1 To Len(value)
    u = AscW(Mid(value, k, 1)) And &HFFFF&
    Select Case u
      Case 9: fragment = "\t"
      Case 10: fragment = "\n"
      Case 13: fragment = "\r"
      Case 92: fragment = "\\"
      Case Else
        If u < 32 Or u > 126 Then
          fragment = "\u" & Right("0000" & Hex(u), 4)
        Else
          fragment = ChrW(u)
        End If
    End Select
    chunk = chunk & fragment
    If Len(chunk) >= 4096 Then
      sink.Write chunk
      chunk = ""
    End If
  Next
  If Len(chunk) > 0 Then sink.Write chunk
End Sub

Sub Say(text)
  Dim opened
  Set opened = fso.OpenTextFile(base & "\opened.tsv", 8, True)
  opened.WriteLine text
  opened.Close
End Sub
