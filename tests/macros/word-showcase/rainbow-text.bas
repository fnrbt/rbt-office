Attribute VB_Name = "RainbowText"
' Name: Rainbow text
' Source: https://gist.github.com/KaitoHH/34840c30fe1b17752068366f5c9a237e ("change into rainbow color for
'         every single characters (tested on Microsoft Word 2016)")
' Author: KaitoHH
' Licence: not stated (a public gist); quoted with attribution, one local variable renamed (see Adaptations)
' What it does: colours every character of the document in turn violet, red, orange, yellow, green,
'               turquoise, blue.
' Adaptations: in `color` the local variable `color` is `colors` (Rbt.Vb 0.4.0 takes `color(i Mod 7)` inside
'              Sub color for a call of the Sub - error 450; Word takes the variable); Setup (added) types a
'              few paragraphs to colour, and Main runs both.
' Needs: some text in the document (Setup types it)
' Entry point: Main
Sub color()
  Application.ScreenUpdating = False
  Dim i As Long
  Dim colors
  'rainbow color
  colors = Array(8388736, 255, 33023, 65535, 65280, 16776960, 16711680)
  With ActiveDocument
    For i = 1 To .Characters.Count
      .Characters(i).Font.color = colors(i Mod 7)
    Next
  End With
  Application.ScreenUpdating = True
End Sub

Sub Setup()
    With ActiveDocument.Content
        .Text = "Somewhere over the rainbow"
        .InsertParagraphAfter
        .InsertAfter "Every character of this document gets a colour of its own, in the order of the rainbow: " & _
            "a macro walks the Characters collection and sets Font.Color on each one."
        .InsertParagraphAfter
        .InsertAfter "VBA in the browser, without Word."
    End With
    With ActiveDocument.Paragraphs(1)
        .Range.Font.Size = 36
        .Range.Font.Bold = True
        .Alignment = wdAlignParagraphCenter
    End With
    ActiveDocument.Paragraphs(2).Range.Font.Size = 16
    ActiveDocument.Paragraphs(3).Range.Font.Size = 24
    ActiveDocument.Paragraphs(3).Alignment = wdAlignParagraphCenter
End Sub

Sub Main()
    Setup
    color
End Sub
