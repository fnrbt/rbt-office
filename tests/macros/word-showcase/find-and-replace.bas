Attribute VB_Name = "FindReplace"
' Name: Finding and replacing text or formatting
' Source: https://learn.microsoft.com/en-us/office/vba/word/concepts/customizing-word/finding-and-replacing-text-or-formatting
' Author: Microsoft (MicrosoftDocs/VBA-Docs)
' Licence: code samples MIT (MicrosoftDocs/VBA-Docs LICENSE-CODE), text CC BY 4.0
' What it does: the page's samples one after the other: bold the first "blue" (Find on a Range and
'               .Parent), replace every "hi" with "hello" (Selection.Find, wdReplaceAll), then remove bold
'               formatting from the document (Find.Font.Bold = True, Replacement.Font.Bold = False).
' Adaptations: the samples (fragments on the page) are wrapped in Subs; Setup, Note and Main added: Main
'              writes a coloured line after each step.
' Needs: text with "blue", "hi" and bold words (Setup types it)
' Entry point: Main
Sub BoldFirstBlue()
With ActiveDocument.Content.Find
 .Text = "blue"
 .Forward = True
 .Execute
 If .Found = True Then .Parent.Bold = True
End With
End Sub

Sub HiToHello()
With Selection.Find
 .ClearFormatting
 .Text = "hi"
 .Replacement.ClearFormatting
 .Replacement.Text = "hello"
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With
End Sub

Sub RemoveBold()
With ActiveDocument.Content.Find
 .ClearFormatting
 .Font.Bold = True
 With .Replacement
 .ClearFormatting
 .Font.Bold = False
 End With
 .Execute FindText:="", ReplaceWith:="", _
 Format:=True, Replace:=wdReplaceAll
End With
End Sub

Sub Setup()
    ActiveDocument.Content.Text = "hi there. The sky is blue, the sea is blue, and BOLD words stand out." & vbCr & _
        "Say hi to the team: hi Ada, hi Grace, hi Linus."
    ActiveDocument.Content.Font.Size = 16
    ActiveDocument.Paragraphs(1).Range.Words(15).Bold = True
End Sub

Sub Note(text As String, color As Long)
    ActiveDocument.Content.InsertParagraphAfter
    ActiveDocument.Content.InsertAfter text
    With ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count).Range.Font
        .Bold = False
        .Color = color
        .Italic = True
        .Size = 11
    End With
End Sub

Sub Main()
    Setup
    BoldFirstBlue
    Note "1. Find on a Range: the first ""blue"" is bold.", wdColorBlue
    Selection.HomeKey wdStory
    HiToHello
    Note "2. Every ""hi"" replaced with ""hello"".", wdColorGreen
    RemoveBold
    Note "3. Formatting replaced: no bold text is left.", wdColorDarkRed
End Sub
