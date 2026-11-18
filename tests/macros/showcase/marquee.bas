' Source: https://stackoverflow.com/a/29394947 (answer to "Make a cell scroll/marquee text right to left?", question 29394524)
' Author: JNevill (https://stackoverflow.com/users/2221001/jnevill)
' Licence: CC BY-SA 3.0 (Stack Exchange user content, posted 2015)
' What: scrolling marquee: shows a 10-character window of a message in Sheet1!M2 and moves it on one character a second.
' Entry: DoMarquee
' Run: it re-schedules itself with Application.OnTime every second and never stops (the answer says to comment out the
' OnTime line to stop it). It names the sheet by its code name, Sheet1.
' Uses Application.OnTime Now + TimeValue("00:00:01") with a macro name, InStr, Mid, Range.Value. Animates (one step a second).
' Code below is verbatim from the source.
Sub DoMarquee()
    Dim sMarquee As String
    Dim iWidth As Integer
    Dim iPosition As Integer
    Dim rCell As Range
    Dim iCurPos As Integer

    'Set the message to be displayed in this cell
    sMarquee = "This is a scrolling Marquee."

    'Set the cell width (how many characters you want displayed at once
    iWidth = 10

    'Which cell are we doing this in?
    Set rCell = Sheet1.Range("M2")

    'determine where we are now with the message.
    '   instr will return the position of the first
    '   character where the current cell value is in
    '   the marquee message
    iCurPos = InStr(1, sMarquee, rCell.Value)

    'If we are position 0, then there is no message, so start over
    '   otherwise, bump the message to the next characterusing mid
    If iCurPos = 0 Then
        'Start it over
        rCell.Value = Mid(sMarquee, 1, iWidth)
    Else
        'bump it
        rCell.Value = Mid(sMarquee, iCurPos + 1, iWidth)
    End If

    'Set excel up to run this thing again in a second or two or whatever
    Application.OnTime Now + TimeValue("00:00:01"), "DoMarquee"
End Sub
