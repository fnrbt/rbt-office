' chart-probe.vbs WORKBOOK OUT.tsv [PNGDIR]: what Excel 2019 makes of every chart of every sheet of a workbook, one
' line per fact (tab separated): the chart's size, each axis' scale (minimum, maximum, major unit, the category axis'
' label spacing and kind), the plot area outside and inside, the title and the legend (and each entry), each data
' label's text and place. Points; positions from the chart area's top left. With PNGDIR each chart is also exported
' there as <sheet>-<chart>.png (Chart.Export, at 100 %). Run it in the desktop session (run-vbs-in-session.ps1).
Option Explicit
Dim book, out, pngDir, fso, f, xl, wb, ws, co, ch, ax, i, j, k, g, s, p, n, t, e
book = WScript.Arguments(0)
out = WScript.Arguments(1)
pngDir = ""
If WScript.Arguments.Count > 2 Then pngDir = WScript.Arguments(2)
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile(out, True, True)
Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
xl.AskToUpdateLinks = False
xl.EnableEvents = False
Set wb = xl.Workbooks.Open(book, False, True)

Sub W(parts)
  f.WriteLine ws.Name & vbTab & co.Name & vbTab & parts
End Sub

On Error Resume Next
For Each ws In wb.Worksheets
  For Each co In ws.ChartObjects
    Set ch = co.Chart
    W "chart" & vbTab & ch.ChartType & vbTab & co.Width & vbTab & co.Height & vbTab & ch.ChartArea.Width & vbTab & ch.ChartArea.Height
    W "plot" & vbTab & ch.PlotArea.Left & vbTab & ch.PlotArea.Top & vbTab & ch.PlotArea.Width & vbTab & ch.PlotArea.Height & vbTab & ch.PlotArea.InsideLeft & vbTab & ch.PlotArea.InsideTop & vbTab & ch.PlotArea.InsideWidth & vbTab & ch.PlotArea.InsideHeight
    For g = 1 To 2
      For k = 1 To 2
        Set ax = Nothing
        Err.Clear
        Set ax = ch.Axes(k, g)
        If Err.Number = 0 And Not ax Is Nothing Then
          If k = 2 Then
            W "axis" & vbTab & g & vbTab & "value" & vbTab & ax.MinimumScale & vbTab & ax.MaximumScale & vbTab & ax.MajorUnit & vbTab & ax.Left & vbTab & ax.Top & vbTab & ax.Width & vbTab & ax.Height
          Else
            t = ""
            Err.Clear
            t = ax.MinimumScale & vbTab & ax.MaximumScale & vbTab & ax.MajorUnit
            If Err.Number <> 0 Then t = "-" & vbTab & "-" & vbTab & "-"
            Err.Clear
            e = ax.TickLabelSpacing
            If Err.Number <> 0 Then e = "-"
            Err.Clear
            n = ax.CategoryType
            If Err.Number <> 0 Then n = "-"
            Err.Clear
            W "axis" & vbTab & g & vbTab & "category" & vbTab & t & vbTab & ax.Left & vbTab & ax.Top & vbTab & ax.Width & vbTab & ax.Height & vbTab & "spacing " & e & vbTab & "type " & n
          End If
        End If
      Next
    Next
    Err.Clear
    If ch.HasTitle Then W "title" & vbTab & ch.ChartTitle.Text & vbTab & ch.ChartTitle.Left & vbTab & ch.ChartTitle.Top & vbTab & ch.ChartTitle.Width & vbTab & ch.ChartTitle.Height
    If ch.HasLegend Then
      W "legend" & vbTab & ch.Legend.Left & vbTab & ch.Legend.Top & vbTab & ch.Legend.Width & vbTab & ch.Legend.Height
      For j = 1 To ch.Legend.LegendEntries.Count
        Set e = ch.Legend.LegendEntries(j)
        W "entry" & vbTab & j & vbTab & e.Left & vbTab & e.Top & vbTab & e.Width & vbTab & e.Height & vbTab & e.Font.Size & vbTab & e.Font.Bold & vbTab & e.Font.Color
      Next
    End If
    For s = 1 To ch.SeriesCollection.Count
      Set p = ch.SeriesCollection(s)
      If p.HasDataLabels Then
        For j = 1 To p.Points.Count
          Err.Clear
          If p.Points(j).HasDataLabel Then
            With p.Points(j).DataLabel
              W "label" & vbTab & s & vbTab & j & vbTab & Replace(Replace(.Text, vbLf, "\n"), vbCr, "\n") & vbTab & .Left & vbTab & .Top & vbTab & .Width & vbTab & .Height & vbTab & .Font.Name & vbTab & .Font.Size & vbTab & .Font.Bold & vbTab & .Position
            End With
          End If
        Next
      End If
    Next
    If pngDir <> "" Then ch.Export pngDir & "\" & Replace(ws.Name, " ", "_") & "-" & Replace(co.Name, " ", "_") & ".png"
  Next
Next
On Error GoTo 0
f.Close
wb.Close False
xl.Quit
