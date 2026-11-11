' cscript probe-tints.vbs OUT.tsv: the colour Excel gives a theme colour with a tint (Interior.ThemeColor +
' TintAndShade, read back as Interior.Color), for every theme colour of the default theme and a set of tints:
' the reference for Xlsx.applyColorTint (tests/formula: tints against Excel). Lines: theme index (0-based, as
' in a file's theme="n"), tint, RRGGBB.
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile(WScript.Arguments(0), True)
Set x = CreateObject("Excel.Application"): x.Visible = False: x.DisplayAlerts = False
Set wb = x.Workbooks.Add(): Set sh = wb.Worksheets(1)
tints = Array(0.8, 0.6, 0.4, 0.2, 0.1, -0.05, -0.1, -0.15, -0.25, -0.35, -0.5, 0.5, 0.9, -0.9, _
  0.79998168889431442, 0.59999389629810485, 0.39997558519241921, -0.249977111117893, -0.499984740745262, _
  -0.14999847407452621, -0.34998626667073579, 0.499984740745262, -4.9989318521683403E-2, 0.249977111117893, 0.14999847407452621)
' xlThemeColorDark1..xlThemeColorAccent6 (1..10) are the theme colours in the file's order (lt1, dk1, lt2, dk2, accents): theme="n" is n - 1
fileIndex = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
For t = 1 To 10
  For Each tint In tints
    Set c = sh.Cells(1, 1)
    c.Interior.ThemeColor = t
    c.Interior.TintAndShade = tint
    v = c.Interior.Color
    out.WriteLine fileIndex(t - 1) & vbTab & tint & vbTab & Right("0" & Hex(v And 255), 2) & Right("0" & Hex((v \ 256) And 255), 2) & Right("0" & Hex((v \ 65536) And 255), 2)
  Next
Next
out.Close
wb.Close False: x.Quit
