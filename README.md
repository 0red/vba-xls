# vba-xls
```vba
Public Sub Master_all_to_one()
' Master_all_to_one Macro
Dim ws  As Worksheet, _
    LR1 As Long, _
    LR2 As Long
Application.ScreenUpdating = False
For Each ws In ActiveWorkbook.Worksheets
    If ws.Name <> "Master" And ws.Range("A2").Value = "Node" Then
        LR1 = Sheets("Master").Range("A" & Rows.Count).End(xlUp).Row + 1
        LR2 = ws.Range("A" & Rows.Count).End(xlUp).Row
        ws.Range("A2:K" & LR2).Copy Destination:=Sheets("Master").Range("A" & LR1)
    End If
Next ws
Application.ScreenUpdating = True
End Sub



Sub ChangeDateFormat()
'
' Macro1 Macro
'

'
    Dim rng As Range, cell As Range
    'Set rng = Range("A1:A3")
    Set rng = Selection
    For Each cell In rng
        If IsDate(cell.Value) Then
            cell.NumberFormat = "yyyy/mm/dd;@"
        End If
    Next cell
    
End Sub

Sub kasuj_lf()
    Dim rng As Range
    Set rng = Selection
    rng.Replace what:=Chr(10), lookat:=xlPart, replacement:="@@"
    rng.Replace what:="@@", lookat:=xlWhole, replacement:=""
End Sub
Sub kasuj_lfpocz()
    Dim rng As Range
    Set rng = Selection
    rng.Replace what:=Chr(10), lookat:=xlWhole, replacement:=""
    
    For Each cell In Selection
        If Mid(cell, 1, 1) = Chr(10) Then
            cell.Value = Mid(cell, 2, 999)
        End If
    Next cell
End Sub


Sub dodaj_lf()
    Dim rng As Range
    Set rng = Selection
    rng.Replace what:="@@", lookat:=xlPart, replacement:=Chr(10)
    rng.Replace what:="@@", lookat:=xlWhole, replacement:=""
End Sub
```
