Execute function at the filter change

[https://stackoverflow.com/questions/15904230/vbatrigger-macro-on-column-filter]

```vba
Private Sub Worksheet_Calculate()
 If ActiveSheet.Name = "Sheet1" Then
     If Cells(Rows.Count, 1).End(xlUp).Row = 1 Then
         MsgBox "No data available"
     Else
         MsgBox "There are filtering results"
     End If
 End If
End Sub
```
