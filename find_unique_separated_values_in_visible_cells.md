Find Number of Uniques Values separated by "sptChar" (like ,) over the VISIBLE range 

```vba
Function FindNoOfUniques(xRg As Variant, sptChar As String)
    Dim dict                   'Create a variable
    Set dict = CreateObject("Scripting.Dictionary")
    Dim rg As Range
    For Each rg In xRg
        If (rg.EntireRow.Hidden = False) And (rg.EntireColumn.Hidden = False) Then
            Dim arr As Variant
            arr = Split(rg.Value, sptChar)
            For Each Item In arr
                If Not dict.Exists(Item) Then
                    dict.Add Item, 1
                End If
            Next Item
        End If
    Next
    FindNoOfUniques = dict.Count
End Function

```

Usage
```vba
Range("A1").Value = FindNoOfUniques(Range("B1:B4"),",")
```
| |A|B|C|
|--:|--|--|--|
|1|**4**|1,2,4| |
|2||1||
|3||3,4||
|4||2,3||
