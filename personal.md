```vba
'https://learn.microsoft.com/en-us/office/troubleshoot/excel/run-macro-cells-change

Private Sub Worksheet_Calculate()
'https://stackoverflow.com/questions/15904230/vbatrigger-macro-on-column-filter

If ActiveSheet.Name = "6000v2OUT" Then
 Range("N4").Value = Calculate_OR(Range("I7:I300"))
 Range("O4").Value = Calculate_OR(Range("L7:L300"))
 Range("P4").Value = Calculate_OR(Range("M7:M300"))
 
 
' https://excelmacromastery.com/excel-vba-array/#Get_the_Array_Size

'    If Cells(Rows.Count, 1).End(xlUp).Row = 1 Then
'        MsgBox "No data available"
'    Else
'        MsgBox "There are filtering results"
'    End If
End If
End Sub


Function ConcatenateVisible(xRg As Variant, sptChar As String)
'Updateby Extendoffice 20160922
'https://www.mrexcel.com/board/threads/vba-to-textjoin-visible-cells-only.1137178/

    Dim rg As Range
    For Each rg In xRg
        If (rg.EntireRow.Hidden = False) And (rg.EntireColumn.Hidden = False) Then
            ConcatenateVisible = ConcatenateVisible & rg.Value & sptChar
        End If
    Next
    ConcatenateVisible = Left(ConcatenateVisible, Len(ConcatenateVisible) - Len(sptChar))
End Function


Function Calculate_OR(xRg As Variant, sptChar As String)
    Dim dict                   'Create a variable
    Set dict = CreateObject("Scripting.Dictionary")
    Dim rg As Range
    For Each rg In xRg
        If (rg.EntireRow.Hidden = False) And (rg.EntireColumn.Hidden = False) Then
            Dim arr As Variant
            arr = Split(rg.Value, ",")
            For Each Item In arr
                If Not dict.Exists(Item) Then
                    dict.Add Item, Item
                End If
            Next Item
        End If
    Next
    Calculate_OR = dict.Count
End Function

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





Public Sub Remove_Cell_MultiLine()
    Dim rgEx As Object
    Set rgEx = CreateObject("VBScript.RegExp")
    rgEx.Pattern = "\n+"
    rgEx.MultiLine = True
    rgEx.Global = True
                    
    
    Dim rng As Range, cell As Range
    Set rng = Selection
        For Each cell In rng
            If Not (IsEmpty(cell.Value)) And Application.IsText(cell) Then
                    ' Adjust Selection part with suitable Excel Range References.
                cell.Value = rgEx.Replace(cell.Value, " ")
            End If
        Next cell
        

End Sub
Public Function CellType(c)
' https://stackoverflow.com/questions/15980484/checking-data-types-in-a-range
' https://spreadsheetpage.com/use-excel-tips-tricks/
    Application.Volatile
    Select Case True
        Case IsEmpty(c): CellType = "Blank"
        Case Application.IsText(c): CellType = "Text"
        Case Application.IsLogical(c): CellType = "Logical"
        Case Application.IsErr(c): CellType = "Error"
        Case IsDate(c): CellType = "Date"
        Case InStr(1, c.Text, ":") <> 0: CellType = "Time"
        Case InStr(1, c.Text, "%") <> 0: CellType = "Percentage"
        Case IsNumeric(c): CellType = "Value"
    End Select
End Function


Sub jr_cell_merge()
' Z zaznaczonych komórek robi jedną ze wspólną zawartością (string połączony spacją)
Dim rng As Range, cell As Range
Set rng = Selection
    For Each cell In rng
        If cell.Row > Selection.Row Then
            If Not (IsEmpty(cell.Value)) Then
                rng.Cells(1, 1).Value = rng.Cells(1, 1).Value & " " & cell.Cells(1, 1).Value
                cell.Cells(1, 1).Value = ""
            End If
        End If
    Next cell
End Sub


Sub ChangeDateFormat()
'
' Ustawia odpowiedni rodzaj daty
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


Sub RowHeightInCentimeters()
Dim cm As Single
cm = Application.InputBox("Enter Row Height in Centimeters", _
"Row Height (cm)", Type:=1)
If cm Then
Selection.RowHeight = Application.CentimetersToPoints(cm)
End If
End Sub


Sub ColumnWidthInCentimeters()
Dim cm As Single, points As Integer, savewidth As Integer
Dim lowerwidth As Integer, upwidth As Integer, curwidth As Integer
Dim Count As Integer
Application.ScreenUpdating = False
cm = Application.InputBox("Enter Column Width in Centimeters", _
"Column Width (cm)", Type:=1)
If cm = False Then Exit Sub
points = Application.CentimetersToPoints(cm)
savewidth = ActiveCell.ColumnWidth
ActiveCell.ColumnWidth = 255
If points > ActiveCell.Width Then
MsgBox "Width of " & cm & " is too large." & Chr(10) & _
"The maximum value is " & _
Format(ActiveCell.Width / 28.3464566929134, _
"0.00"), vbOKOnly + vbExclamation, "Width Error"
ActiveCell.ColumnWidth = savewidth
Exit Sub
End If
lowerwidth = 0
upwidth = 255
ActiveCell.ColumnWidth = 127.5
curwidth = ActiveCell.ColumnWidth
Count = 0
While (ActiveCell.Width <> points) And (Count < 20)
If ActiveCell.Width < points Then
lowerwidth = curwidth
Selection.ColumnWidth = (curwidth + upwidth) / 2
Else
upwidth = curwidth
Selection.ColumnWidth = (curwidth + lowerwidth) / 2
End If
curwidth = ActiveCell.ColumnWidth
Count = Count + 1
Wend
End Sub

' *****************************************************
' https://www.contextures.com/excellatitudelongitude.html
' *****************************************
Private Const PI = 3.14159265358979
Private Const EPSILON As Double _
    = 0.000000000001
'======================================
Public Function distVincenty(ByVal _
  lat1 As Double, ByVal lon1 As Double, _
    ByVal lat2 As Double, _
      ByVal lon2 As Double) As Double
'INPUTS: Latitude and Longitude of
'  initial and destination points
'  in decimal format.
'OUTPUT: Distance between the
'  two points in Meters.
'
'=============================
' Calculate geodesic distance (in m)
'  between two points specified by
'  latitude/longitude (in numeric
'  [decimal] degrees)
' using Vincenty inverse formula
'  for ellipsoids
'==============================
' Code has been ported by lost_species
'  from www.aliencoffee.co.uk to VBA
'  from javascript published at:
' https://www.movable-type.co.uk/scripts
'  /latlong-vincenty.html
' * from: Vincenty inverse formula -
'  T Vincenty, "Direct and Inverse
'  Solutions of Geodesics on the
' *       Ellipsoid with application
'  of nested equations", Survey Review,
'  vol XXII no 176, 1975
' *       https://www.ngs.noaa.gov/
'               PUBS_LIB/inverse.pdf
'Additional Reference:
'  https://en.wikipedia.org/wiki/
'      Vincenty%27s_formulae
'=============================
' Copyright lost_species 2008 LGPL
'  https://www.fsf.org/licensing/
'      licenses/lgpl.html
'=============================
' Code modifications to prevent
'  "Formula Too Complex" errors in
'  Excel (2010) VBA implementation
' provided by Jerry Latham,
'  Microsoft MVP Excel, 2005-2011
' July 23 2011
'=============================

  Dim low_a As Double
  Dim low_b As Double
  Dim f As Double
  Dim L As Double
  Dim U1 As Double
  Dim U2 As Double
  Dim sinU1 As Double
  Dim sinU2 As Double
  Dim cosU1 As Double
  Dim cosU2 As Double
  Dim lambda As Double
  Dim lambdaP As Double
  Dim iterLimit As Integer
  Dim sinLambda As Double
  Dim cosLambda As Double
  Dim sinSigma As Double
  Dim cosSigma As Double
  Dim sigma As Double
  Dim sinAlpha As Double
  Dim cosSqAlpha As Double
  Dim cos2SigmaM As Double
  Dim c As Double
  Dim uSq As Double
  Dim upper_A As Double
  Dim upper_B As Double
  Dim deltaSigma As Double
  Dim s As Double ' final result,
'  will be returned rounded to
'  3 decimals (mm).
'added by JLatham to break up
'  "Too Complex" formulas
'into pieces to properly calculate
'  those formulas as noted below
'and to prevent overflow errors when
'  using Excel 2010 x64 on
'  Windows 7 x64 systems
  Dim P1 As Double ' used to calculate
'  a portion of a complex formula
  Dim P2 As Double ' used to calculate
'  a portion of a complex formula
  Dim P3 As Double ' used to calculate
'  a portion of a complex formula

'See https://en.wikipedia.org/wiki
'  /World_Geodetic_System
'for information on various Ellipsoid
'  parameters for other standards.
'low_a and low_b in meters
' === GRS-80 ===
' low_a = 6378137
' low_b = 6356752.314245
' f = 1 / 298.257223563
'
' === Airy 1830 ===  Reported best
'  accuracy for England
'  and Northern Europe.
' low_a = 6377563.396
' low_b = 6356256.910
' f = 1 / 299.3249646
'
' === International 1924 ===
' low_a = 6378388
' low_b = 6356911.946
' f = 1 / 297
'
' === Clarke Model 1880 ===
' low_a = 6378249.145
' low_b = 6356514.86955
' f = 1 / 293.465
'
' === GRS-67 ===
' low_a = 6378160
' low_b = 6356774.719
' f = 1 / 298.247167

'== WGS-84 Ellipsoid Parameters ===
  low_a = 6378137       ' +/- 2m
  low_b = 6356752.3142
  f = 1 / 298.257223563
'=========================
  L = toRad(lon2 - lon1)
  U1 = Atn((1 - f) * Tan(toRad(lat1)))
  U2 = Atn((1 - f) * Tan(toRad(lat2)))
  sinU1 = Sin(U1)
  cosU1 = Cos(U1)
  sinU2 = Sin(U2)
  cosU2 = Cos(U2)

  lambda = L
  lambdaP = 2 * PI
  iterLimit = 100 ' can be set
'  as low as 20 if desired.

  While (Abs(lambda - lambdaP) > _
      EPSILON) And (iterLimit > 0)
    iterLimit = iterLimit - 1

    sinLambda = Sin(lambda)
    cosLambda = Cos(lambda)
    sinSigma = Sqr(((cosU2 * sinLambda) _
        ^ 2) + ((cosU1 * sinU2 - sinU1 _
        * cosU2 * cosLambda) ^ 2))
    If sinSigma = 0 Then
     distVincenty = 0 'co-incident points
      Exit Function
    End If
    cosSigma = sinU1 * sinU2 + cosU1 _
      * cosU2 * cosLambda
    sigma = Atan2(cosSigma, sinSigma)
    sinAlpha = cosU1 * cosU2 * _
      sinLambda / sinSigma
    cosSqAlpha = 1 - sinAlpha * sinAlpha

    If cosSqAlpha = 0 Then 'check for
    'a divide by zero
      cos2SigmaM = 0 '2 points on equator
    Else
      cos2SigmaM = cosSigma - 2 _
        * sinU1 * sinU2 / cosSqAlpha
    End If

    c = f / 16 * cosSqAlpha * (4 + f _
        * (4 - 3 * cosSqAlpha))
    lambdaP = lambda

'the original calculation is
'  "Too Complex" for Excel VBA
'  to deal with
'so it is broken into segments
'  to calculate without that issue
'the original implementation
'  to calculate lambda
'lambda = L + (1 - C) * f * sinAlpha * _
  (sigma + C * sinSigma * (cos2SigmaM
'  + C * cosSigma * (-1 + 2
'  * (cos2SigmaM ^ 2))))
      'calculate portions
    P1 = -1 + 2 * (cos2SigmaM ^ 2)
    P2 = (sigma + c * sinSigma * _
      (cos2SigmaM + c * cosSigma * P1))
    'complete the calculation
    lambda = L + (1 - c) * f _
      * sinAlpha * P2

  Wend

  If iterLimit > 1 Then
   MsgBox _
   "iteration limit has been reached," _
        & " something didn't work."
    Exit Function
  End If

  uSq = cosSqAlpha * (low_a ^ 2 _
    - low_b ^ 2) / (low_b ^ 2)

'the original calculation is
'  "Too Complex" for Excel VBA
'  to deal with
'so it is broken into segments to
'  calculate without that issue
  'the original implementation to
'  calculate upper_A
  'upper_A = 1 + uSq / 16384 *
'  (4096 + uSq * (-768 + uSq *
'  (320 - 175 * uSq)))
  'calculate one piece of the equation
  P1 = (4096 + uSq * (-768 _
    + uSq * (320 - 175 * uSq)))
  'complete the calculation
  upper_A = 1 + uSq / 16384 * P1

  'oddly enough, upper_B calculates
'  without any issues - JLatham
  upper_B = uSq / 1024 * (256 + uSq _
    * (-128 + uSq * (74 - 47 * uSq)))

'the original calculation is
'  "Too Complex" for Excel VBA
'  to deal with
'so it is broken into segments to
'  calculate without that issue
  'the original implementation to
'  calculate deltaSigma
  'deltaSigma = upper_B * sinSigma *
'  (cos2SigmaM + upper_B / 4 *
'  (cosSigma * (-1 + 2
'   * cos2SigmaM ^ 2) _
     - upper_B / 6 * cos2SigmaM *
'   (-3 + 4 * sinSigma ^ 2) *
'   (-3 + 4 *cos2SigmaM ^ 2)))
  'calculate pieces of the
'  deltaSigma formula
  'broken into 3 pieces to prevent
'  overflow error that may occur in
  'Excel 2010 64-bit version.
  P1 = (-3 + 4 * sinSigma ^ 2) * _
    (-3 + 4 * cos2SigmaM ^ 2)
  P2 = upper_B * sinSigma
  P3 = (cos2SigmaM + upper_B / 4 * _
   (cosSigma * (-1 + 2 _
      * cos2SigmaM ^ 2) - _
     upper_B / 6 * cos2SigmaM * P1))
  'complete deltaSigma calculation
  deltaSigma = P2 * P3

  'calculate the distance
  s = low_b * upper_A * _
    (sigma - deltaSigma)
  'round distance to millimeters
  distVincenty = Round(s, 3)

End Function
'======================================
Function SignIt(Degree_Dec As String) _
    As Double
'Input:  a string representation of
'  a lat or long in the
'         format of 10° 27' 36" S/N
'  or 10~ 27' 36" E/W
'OUTPUT:  signed decimal value
'  ready to convert to radians
'
  Dim decimalValue As Double
  Dim tempString As String
  tempString = UCase(Trim(Degree_Dec))
  decimalValue = _
    Convert_Decimal(tempString)
  If Right(tempString, 1) = "S" _
    Or Right(tempString, 1) = "W" Then
    decimalValue = decimalValue * -1
  End If
  SignIt = decimalValue
End Function
'======================================
Function Convert_Degree(Decimal_Deg) _
  As Variant
'source: https://support.microsoft.com/
'  kb/213449
'
'converts a decimal degree
'  representation to deg min sec
'as 10.46 returns 10° 27' 36"
'
  Dim degrees As Variant
  Dim minutes As Variant
  Dim seconds As Variant
  With Application
     'Set degree to Integer of
'  Argument Passed
     degrees = Int(Decimal_Deg)
     'Set minutes to 60 times the
'  number to the right
     'of the decimal for the
'  variable Decimal_Deg
     minutes = (Decimal_Deg - _
    degrees) * 60
     'Set seconds to 60 times the
'  number to the right of the
     'decimal for the variable Minute
     seconds = Format(((minutes - _
    Int(minutes)) * 60), "0")
     'Returns the Result of degree
'  conversion
    '(for example, 10.46 = 10° 27' 36")
     Convert_Degree = " " & degrees _
    & "° " & Int(minutes) & "' " _
         & seconds + Chr(34)
  End With
End Function
'======================================
Function Convert_Decimal _
    (Degree_Deg As String) As Double
'source: https://support.microsoft.com/
'  kb/213449
   ' Declare the variables to be
'  double precision floating-point.
   ' Converts text angular entry to
'  decimal equivalent, as:
   ' 10° 27' 36" returns 10.46
   ' alternative to ° is permitted:
'  Use ~ instead, as:
   ' 10~ 27' 36" also returns 10.46
   Dim degrees As Double
   Dim minutes As Double
   Dim seconds As Double
   '
   'modification by JLatham
   'allow the user to use the ~
'  symbol instead of ° to denote degrees
   'since ~ is available from the
'  keyboard and ° has to be entered
   'through [Alt] [0] [1] [7] [6]
'  on the number pad.
   Degree_Deg = Replace(Degree_Deg, _
    "~", "°")

   ' Set degree to value before
'  "°" of Argument Passed.
   degrees = Val(Left(Degree_Deg, _
    InStr(1, Degree_Deg, "°") - 1))
   ' Set minutes to the value between
'  the "°" and the "'"
   ' of the text string for the variable
'   Degree_Deg divided by
   ' 60. The Val function converts the
'   text string to a number.
   minutes = Val(Mid(Degree_Deg, _
    InStr(1, Degree_Deg, "°") + 2, _
      InStr(1, Degree_Deg, "'") - _
    InStr(1, Degree_Deg, "°") - 2)) / 60
   ' Set seconds to the number to the
'  right of "'" that is
   ' converted to a value and then
'  divided by 3600.
   seconds = Val(Mid(Degree_Deg, _
    InStr(1, Degree_Deg, "'") + _
      2, Len(Degree_Deg) - _
    InStr(1, Degree_Deg, "'") - 2)) _
    / 3600
   Convert_Decimal = degrees _
    + minutes + seconds
End Function
'======================================
Private Function toRad(ByVal _
    degrees As Double) As Double
    toRad = degrees * (PI / 180)
End Function
'======================================
Private Function Atan2( _
    ByVal x As Double, _
    ByVal y As Double) As Double
 ' code nicked from:
 ' https://en.wikibooks.org/wiki/
'  Programming:Visual_Basic_Classic/
'  Simple_Arithmetic
'  #Trigonometrical_Functions
 ' If you re-use this watch out:
'  the x and y have been reversed from
'  typical use.
    If y > 0 Then
        If x >= y Then
            Atan2 = Atn(y / x)
        ElseIf x <= -y Then
            Atan2 = Atn(y / x) + PI
        Else
        Atan2 = PI / 2 - Atn(x / y)
    End If
        Else
            If x >= -y Then
            Atan2 = Atn(y / x)
        ElseIf x <= y Then
            Atan2 = Atn(y / x) - PI
        Else
            Atan2 = -Atn(x / y) - PI / 2
        End If
    End If
End Function
'======================================


Sub HyperAdd()

    'Converts each text hyperlink selected into a working hyperlink
    
    ' https://bettersolutions.com/excel/cells-ranges/vba-finding-last-row-column.htm
    Dim llastrow As Long
        'llastrow = Range(Range("A65536").End(XlDirection.xlUp).Address).Row
        llastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        
    
 '   MsgBox "The last row with data is number: " & Selection.End(xlUp).Row
    Debug.Print "---" & llastrow&; "--" & Range(Range(Cells(1, 1), Cells(Selection.Rows.Count, Selection.Columns.Count)).End(XlDirection.xlUp).Address)
    Dim s As Range
    'Set s = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    For Each xCell In Selection
        Debug.Print xCell.Column & ":" & xCell.Row & ":" & xCell.Value
        If Not (IsEmpty(xCell.Value)) And InStr(1, xCell.Value, "http") Then
            ActiveSheet.Hyperlinks.Add Anchor:=xCell, Address:=xCell.Formula, TextToDisplay:=Mid(Mid(xCell, InStrRev(xCell, "/") + 1), InStrRev(xCell, "\") + 1)
            Debug.Print xCell.Column & ":" & xCell.Row & "match"
            
        End If
 '       If xCell.Row > llastrow Then
 '           Exit For
 '       End If
            
        
    Next xCell

End Sub


Function GiveMeURL(rng As Range) As String
On Error Resume Next
GiveMeURL = rng.Hyperlinks(1).Address
End Function

Function GiveMeURLLink(rng As Range) As String
On Error Resume Next
GiveMeURL = rng.Hyperlinks(1).SubAddress
End Function


Sub DownloadURL(fURL As String)

' pobiera podany url do "E:\Excel_download\"
' Katalog nalezy stworzyc recznie wczesniej !!!
' https://stackoverflow.com/questions/14675830/how-to-download-all-links-in-column-a-in-a-folder

Dim myURL As String
myURL = "https://ocdn.eu/pulscms-transforms/1/2yTk9kpTURBXy9jNTNlZTYzZWMzZjA2ZWZiMTJmMTQ5ZTQ3YTY1ZThhNC5qcGeTlQMAzFrNBkDNA4STBc0DFM0BvJMJpjkzNDg4NQbeAAGhMAU/eksplozja-w-miejscowosci-przewodow.webp"

Dim WinHTTPReq As Object
Set WinHTTPReq = CreateObject("Microsoft.XMLHTTP")
Call WinHTTPReq.Open("GET", fURL, False)
WinHTTPReq.send

If WinHTTPReq.Status = 200 Then
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1
    oStream.Write WinHTTPReq.responseBody
    oStream.SaveToFile "E:\Excel_download\" + FunctionGetFileName(fURL), 2
    oStream.Close
End If

End Sub

Function FunctionGetFileName(FullPath As String) As String
'Update 20140210
Dim splitList As Variant
splitList = VBA.Split(FullPath, "/")
FunctionGetFileName = splitList(UBound(splitList, 1))
End Function



Sub DownloadLinks()

' sluzy do sciagniecia linkow z Zaznaczonych komorek
' przydatne np. SharePoint --> export to Excel
' zaznaczasz kolumne --> i się sciąga --> tu do katalogu E:\Excel_download\ DownloadURL

Dim rng As Range
Dim download As Boolean
download = True  ' jak false to robi na sucho (czyli bez sciagniecia plikow)
UserForm1.Show
UserForm1.Frame1.Width = 0
Dim x As Long
x = UserForm1.Width - (2 * UserForm1.Frame1.Left)
Dim Files_processed As Long
Files_processed = 0
Dim Files_max As Integer

' wylicza ile jest rzeczywistych linkow (jak zaznaczy sie cala kolumne to jest ponad milion celek

Files_max = 0
UserForm1.TextBox1 = "Calculate items "
For Each rng In Selection
    Files_processed = Files_processed + 1
    If rng.Hyperlinks.Count > 0 Then Files_max = Files_max + 1
    If Files_processed Mod 70000 = 1 Then
        UserForm1.Frame1.Width = Files_processed / Selection.Count * x
        UserForm1.Label1.Caption = Int(Files_processed / Selection.Count * 100) & "%"
        DoEvents
    End If
    
Next rng

' wlasciwa petla
Files_processed = 0

For Each rng In Selection
    Files_processed = Files_processed + 1
    If rng.Hyperlinks.Count > 0 Then  ' czy jest to link
        Dim url1 As String
        url1 = rng.Hyperlinks(1).Address
        Dim fil1 As String
        fil1 = FunctionGetFileName(rng.Hyperlinks(1).Address)
       ' rng.Offset(0, 1).Value = rng.Hyperlinks(1).Address
       ' rng.Offset(0, 2).Value = FunctionGetFileName(rng.Hyperlinks(1).Address)
        UserForm1.TextBox1 = fil1 & vbCrLf & url1 & vbCrLf & Files_processed & "/" & Files_max
        UserForm1.Frame1.Width = Files_processed / Files_max * x
        UserForm1.Label1.Caption = Files_processed & "/" & Files_max & " " & Int(Files_processed / Files_max * 100) & "%"
        DoEvents ' aby sie wyswietlal pasek postepu
        If download Then
            DownloadURL (rng.Hyperlinks(1).Address)
        End If
    End If
Next rng
UserForm1.Hide
End Sub



Public Function ISOWEEKNUMBER(InDate As Date) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Return week number ( ISO )                                          '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim D As Date

    D = DateSerial(Year(InDate - Weekday(InDate - 1) + 4), 1, 3)
    ISOWEEKNUMBER = Int((InDate - D + Weekday(D) + 5) / 7)

End Function

Function FirstDayOfWeek(nWeek As Integer, ReportYear As Integer) As Date

    FirstDayOfWeek = (7 * (nWeek - 1) + DateSerial(ReportYear, 1, 1)) - Weekday(7 * (nWeek - 1) + DateSerial(ReportYear, 1, 1)) + 2

End Function




```
