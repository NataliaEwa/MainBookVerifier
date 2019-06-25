VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
1 ……………………………………
 
Sub M()
 
Call Dane
Call Zlacz
Call Unikatowe
Call Przypasuj
Call Dopasuj
 
End Sub
 
 
2 ……………………………………
 
Sub Dane()
 
' 1
 
Range("A:A").Clear
Range("C:C").Clear
Range("B:B").Clear
 
Range("A1").Value = "E"
Range("B1").Value = "R"
Range("C1").Value = "RAPORT"
 
Range("E:E").Clear
Range("E1").Value = "ALL"
 
Range("G:G").Clear
Range("J:J").Clear
Range("G1").Value = "Lp."
Range("J1").Value = "Lp."
 
Range("H:H").Clear
Range("H1").Value = "E+R"
 
Range("K:K").Clear
Range("K1").Value = "RAPORT"
 
Raport = "C:\Users\NLaskarzewska\Desktop\SGW-Raport z zaksi´gowanych dokumentów w Ksi´dze G∏ównej csv 07-2018.xls"
Ewidencja = "C:\Users\NLaskarzewska\Desktop\EWIDENCJA 07 2018.xls"
Ryczalt = "C:\Users\NLaskarzewska\Desktop\RYCZA¸T 07 2018.xls"
 
sprawdzenie = ActiveWorkbook.Name
Workbooks.Open Ewidencja
w = ActiveWorkbook.Name
Sheets(1).Range("C2:C65536").Copy
ThisWorkbook.Activate
Sheets("Arkusz1").Activate
Range("A2:A500").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Workbooks(w).Close
ThisWorkbook.Activate
 
sprawdzenie = ActiveWorkbook.Name
Workbooks.Open Ryczalt
w = ActiveWorkbook.Name
Sheets(1).Range("C2:C65536").Copy
ThisWorkbook.Activate
Sheets("Arkusz1").Activate
Range("B2:B500").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Workbooks(w).Close
ThisWorkbook.Activate
 
sprawdzenie = ActiveWorkbook.Name
Workbooks.Open Raport
w = ActiveWorkbook.Name
Sheets(1).Range("W13:W65536").Copy
ThisWorkbook.Activate
Sheets("Arkusz1").Activate
Range("C2:C500").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Workbooks(w).Close
ThisWorkbook.Activate
 
Range("A:A").Select
Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
 
 
Range("B:B").Select
Selection.Sort Key1:=Range("B2"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
 
 
Range("C:C").Select
Selection.Sort Key1:=Range("C2"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
 
 
End Sub
 
3 ……………………………………
 
Sub Zlacz()
 
' 2
 
ThisWorkbook.Activate
Range("A2:A500").Copy
    If Range("E1").Value = "" Then
        Range("E1").Select
    Else
        Range("E" & ActiveSheet.Cells.Rows.Count).End(xlUp).Offset(1, 0).Select
   End If
ActiveSheet.Paste
 
ThisWorkbook.Activate
Range("B2:B500").Copy
    If Range("E1").Value = "" Then
        Range("E1").Select
    Else
        Range("E" & ActiveSheet.Cells.Rows.Count).End(xlUp).Offset(1, 0).Select
    End If
ActiveSheet.Paste
 
ThisWorkbook.Activate
Range("C2:C500").Copy
    If Range("E1").Value = "" Then
        Range("E1").Select
    Else
        Range("E" & ActiveSheet.Cells.Rows.Count).End(xlUp).Offset(1, 0).Select
    End If
ActiveSheet.Paste
 
Range("E:E").Select
Selection.Sort Key1:=Range("E2"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
'Selection.AutoFilter
 
End Sub
 
 
4 ……………………………………
 
Sub Unikatowe()
 
' 3
 
Dim rng As Range
Set Zakres = Range("E2:E200")
Dim xLR As Long
Dim xLR2 As Long
Dim I As Integer
On Error Resume Next
Zakres.Copy Range("G2")
xLR = Zakres.Rows.Count + 1
ActiveSheet.Range("G2:G" & xLR).RemoveDuplicates Columns:=1, Header:=xlNo
xLR2 = Cells(Rows.Count, "E").End(xlUp).Row
 
For I = 1 To xLR2
    If ActiveSheet.Range("G3:T" & xLR).Cells(I).Value = "" Then
    ActiveSheet.Range("G3:G" & xLR).Cells(I).Value.Delete
    End If
Next
   
End Sub
 
5 ……………………………………
 
Sub Przypasuj()
 
' 4.1
ThisWorkbook.Activate
Dim CompareRange2 As Variant, a As Variant, b As Variant
Set CompareRange2 = Range("A2:B100")
Range("G2:G200").Select
 
For Each a In Selection
    For Each b In CompareRange2
            If a = b Then a.Offset(0, 1) = a
        Next b
    Next a
End
 
End Sub
 
6 ……………………………………
 
Sub Dopasuj()
 
' 4.2
ThisWorkbook.Activate
Range("G:G").Copy
Range("J:J").Select
ActiveSheet.Paste
 
ThisWorkbook.Activate
Dim CompareRange2 As Variant, a As Variant, b As Variant
Set CompareRange2 = Range("C2:C100")
Range("J2:J200").Select
 
For Each a In Selection
    For Each b In CompareRange2
            If a = b Then a.Offset(0, 1) = a
        Next b
    Next a
End
 
End Sub
