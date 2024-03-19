Attribute VB_Name = "guper"
Option Explicit

Sub openphto()
Attribute openphto.VB_ProcData.VB_Invoke_Func = "Q\n14"
    Dim c As Range, s As Range
    Set s = Cells(ActiveCell.Row, ActiveCell.Column)
    For Each c In Selection
        ThisWorkbook.FollowHyperlink ("C:\Users\ƒј ј– 2\Desktop\ра общее\" & c)
    Next
    Cells(s.Row, s.Column).Select
End Sub
Private Sub ѕереходѕо√иперссылке»зјктивнойячейки()
    Dim url$
 
    ' получаем гиперссылку из активной €чейки листа
   url = FormulaHyperlink(ActiveCell)
     
    ' если гиперссылка найдена - переходим по ней
   If Len(url) Then ThisWorkbook.FollowHyperlink (url)
     
End Sub
 
Function FormulaHyperlink(ByRef cell As Range) As String
    If cell.HasFormula And (cell.Hyperlinks.Count = 0) Then
        If cell.Formula Like "=HYPERLINK*" Then
            FormulaHyperlink = Evaluate(Mid$(Split(cell.Formula, ",")(0), 12))
        End If
    End If
End Function

Sub killphoto()
Attribute killphoto.VB_ProcData.VB_Invoke_Func = "K\n14"
    Dim c As Range, s As Range
    Set s = Cells(ActiveCell.Row, ActiveCell.Column)
    For Each c In Selection
        Kill "C:\Users\ƒј ј– 2\Desktop\ра общее\" & c
    Next
    Cells(s.Row, s.Column).Select
End Sub

Sub test2()
    Name Cells(ActiveCell.Row, 2) & Cells(ActiveCell.Row, 1) As Cells(ActiveCell.Row, 2) & Cells(ActiveCell.Row, 4)
End Sub

