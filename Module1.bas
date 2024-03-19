Attribute VB_Name = "Module1"
Option Explicit

Sub changePrices()
    Dim listName() As String, Tabs() As String, QuantityChecks As Integer, i As Integer, adress As String, BaseName As String
    adress = ActiveWorkbook.Path
    Workbooks.Open Filename:=Workbooks("PERSONAL.XLSB").Sheets("Смена цен").Cells(1, 4)
    BaseName = ActiveWorkbook.name
    QuantityChecks = Workbooks("PERSONAL.XLSB").Sheets("Смена цен").Cells(Rows.Count, 1).End(xlUp).Row - 1
    ReDim listName(QuantityChecks)
    ReDim Tabs(QuantityChecks)
    For i = 0 To QuantityChecks
        If Not IsBookOpen(adress + "\" + Workbooks("PERSONAL.XLSB").Sheets("Смена цен").Cells(i + 2, 1)) Then _
        Workbooks.Open Filename:=adress + "\" + Workbooks("PERSONAL.XLSB").Sheets("Смена цен").Cells(i + 2, 1)
        Application.Goto Cells(2, 6)
        ActiveCell.FormulaR1C1 = "=VLOOKUP(c1,'[" + BaseName + "]" + Workbooks("PERSONAL.XLSB").Sheets("Смена цен").Cells(i + 2, 2) + "'!C4:C5,2,0)"
    Next
End Sub
