Attribute VB_Name = "Leftovers"
Option Explicit

Private Base As String, BaseName As String, StartBookName As String, Address As String
Private listName() As String, MsngArr() As String, QuantityPrice As Byte
Private Kia As String, KiaName As String, KiaErr As String, KiaErrName As String
Private TrigerText As String, BonusColumn As Byte
Private QuantityChecks As Byte, CheckList() As String

Sub main()
    If IsBookOpen("Прайс Питер.xls") Then
        Call setValueForMoscow
        Call check
        Call prepare1C
        Call calculations
        Columns(1).AutoFilter Field:=8, Criteria1:="#Н/Д"
        Columns(1).AutoFilter Field:=5, Criteria1:="<>0"
    ElseIf IsBookOpen("Прайс Смарт-Д Питер.xls") Then
        Call setValueForSamara
        Call check
        Call deleteKia
        Call prepare1C
    Else
        MsgBox "Не открыт прайс для Питера!" & vbNewLine & "Откройте этот прайс и попробуйте снова."
        End
    End If
    Address = ActiveWorkbook.Path
    ND.Show
End Sub

Private Sub setValueForMoscow()
    Dim i As Integer
    
    Base = Workbooks("PERSONAL.XLSB").Sheets("Москва").Cells(1, 3)
    QuantityPrice = Workbooks("PERSONAL.XLSB").Sheets("Москва").Cells(Rows.Count, 1).End(xlUp).Row - 2
    BonusColumn = 1
    ReDim listName(QuantityPrice)
    ReDim MsngArr(QuantityPrice)
    For i = 0 To QuantityPrice
        listName(i) = Workbooks("PERSONAL.XLSB").Sheets("Москва").Cells(i + 2, 1)
    Next
    QuantityChecks = Workbooks("PERSONAL.XLSB").Sheets("Москва").Cells(Rows.Count, 5).End(xlUp).Row - 1
    ReDim CheckList(QuantityChecks)
    For i = 0 To QuantityChecks
        CheckList(i) = Workbooks("PERSONAL.XLSB").Sheets("Москва").Cells(i + 1, 5)
    Next
End Sub

Private Sub setValueForSamara()
    Dim i As Integer
    
    Base = Workbooks("PERSONAL.XLSB").Sheets("Самара").Cells(1, 3)
    Kia = Workbooks("PERSONAL.XLSB").Sheets("Самара").Cells(2, 3)
    KiaErr = Workbooks("PERSONAL.XLSB").Sheets("Самара").Cells(3, 3)
    QuantityPrice = Workbooks("PERSONAL.XLSB").Sheets("Самара").Cells(Rows.Count, 1).End(xlUp).Row - 2
    BonusColumn = 0
    ReDim listName(QuantityPrice)
    ReDim MsngArr(QuantityPrice)
    For i = 0 To QuantityPrice
        listName(i) = Workbooks("PERSONAL.XLSB").Sheets("Самара").Cells(i + 2, 1)
    Next
    TrigerText = Workbooks("PERSONAL.XLSB").Sheets("Самара").Cells(1, 5)
    QuantityChecks = Workbooks("PERSONAL.XLSB").Sheets("Самара").Cells(Rows.Count, 5).End(xlUp).Row - 2
    ReDim CheckList(QuantityChecks)
    For i = 0 To QuantityChecks
        CheckList(i) = Workbooks("PERSONAL.XLSB").Sheets("Самара").Cells(i + 2, 5)
    Next
End Sub

Private Sub prepare1C()
    Dim i As Integer
'Чтение и подготовка базы
    StartBookName = ActiveWorkbook.name
    Workbooks.Open Filename:=Base
    BaseName = ActiveWorkbook.name
    Rows("1:6").Delete Shift:=xlUp
    Rows(Cells(Rows.Count, 1).End(xlUp).Row).Delete Shift:=xlUp
    Columns(4).NumberFormat = "General"
    Columns(4).Value = Columns(4).Value

'Введение формулы
    Application.Goto Workbooks(StartBookName).Sheets("Лист1").Cells(2, 6 + BonusColumn)
    ActiveCell.FormulaR1C1 = "=VLOOKUP(c1,'[" + BaseName + "]TDSheet'!C4:C5,2,0)"

'Заполнение Таблицы
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 1).End(xlUp).Row, ActiveCell.Column))
    ActiveCell(1, 2).Select
    ActiveCell.FormulaR1C1 = "=RC[-" & 2 + BonusColumn & "]-RC[-1]"
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 1).End(xlUp).Row, ActiveCell.Column))
End Sub

Private Sub calculations()
    Dim i As Integer

    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        If Not IsError(Cells(i, 8)) Then
            If Cells(i, 8) <> 0 Then
                If Cells(i, 6) = 1 Then
                    If Cells(i, 7) <= 105 Then Cells(i, 5) = Cells(i, 7) _
                    Else If Cells(i, 1) <> "4h0951253a" Then Cells(i, 5) = 105 _
                    Else Cells(i, 5) = Cells(i, 7)
                Else
                    If Cells(i, 7) <= 300 Then Cells(i, 5) = Cells(i, 7) Else Cells(i, 5) = 300
                End If
            End If
        End If
    Next
End Sub

Sub FPTA(Optional v = 0)
    Dim i, n As Integer
    Dim Msng As String, test As Boolean
    
    If Not IsBookOpen(StartBookName) Then Workbooks.Open Filename:=Address + "\" + StartBookName
    If IsBookOpen(BaseName) Then Workbooks(BaseName).Close False
    With ActiveSheet
        If .FilterMode Then .ShowAllData
    End With
    Range(Columns(6 + BonusColumn), Columns(7 + BonusColumn)).Clear
    If BonusColumn = 0 Then Call prepareKia
    
    Msng = "Ошибки в прайсе(ах):" & vbNewLine
    For n = 0 To QuantityPrice
        MsngArr(n) = ""
    Next
    Application.DisplayAlerts = False
    For n = 0 To QuantityPrice
    'Открытие
        Workbooks.Open Filename:=Address + "\" + listName(n)
    'Заполнение Таблицы
        Application.Goto Cells(2, 6 + BonusColumn)
        If n <> 0 And BonusColumn = 0 Then Call deleteKia
        ActiveCell.FormulaR1C1 = "=VLOOKUP(c1,'[" + StartBookName + "]Лист1'!C1:C5,5,0)"
        Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 1).End(xlUp).Row, ActiveCell.Column))
        ActiveCell(1, 2).Select
        ActiveCell.FormulaR1C1 = "=RC[-" & 2 + BonusColumn & "]-RC[-1]"
        Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 1).End(xlUp).Row, ActiveCell.Column))
    'Расчеты
        For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
            If Not IsError(Cells(i, 7 + BonusColumn)) Then
                If Cells(i, 7 + BonusColumn) <> 0 Then Cells(i, 5) = Cells(i, 6 + BonusColumn)
            Else
                test = True
            End If
        Next
        If n <> 0 And BonusColumn = 0 Then
            Call addKia(ActiveWorkbook.name)
        End If
        If n = 3 And BonusColumn = 0 Then
            Columns(4).Clear
            Columns(4).Hidden = True
        End If
    'Чистка и закрытие
        Range(Columns(6 + BonusColumn), Columns(7 + BonusColumn)).Clear
        Cells(2, 7).Select
        ActiveWorkbook.Close True
    'Анализ ошибок
        If test Then MsngArr(n) = listName(n)
        test = False
    Next
    If BonusColumn = 0 Then
        Call addKia(StartBookName)
        Workbooks(KiaName).Close False
    End If
    Cells(2, 7).Select
    Workbooks(StartBookName).Save
'Подводим итоги
    For n = 0 To QuantityPrice
        If MsngArr(n) <> "" Then Msng = Msng & MsngArr(n) & vbNewLine
    Next
    If Msng <> "Ошибки в прайсе(ах):" & vbNewLine Then MsgBox Msng Else MsgBox "Ошибок нет."
    Application.DisplayAlerts = True
End Sub

Private Sub prepareKia()
    Dim i As Integer
    Workbooks.Open Filename:=KiaErr
    KiaErrName = ActiveWorkbook.name
    Workbooks.Open Filename:=Kia
    KiaName = ActiveWorkbook.name
    Rows(Cells(Rows.Count, 2).End(xlUp).Row + 1).Select
    For i = 2 To Cells(Rows.Count, 2).End(xlUp).Row
        If Cells(i, 6) <> "" And Cells(i, 6) <= 500 Then
            Union(Selection, Rows(i)).Select
        End If
    Next
    Selection.Delete Shift:=xlUp
    Application.Goto Cells(2, 7)
    ActiveCell.FormulaR1C1 = "=VLOOKUP(c3,'[" + KiaErrName + "]Лист1'!C2:C3,2,0)"
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 2).End(xlUp).Row, ActiveCell.Column))
    For i = 2 To Cells(Rows.Count, 2).End(xlUp).Row
        If (Not IsError(Cells(i, 7))) Then
            If (Not IsEmpty(Cells(i, 7))) Then
                Rows(i).Delete Shift:=xlUp
                i = i - 1
            End If
        End If
    Next
    Workbooks(KiaErrName).Close
    Application.Goto Cells(2, 7)
    ActiveCell.FormulaR1C1 = "=c[-1]*1.2"
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 2).End(xlUp).Row, ActiveCell.Column))
    ActiveCell(1, 2).Select
    ActiveCell.FormulaR1C1 = "=Round(c[-1],0)"
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 2).End(xlUp).Row, ActiveCell.Column))
    For i = 2 To Cells(Rows.Count, 2).End(xlUp).Row
        Cells(i, 6) = Cells(i, 8).Value
        If Cells(i, 2).Value = "KIA MOTORS" Then Cells(i, 2) = "Hyundai/Kia"
    Next
    Range(Columns(7), Columns(8)).Clear
    Range(Columns(5), Columns(6)).NumberFormat = "0"
    Range(Columns(5), Columns(6)).HorizontalAlignment = xlCenter
    Columns(2).HorizontalAlignment = xlCenter
    Columns(3).Cut
    Columns(2).Insert
    Columns(6).Cut
    Columns(5).Insert
    Columns(1).Delete Shift:=xlLeft
    With Range(Cells(1, 1), Cells(Cells(Rows.Count, 2).End(xlUp).Row, 5))
        .Font.name = "Calibri"
        .Font.Size = 11
        .RowHeight = 15
    End With
End Sub

Private Sub addKia(Book As String)
    Application.Goto Workbooks(KiaName).Sheets("TDSheet").Cells(1, 1)
    Range(Rows(2), Rows(Cells(Rows.Count, 1).End(xlUp).Row)).Copy
    Application.Goto Workbooks(Book).Sheets("лист1").Cells(1, 1)
    Application.Goto Workbooks(Book).Sheets("лист1").Cells(Cells(Rows.Count, 1).End(xlUp).Row + 1, 1)
    ActiveCell.PasteSpecial Paste:=xlPasteValues, SkipBlanks:=False, Transpose:=False
    Range(Cells(1, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, 5)).Borders.LineStyle = True
End Sub

Private Sub deleteKia()
    Dim i As Integer
    For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
        If Cells(i, 1) = TrigerText Then
            Range(Rows(i + 1), Rows(Cells(Rows.Count, 1).End(xlUp).Row)).Delete Shift:=xlUp
            Exit For
        End If
    Next
End Sub

Private Sub check()
    Dim i As Byte, list As String, Messeng As String, Desigion As Boolean
    Messeng = "Следующие книги будут закрыты:"
    list = ""
    For i = 0 To QuantityChecks
        If IsBookOpen(CheckList(i)) Then list = list + CheckList(i) + ", "
    Next
    Messeng = Messeng & vbNewLine & list & vbNewLine & "Сохранить?"
    If list <> "" Then
        Select Case MsgBox(Messeng, vbYesNoCancel, "Внимание!", vbQuestion, vbApplicationModal)
        Case 6
            Desigion = True
        Case 7
            Desigion = False
        Case Else
            End
        End Select
        For i = 0 To QuantityChecks
            If IsBookOpen(CheckList(i)) Then Workbooks(CheckList(i)).Close Desigion
        Next
    End If
End Sub
