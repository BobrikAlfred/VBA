Attribute VB_Name = "LeftoversMoscowWithLiart"
Option Explicit

Public decision As Boolean
Public list As String
Public lists(4) As String

Sub withLiart()
    Dim CatalogAdress, temporaryBook As String
    
    CatalogAdress = ActiveWorkbook.Path
    If CheckR Then
        Select Case Menu.choises
        Case 1
            temporaryBook = Workbooks.Add.name
            If Redundant.checkThis("Прайс Питер.xls") Then Workbooks("Прайс Питер.xls").Close False
            Workbooks.Open Filename:="C:\Users\ДАКАР 2\Dropbox\москва\Прайсы\Clear\Прайс Питер.xls"
            Workbooks(temporaryBook).Close
            Call prepare
            Call FPTA(CatalogAdress, False, True)
        Case 2
            temporaryBook = Workbooks.Add.name
            If checkThis("Прайс Питер.xls") Then Workbooks("Прайс Питер.xls").Close False
            Workbooks.Open Filename:="C:\Users\ДАКАР 2\Dropbox\москва\Прайсы\Clear\Прайс Питер.xls"
            Workbooks(temporaryBook).Close
            Call from1C("Прайс Питер.xls")
            Call check.ND("LeftoversMoscowWithLiart", CatalogAdress)
        Case Else
            End
        End Select
    End If
End Sub

Sub from1C(ByVal BookName As String)
    Dim art, BookAdress As String
    Dim i, EoF, c As Integer
    
'Чтение и подготовка базы
    Workbooks.Open Filename:="C:\Users\Public\Leftovers\Остатки товаров.xls"
    Rows("1:6").Delete Shift:=xlUp
    Rows(Cells(Rows.Count, 1).End(xlUp).Row).Delete Shift:=x1Up
    Columns(4).NumberFormat = "General"
    Columns(4).Value = Columns(4).Value

'Введение формулы
    Application.Goto Workbooks(BookName).Sheets("Лист1").Cells(2, 7)
    ActiveCell.FormulaR1C1 = "=VLOOKUP(c1,'[Остатки товаров.xls]TDSheet'!C4:C5,2,0)"

'Заполнение Таблицы
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 1).End(xlUp).Row, ActiveCell.Column))
    ActiveCell(1, 2).Select
    ActiveCell.FormulaR1C1 = "=RC[-3]-RC[-1]"
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 1).End(xlUp).Row, ActiveCell.Column))

'Расчеты
    EoF = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To EoF
        If Not IsError(Cells(i, 8)) Then
            If Cells(i, 8) <> 0 Then
                If Cells(i, 6) = 1 Then
                    If Cells(i, 7) <= 100 Then Cells(i, 5) = Cells(i, 7) _
                    Else If Cells(i, 1) <> "4h0951253a" Then Cells(i, 5) = 100 _
                    Else Cells(i, 5) = Cells(i, 7)
                Else
                    If Cells(i, 7) <= 300 Then Cells(i, 5) = Cells(i, 7) Else Cells(i, 5) = 300
                End If
            End If
        End If
    Next
'Проверка
    Columns(1).AutoFilter Field:=8, Criteria1:="#Н/Д"
    Columns(1).AutoFilter Field:=5, Criteria1:="<>0"
End Sub

Sub FPTA(ByVal BookAdress As String, ByVal from1C As Boolean, ByVal fromLiart As Boolean)
    Dim i, n, EoF, AmPrice As Integer
    Dim ListAll, Listadr, Msng, ListBook, SavedName As String
    Dim listName(15), MsngArr(15) As String
    Dim test As Boolean
    
    Application.Goto Workbooks("Прайс Питер.xls").Sheets("Лист1").Cells(1, 1)
'Задаем начальные параметры
    AmPrice = 15
    Listadr = ActiveWorkbook.Path
    listName(0) = "\Li-art ОСОБЫЕ цены\Прайс СД.xls"
    listName(1) = "\ZZap\ZZAP.xls"
    listName(2) = "\Авто_Партс\Прайс СД.xls"
    listName(6) = "\Авто_ТО  ОСОБЫЕ цены\Прайс СД.xls"
    listName(4) = "\Автодок\Прайс СД.xls"
    listName(7) = "\Автоформула Техно Цены питер\Прайс СД.xls"
    listName(3) = "\Вендор\Прайс СД.xls"
    listName(5) = "\Джапартс\SMDTB.xls"
    listName(8) = "\М-Партс\Москва СД.xls"
    listName(9) = "\Профит_Лига  ОСОБЫЕ цены\Прайс СД.xlsx"
    listName(10) = "\Стелс\Прайс СД.xls"
    listName(11) = "\Стратегия        ОСОБЫЕ цены\Прайс СД.xls"
    listName(12) = "\сфера\Прайс СД.xls"
    listName(13) = "\Фроза Цены Емекс\Прайс СД.xls"
    listName(14) = "\Прайс Емекс.xls"
    listName(15) = "\Прайс СД.xls"
'отслеживания ошибок
    Msng = "Ошибки в прайсе(ах):" & vbNewLine
    For n = 0 To AmPrice
        MsngArr(n) = ""
    Next

'Передаем остатки в прайсы
    For n = 0 To AmPrice
    'Открытие
        ListAll = Listadr + listName(n)
        Workbooks.Open Filename:=ListAll
        ListBook = ActiveWorkbook.name
    If from1C Then
    'Введение формулы
        Cells(2, 7).FormulaR1C1 = "=VLOOKUP(c1,'[Прайс Питер.xls]Лист1'!C1:C5,5,0)"
        Application.Goto Cells(2, 7)
    'Заполнение Таблицы
        Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 1).End(xlUp).Row, ActiveCell.Column))
        ActiveCell(1, 2).Select
        ActiveCell.FormulaR1C1 = "=RC[-3]-RC[-1]"
        Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 1).End(xlUp).Row, ActiveCell.Column))
    'Расчеты
        EoF = Cells(Rows.Count, 1).End(xlUp).Row
        For i = 1 To EoF
            If Not IsError(Cells(i, 8)) Then
                If Cells(i, 8) <> 0 Then Cells(i, 5) = Cells(i, 7)
            Else
                test = True
            End If
        Next
    'Чистка
        Range(Columns(7), Columns(8)).Clear
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs Filename:=("C:\Users\ДАКАР 2\Dropbox\москва\Прайсы\Clear" + listName(n))
        Application.DisplayAlerts = True
    End If
    If fromLiart Then
    'Добавление ЛиАрт
        If n = 6 Then Call DeleteDubl
        If 3 < n Then
            Call Copy
            Application.Goto Workbooks(ListBook).Sheets("Лист1").Cells(1, 1)
            Application.Goto Workbooks(ListBook).Sheets("Лист1").Cells(Cells(Rows.Count, 1).End(xlUp).Row + 1, 1)
            ActiveCell.PasteSpecial Paste:=xlPasteValues, SkipBlanks:=False, Transpose:=False
            'Границы
            Range(Cells(1, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, 6)).Borders.LineStyle = True
        End If
    End If
    'Закрытие
        Cells(2, 7).Select
        Application.DisplayAlerts = False
        SavedName = BookAdress + listName(n)
        ActiveWorkbook.SaveAs Filename:=(SavedName)
        ActiveWorkbook.Close False
        Application.DisplayAlerts = True
    'Анализ ошибок
        If test Then MsngArr(n) = listName(n)
        test = False
    Next
    
'Подводим итоги
    Application.DisplayAlerts = False
    Application.Goto Workbooks("Прайс Питер.xls").Sheets("Лист1").Cells(1, 1)
    ActiveWorkbook.SaveAs Filename:="C:\Users\ДАКАР 2\Dropbox\москва\Прайсы\Clear\Прайс Питер.xls"
    Application.DisplayAlerts = True
If fromLiart Then
    Call Copy
    Application.Goto Workbooks("Прайс Питер.xls").Sheets("Лист1").Cells(1, 1)
    Application.Goto Workbooks("Прайс Питер.xls").Sheets("Лист1").Cells(Cells(Rows.Count, 1).End(xlUp).Row + 1, 1)
    ActiveCell.PasteSpecial Paste:=xlPasteValues, SkipBlanks:=False, Transpose:=False
    Range(Cells(1, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, 6)).Borders.LineStyle = True
    Cells(2, 7).Select
    Application.DisplayAlerts = False
    Workbooks("Reestr.xlsx").Close False
    Application.DisplayAlerts = True
End If
    For n = 0 To AmPrice
        If MsngArr(n) <> "" Then Msng = Msng & MsngArr(n) & vbNewLine
    Next
    If Msng <> "Ошибки в прайсе(ах):" & vbNewLine Then MsgBox Msng Else MsgBox "Ошибок нет."
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=BookAdress + "\Прайс Питер.xls"
    Application.DisplayAlerts = True
End Sub

Sub prepare(Optional v = 0) 'Подготовка
    Dim EoF As Integer
    
    Workbooks.Open Filename:="C:\Users\Public\Leftovers\Reestr.xlsx"
    Rows("1:2").Delete Shift:=x1Up
    Columns(4).Delete Shift:=x1Left
    Columns(6).Delete Shift:=x1Left
    Columns(2).Cut
    Columns(1).Insert
'Наценка
    Application.Goto Cells(1, 7)
    ActiveCell.FormulaR1C1 = "=c[-3]*1.25"
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 1).End(xlUp).Row, ActiveCell.Column))
    ActiveCell(1, 2).Select
    ActiveCell.FormulaR1C1 = "=Round(c[-1],0)"
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 1).End(xlUp).Row, ActiveCell.Column))
    EoF = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To EoF
        Cells(i, 4) = Cells(i, 8).Value
    Next
    Range(Columns(7), Columns(8)).Clear
End Sub

Sub DeleteDubl(Optional v = 0) 'Удаление дубликатов
    Dim EoF As Integer
    
    Application.Goto Workbooks("Reestr.xlsx").Sheets("TDSheet").Cells(1, 7)
    ActiveCell.FormulaR1C1 = "=VLOOKUP(c1,'[Прайс Питер.xls]Лист1'!C1:C5,5,0)"
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 1).End(xlUp).Row, ActiveCell.Column))
    EoF = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To EoF
        If (Not IsError(Cells(i, 7))) Then
            If (Not IsEmpty(Cells(i, 7))) Then
                Rows(i).Delete Shift:=x1Up
                i = i - 1
            End If
        End If
    Next
    Columns(7).Clear
End Sub

Sub Copy(Optional v = 0) 'Добавление в прайс
    Dim EoF As String
    
    Application.Goto Workbooks("Reestr.xlsx").Sheets("TDSheet").Cells(1, 1)
    EoF = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Rows(1), Rows(EoF)).Copy
End Sub

Function CheckR() As Boolean
    If checkThis("Прайс СД.xls") Then lists(0) = "Прайс СД.xls"
    If checkThis("ZZAP.xls") Then lists(1) = "ZZAP.xls"
    If checkThis("SMDTB.xls") Then lists(2) = "SMDTB.xls"
    If checkThis("Прайс СД.xlsx") Then lists(3) = "Прайс СД.xlsx"
    If checkThis("Прайс Емекс.xls") Then lists(4) = "Прайс Емекс.xls"
    For i = 0 To 4
        If lists(i) <> "" Then list = list + lists(i) + " "
    Next
    If list <> "" Then Note.Show
    For i = 0 To 4
        If lists(i) <> "" Then Call CloseThis(lists(i))
    Next
    check = True
End Function

Private Sub CloseThis(ByVal Book As String)
    If decision Then
        If checkThis(Book) Then Workbooks(Book).Close True
    Else
        If checkThis(Book) Then Workbooks(Book).Close False
    End If
End Sub

Function checkThis(ByVal Book As String) As Boolean
    Dim wb As Workbook
    
    On Error Resume Next                       '//this is VBA way of saying "try"'
    Set wb = Application.Workbooks(Book)
    If Err.Number = 9 Then                     '//this is VBA way of saying "catch"'
        'the file is not opened...'
        checkThis = False
    Else
        checkThis = True
    End If
End Function
