Attribute VB_Name = "MyLibrary"
Option Explicit

Public Sub Вставить_значение()
Attribute Вставить_значение.VB_ProcData.VB_Invoke_Func = "V\n14"
'
' ВставитьЗ Макрос
'
' Сочетание клавиш: Ctrl+Shift+V
'
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Public Sub Вставить_ВПР_из_Прайс_Питер()
'
' Сочетание клавиш: Ctrl+Shift+E (англ.)
'
    ActiveCell.FormulaR1C1 = "=VLOOKUP(C[-6],'[Прайс питер.xls]Лист1'!C1:C5,5,0)"
End Sub

Public Sub Окрасить_повторы()
    Dim i As Integer, k As Integer
    Dim cell As Range
     
    Dim Dupes()     'объявляем массив для хранения дубликатов
    ReDim Dupes(1 To Selection.Cells.Count, 1 To 2)
     
    Selection.Interior.ColorIndex = -4142   'убираем заливку если была
    i = 3
    For Each cell In Selection
        If WorksheetFunction.CountIf(Selection, cell.Value) > 1 Then
            For k = LBound(Dupes) To UBound(Dupes)
                'если ячейка уже есть в массиве дубликатов - заливаем
                If Dupes(k, 1) = cell Then cell.Interior.ColorIndex = Dupes(k, 2)
            Next k
            'если ячейка содержит дубликат, но еще не в массиве - добавляем ее в массив и заливаем
            If cell.Interior.ColorIndex = -4142 Then
                cell.Interior.ColorIndex = i
                Dupes(i, 1) = cell.Value
                Dupes(i, 2) = i
                i = i + 1
            End If
        End If
    Next cell
End Sub

Public Sub AsNumbers()
    Selection.NumberFormat = "General"
    Selection.Value = Selection.Value
End Sub

Public Sub adress()
    Dim BookName, art As String
    Dim i, EoF, c As Integer
    
    BookName = ActiveWorkbook.name
    Workbooks.Open Filename:="C:\Users\ДАКАР 2\Desktop\Пересчёт товара СМР.xlsx"
    Workbooks(BookName).Activate
    Select Case Cells(3, 7)
    Case "Примечание"
        ActiveCell.FormulaR1C1 = "=VLOOKUP(c2,'[Пересчёт товара СМР.xlsx]Лист1'!R3C4:R467C5,2,0)"
    Case ""
        ActiveCell.FormulaR1C1 = "=VLOOKUP(c6,'[Пересчёт товара СМР.xlsx]Лист1'!R3C4:R467C5,2,0)"
    Case "SLC4"
        ActiveCell.FormulaR1C1 = "=VLOOKUP(c3,'[Пересчёт товара СМР.xlsx]Лист1'!R3C4:R467C5,2,0)"
    End Select
    Workbooks("Пересчёт товара СМР.xlsx").Close False
End Sub

Sub writeDown()
    Dim EoF As Integer
    Dim art As String, brend As String, name As String, price As Integer, quantity As Byte, thisRow As Integer
    
    name = ActiveWorkbook.name
    If Not IsBookOpen("Ручкин продажа.xlsx") Then Workbooks.Open Filename:="C:\Users\ДАКАР 2\Dropbox\самара\Ручкин продажа.xlsx"
    Workbooks(name).Activate
    thisRow = ActiveCell.Row
    Select Case Cells(3, 7)
    Case "Примечание"
        art = Cells(thisRow, 2)
        brend = Cells(thisRow, 1)
        name = Cells(thisRow, 3)
        price = Cells(thisRow, 5)
        quantity = Cells(thisRow, 4)
    Case ""
        art = Cells(thisRow, 6)
        brend = Cells(thisRow, 8)
        name = Cells(thisRow, 4)
        price = Cells(thisRow, 11)
        quantity = Cells(thisRow, 10)
    Case "SLC4"
        art = Cells(thisRow, 3)
        brend = Cells(thisRow, 2)
        name = Cells(thisRow, 3)
        price = Cells(thisRow, 11)
        quantity = Cells(thisRow, 10)
    Case Else
        Exit Sub
    End Select
    EoF = Workbooks("Ручкин продажа.xlsx").Sheets("Лист1").Cells(Rows.Count, 1).End(xlUp).Row + 1
    With Workbooks("Ручкин продажа.xlsx").Sheets("Лист1")
        .Cells(EoF, 1) = art
        .Cells(EoF, 2) = brend
        .Cells(EoF, 3) = name
        .Cells(EoF, 4) = price
        .Cells(EoF, 5) = quantity
        .Cells(EoF, 6) = Date
    End With
    ActiveCell = "Добавлено"
End Sub

Sub SendMail_Ruchkin()
    Select Case ActiveWorkbook.name
    Case "Ручкин продажа.xlsx"
        If 6 = MsgBox("Отправить письмо?", vbYesNo, vbApplicationModal) Then _
            Call SendMail_Default("buh@smart-d.ru", "Ручкин товар на приход")
    End Select
End Sub

Public Function IsBookOpen(ByVal Book As String) As Boolean
    Dim wb As Workbook
    
    On Error Resume Next                       '//this is VBA way of saying "try"'
    Set wb = Application.Workbooks(Book)
    If Err.Number = 9 Then                     '//this is VBA way of saying "catch"'
        'the file is not opened...'
        IsBookOpen = False
    Else
        IsBookOpen = True
    End If
End Function

Public Sub SendMail_Default(ByVal sTo As String, Optional ByVal sSubject As String)
    ActiveWorkbook.SendMail sTo, sSubject
End Sub

Public Sub SendMail( _
    ByVal sTo As String, _
    ByVal sCC As String, _
    ByVal sBCC As String, _
    ByVal sSubject As String, _
    ByVal sBody As String, _
    Optional ByVal sAttachment As String = "")
'====================================================================================================================
'    sTo = Кому
'    sCC = Копия
'    sBCC = Скрытая копия
'    sSubject = Тема письма
'    sBody = Текст письма
'    sAttachment = Вложение(полный путь к файлу)
'====================================================================================================================
    Const CDO_Cnf = "http://schemas.microsoft.com/cdo/configuration/"
    Dim oCDOCnf As Object, oCDOMsg As Object
    Dim SMTPserver As String, sUsername As String, sPass As String, sMsg As String
    Dim sFrom As String
    On Error Resume Next
    'sFrom – как правило совпадает с sUsername
    SMTPserver = "smtp.yandex.com"    ' SMTPServer: для Mail.ru "smtp.mail.ru"; для Яндекса "smtp.yandex.ru"; для Рамблера "mail.rambler.ru"
    sUsername = "Samara@smart-d.ru"    ' Учетная запись на сервере
    sPass = "R2wovJ38i0"    ' Пароль к почтовому аккаунту
    sFrom = "Samara@smart-d.ru"    'От кого
    
    If Len(SMTPserver) = 0 Then MsgBox "Не указан SMTP сервер", vbInformation, "www.Excel-VBA.ru": Exit Sub
    If Len(sUsername) = 0 Then MsgBox "Не указана учетная запись", vbInformation, "www.Excel-VBA.ru": Exit Sub
    If Len(sPass) = 0 Then MsgBox "Не указан пароль", vbInformation, "www.Excel-VBA.ru": Exit Sub

    'Назначаем конфигурацию CDO
    Set oCDOCnf = CreateObject("CDO.Configuration")
    With oCDOCnf.Fields
        .Item(CDO_Cnf & "sendusing") = 2
        .Item(CDO_Cnf & "smtpauthenticate") = 1
        .Item(CDO_Cnf & "smtpserver") = SMTPserver
        'если необходимо указать SSL
        .Item(CDO_Cnf & "smtpserverport") = 465 'для Яндекса и Gmail 465
        .Item(CDO_Cnf & "smtpusessl") = True
        '=====================================
        .Item(CDO_Cnf & "sendusername") = sUsername
        .Item(CDO_Cnf & "sendpassword") = sPass
        .Update
    End With
    'Создаем сообщение
    Set oCDOMsg = CreateObject("CDO.Message")
    With oCDOMsg
        Set .Configuration = oCDOCnf
        .BodyPart.Charset = "koi8-r"
        .From = sFrom
        .To = sTo
        .CC = sCC
        .BCC = sBCC
        .Subject = sSubject
        .TextBody = sBody
        'Проверка наличия файла по указанному пути
        If Len(sAttachment) > 0 Then
            If Dir(sAttachment, 16) <> "" Then
                .AddAttachment sAttachment
            End If
        End If
        .send
    End With
 
    Select Case Err.Number
    Case -2147220973: sMsg = "Нет доступа к Интернет"
    Case -2147220975: sMsg = "Отказ сервера SMTP"
    Case 0: sMsg = "Письмо отправлено"
    Case Else: sMsg = "Ошибка номер: " & Err.Number & vbNewLine & "Описание ошибки: " & Err.Description
    End Select
    MsgBox sMsg, vbInformation, "www.Excel-VBA.ru"
    Set oCDOMsg = Nothing: Set oCDOCnf = Nothing
End Sub
