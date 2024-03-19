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
            If Redundant.checkThis("����� �����.xls") Then Workbooks("����� �����.xls").Close False
            Workbooks.Open Filename:="C:\Users\����� 2\Dropbox\������\������\Clear\����� �����.xls"
            Workbooks(temporaryBook).Close
            Call prepare
            Call FPTA(CatalogAdress, False, True)
        Case 2
            temporaryBook = Workbooks.Add.name
            If checkThis("����� �����.xls") Then Workbooks("����� �����.xls").Close False
            Workbooks.Open Filename:="C:\Users\����� 2\Dropbox\������\������\Clear\����� �����.xls"
            Workbooks(temporaryBook).Close
            Call from1C("����� �����.xls")
            Call check.ND("LeftoversMoscowWithLiart", CatalogAdress)
        Case Else
            End
        End Select
    End If
End Sub

Sub from1C(ByVal BookName As String)
    Dim art, BookAdress As String
    Dim i, EoF, c As Integer
    
'������ � ���������� ����
    Workbooks.Open Filename:="C:\Users\Public\Leftovers\������� �������.xls"
    Rows("1:6").Delete Shift:=xlUp
    Rows(Cells(Rows.Count, 1).End(xlUp).Row).Delete Shift:=x1Up
    Columns(4).NumberFormat = "General"
    Columns(4).Value = Columns(4).Value

'�������� �������
    Application.Goto Workbooks(BookName).Sheets("����1").Cells(2, 7)
    ActiveCell.FormulaR1C1 = "=VLOOKUP(c1,'[������� �������.xls]TDSheet'!C4:C5,2,0)"

'���������� �������
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 1).End(xlUp).Row, ActiveCell.Column))
    ActiveCell(1, 2).Select
    ActiveCell.FormulaR1C1 = "=RC[-3]-RC[-1]"
    Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 1).End(xlUp).Row, ActiveCell.Column))

'�������
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
'��������
    Columns(1).AutoFilter Field:=8, Criteria1:="#�/�"
    Columns(1).AutoFilter Field:=5, Criteria1:="<>0"
End Sub

Sub FPTA(ByVal BookAdress As String, ByVal from1C As Boolean, ByVal fromLiart As Boolean)
    Dim i, n, EoF, AmPrice As Integer
    Dim ListAll, Listadr, Msng, ListBook, SavedName As String
    Dim listName(15), MsngArr(15) As String
    Dim test As Boolean
    
    Application.Goto Workbooks("����� �����.xls").Sheets("����1").Cells(1, 1)
'������ ��������� ���������
    AmPrice = 15
    Listadr = ActiveWorkbook.Path
    listName(0) = "\Li-art ������ ����\����� ��.xls"
    listName(1) = "\ZZap\ZZAP.xls"
    listName(2) = "\����_�����\����� ��.xls"
    listName(6) = "\����_��  ������ ����\����� ��.xls"
    listName(4) = "\�������\����� ��.xls"
    listName(7) = "\����������� ����� ���� �����\����� ��.xls"
    listName(3) = "\������\����� ��.xls"
    listName(5) = "\��������\SMDTB.xls"
    listName(8) = "\�-�����\������ ��.xls"
    listName(9) = "\������_����  ������ ����\����� ��.xlsx"
    listName(10) = "\�����\����� ��.xls"
    listName(11) = "\���������        ������ ����\����� ��.xls"
    listName(12) = "\�����\����� ��.xls"
    listName(13) = "\����� ���� �����\����� ��.xls"
    listName(14) = "\����� �����.xls"
    listName(15) = "\����� ��.xls"
'������������ ������
    Msng = "������ � ������(��):" & vbNewLine
    For n = 0 To AmPrice
        MsngArr(n) = ""
    Next

'�������� ������� � ������
    For n = 0 To AmPrice
    '��������
        ListAll = Listadr + listName(n)
        Workbooks.Open Filename:=ListAll
        ListBook = ActiveWorkbook.name
    If from1C Then
    '�������� �������
        Cells(2, 7).FormulaR1C1 = "=VLOOKUP(c1,'[����� �����.xls]����1'!C1:C5,5,0)"
        Application.Goto Cells(2, 7)
    '���������� �������
        Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 1).End(xlUp).Row, ActiveCell.Column))
        ActiveCell(1, 2).Select
        ActiveCell.FormulaR1C1 = "=RC[-3]-RC[-1]"
        Selection.AutoFill Destination:=Range(ActiveCell, Cells(Cells(Rows.Count, 1).End(xlUp).Row, ActiveCell.Column))
    '�������
        EoF = Cells(Rows.Count, 1).End(xlUp).Row
        For i = 1 To EoF
            If Not IsError(Cells(i, 8)) Then
                If Cells(i, 8) <> 0 Then Cells(i, 5) = Cells(i, 7)
            Else
                test = True
            End If
        Next
    '������
        Range(Columns(7), Columns(8)).Clear
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs Filename:=("C:\Users\����� 2\Dropbox\������\������\Clear" + listName(n))
        Application.DisplayAlerts = True
    End If
    If fromLiart Then
    '���������� �����
        If n = 6 Then Call DeleteDubl
        If 3 < n Then
            Call Copy
            Application.Goto Workbooks(ListBook).Sheets("����1").Cells(1, 1)
            Application.Goto Workbooks(ListBook).Sheets("����1").Cells(Cells(Rows.Count, 1).End(xlUp).Row + 1, 1)
            ActiveCell.PasteSpecial Paste:=xlPasteValues, SkipBlanks:=False, Transpose:=False
            '�������
            Range(Cells(1, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, 6)).Borders.LineStyle = True
        End If
    End If
    '��������
        Cells(2, 7).Select
        Application.DisplayAlerts = False
        SavedName = BookAdress + listName(n)
        ActiveWorkbook.SaveAs Filename:=(SavedName)
        ActiveWorkbook.Close False
        Application.DisplayAlerts = True
    '������ ������
        If test Then MsngArr(n) = listName(n)
        test = False
    Next
    
'�������� �����
    Application.DisplayAlerts = False
    Application.Goto Workbooks("����� �����.xls").Sheets("����1").Cells(1, 1)
    ActiveWorkbook.SaveAs Filename:="C:\Users\����� 2\Dropbox\������\������\Clear\����� �����.xls"
    Application.DisplayAlerts = True
If fromLiart Then
    Call Copy
    Application.Goto Workbooks("����� �����.xls").Sheets("����1").Cells(1, 1)
    Application.Goto Workbooks("����� �����.xls").Sheets("����1").Cells(Cells(Rows.Count, 1).End(xlUp).Row + 1, 1)
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
    If Msng <> "������ � ������(��):" & vbNewLine Then MsgBox Msng Else MsgBox "������ ���."
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=BookAdress + "\����� �����.xls"
    Application.DisplayAlerts = True
End Sub

Sub prepare(Optional v = 0) '����������
    Dim EoF As Integer
    
    Workbooks.Open Filename:="C:\Users\Public\Leftovers\Reestr.xlsx"
    Rows("1:2").Delete Shift:=x1Up
    Columns(4).Delete Shift:=x1Left
    Columns(6).Delete Shift:=x1Left
    Columns(2).Cut
    Columns(1).Insert
'�������
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

Sub DeleteDubl(Optional v = 0) '�������� ����������
    Dim EoF As Integer
    
    Application.Goto Workbooks("Reestr.xlsx").Sheets("TDSheet").Cells(1, 7)
    ActiveCell.FormulaR1C1 = "=VLOOKUP(c1,'[����� �����.xls]����1'!C1:C5,5,0)"
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

Sub Copy(Optional v = 0) '���������� � �����
    Dim EoF As String
    
    Application.Goto Workbooks("Reestr.xlsx").Sheets("TDSheet").Cells(1, 1)
    EoF = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Rows(1), Rows(EoF)).Copy
End Sub

Function CheckR() As Boolean
    If checkThis("����� ��.xls") Then lists(0) = "����� ��.xls"
    If checkThis("ZZAP.xls") Then lists(1) = "ZZAP.xls"
    If checkThis("SMDTB.xls") Then lists(2) = "SMDTB.xls"
    If checkThis("����� ��.xlsx") Then lists(3) = "����� ��.xlsx"
    If checkThis("����� �����.xls") Then lists(4) = "����� �����.xls"
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
