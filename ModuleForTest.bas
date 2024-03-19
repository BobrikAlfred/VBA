Attribute VB_Name = "ModuleForTest"
Option Explicit

Sub test4()
    Dim i As Integer, n As Integer, k As Integer, list() As String, quantity As Integer, numers As Integer
    quantity = 1
    numers = 1
    n = 1
    k = 1
    ReDim list(1736, 1)
    
    For i = ActiveCell.Row To Cells(Rows.Count, 1).End(xlUp).Row
            If Cells(i - 1, 2) = Cells(i, 2) Then
                quantity = quantity + 1
            Else
                Cells(i, 7) = numers
                Do While n < quantity
                    Cells(i - quantity + 1, 7) = Cells(i - quantity, 7)
                    quantity = quantity - 1
                Loop
                numers = numers + 1
            End If
    Next
End Sub

Sub test3()
    Dim i As Integer
    Dim ans As String
    
    For i = ActiveCell.Row To Cells(Rows.Count, 1).End(xlUp).Row
        If Cells(i, 6) = 0 Then
        Cells(i, 2).Copy
        Cells(i, 7).Select
        ans = InputBox(Cells(i, 2))
        Select Case ans
        Case ""
            Cells(i, 7) = "нет"
        Case "с"
            Cells(i, 7) = "средне"
            Cells(i, 8) = InputBox(Cells(i, 2))
        Case "о"
            Cells(i, 7) = "отлично"
            Cells(i, 8) = InputBox(Cells(i, 2))
        Case "end"
            Exit For
        Case Else
            Cells(i, 7) = "пусто"
            Cells(i, 8) = ans
        End Select
        End If
    Next
End Sub

Public Sub mailing()
    Dim sTo As String, sAttachment As String, sSubject As String
    sTo = Workbooks("PERSONAL.XLSB").Sheets("Самара").Cells(2, 6)
    sSubject = Workbooks("PERSONAL.XLSB").Sheets("Самара").Cells(2, 8)
    sAttachment = ActiveWorkbook.Path + "\"
    sAttachment = sAttachment + Workbooks("PERSONAL.XLSB").Sheets("Самара").Cells(2, 7)
'    If Not IsBookOpen("SLC4.xls") Then Workbooks.Open Filename:="C:\Users\ДАКАР 2\Dropbox\самара\Прайсы с остатками\Октябрь\SLC4.xls"
    Call SendMail( _
        "", _
        "samara@smart-d.ru", _
        "p.e.tr.o.vich@mail.ru," + sTo, _
        sSubject, _
        "", _
        sAttachment _
    )
    
End Sub
