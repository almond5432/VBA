Attribute VB_Name = "Module1"
Option Explicit

Sub �ʺA�g��()
Dim MachineName As String
MachineName = InputBox("�п�J�]�ƦW��")
Cells(2, 1).Value = MachineName

Dim modeName As String
modeName = InputBox("�п�J�Ҩ�W��")
Cells(2, 2).Value = modeName

Dim unitPrice As Integer
unitPrice = InputBox("�п�J���")
Cells(2, 3).Value = CInt(unitPrice)

Dim qty As Integer
qty = InputBox("�п�J�ƶq")
Cells(2, 4).Value = CInt(qty)

Dim total As Integer
total = unitPrice * qty
Cells(2, 5) = total

End Sub
