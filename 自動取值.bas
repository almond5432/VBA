Attribute VB_Name = "Module1"
Option Explicit

Sub �ʺA����()
Dim deviceName As String
deviceName = Cells(2, 1).Value
MsgBox "�]�ƦW��:" & deviceName

Dim modeName As String
modeName = Cells(2, 2).Value
MsgBox "�Ҩ�W��:" & modeName

Dim unitPrice As Integer
unitPrice = Cells(2, 3).Value
MsgBox "���:" & unitPrice

Dim qty As Integer
qty = Cells(2, 4).Value
MsgBox "�ƶq:" & qty

Dim total As Integer
total = qty * unitPrice
MsgBox "�`��:" & total

End Sub
