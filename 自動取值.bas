Attribute VB_Name = "Module1"
Option Explicit

Sub 動態取值()
Dim deviceName As String
deviceName = Cells(2, 1).Value
MsgBox "設備名稱:" & deviceName

Dim modeName As String
modeName = Cells(2, 2).Value
MsgBox "模具名稱:" & modeName

Dim unitPrice As Integer
unitPrice = Cells(2, 3).Value
MsgBox "單價:" & unitPrice

Dim qty As Integer
qty = Cells(2, 4).Value
MsgBox "數量:" & qty

Dim total As Integer
total = qty * unitPrice
MsgBox "總價:" & total

End Sub
