Attribute VB_Name = "Module1"
Option Explicit

Sub ¬d¸ß()
Dim qman As String
Dim rowNum As Integer
Dim content As String
Dim payStatus As Boolean
qman = Range("G1").Value

For rowNum = 2 To 7
If (Cells(rowNum, "A").Value = qman) Then
    Range("G2").Value = Cells(rowNum, 2).Value
    
    If (Cells(rowNum, "C").Value = "Y") Then
    payStatus = True
    content = qman & "©µ¿ðª¬ºA" & payStatus
    MsgBox (content)
    Else
    payStatus = False
    End If
    
    Else
    
    End If
    Next
    
End Sub
