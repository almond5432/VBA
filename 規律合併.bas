Attribute VB_Name = "Module1"
Option Explicit

Sub �W�ߦX��()
Dim shtsIdx As Integer
For shtsIdx = 1 To Sheets.Count
    Dim rangeStr As String
    rangeStr = "A" & k & ":A" & k + 2
    MsgBox "�ثe�X�ֽd��" & rangeStr
    Range(rangeStr).Merge
    Next
    
End Sub


Sub �u�@��W�ߦX��()
Dim shtsIdx As Integer
Dim k As Integer
For shtsIdx = 1 To Sheets.Count
Sheets(shtsIdx).Activate
For k = 2 To 11 Step 3

    Dim rangeStr As String
    rangeStr = "A" & k & ":A" & k + 2

    Range(rangeStr).Merge
    Next
    Next
End Sub
