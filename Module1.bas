Attribute VB_Name = "Module1"
Option Explicit



Sub �����X�ֶi����()
Dim shtIdx As Integer
For shtIdx = 1 To Sheets.Count
    Sheets(shtIdx).Activate
    Dim rowCnt, mergeRow As Long
    Dim myrng As Range
    rowCnt = Sheets(shtIdx).UsedRange.Rows.Count 'rowCnt=�C��
    For Each myrng In Range(Cells(2, "A"), Cells(rowCnt, "A"))
            myrng.Select
            mergeRow = myrng.MergeArea.Count
            
            myrng.UnMerge
            myrng.Resize(mergeRow, 1) = myrng
        Next
        
        Sheets(shtIdx).Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
Next
End Sub
