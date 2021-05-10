Attribute VB_Name = "Module1"
Option Explicit



Sub 取消合併進階版()
Dim shtIdx As Integer
For shtIdx = 1 To Sheets.Count
    Sheets(shtIdx).Activate
    Dim rowCnt, mergeRow As Long
    Dim myrng As Range
    rowCnt = Sheets(shtIdx).UsedRange.Rows.Count 'rowCnt=列數
    For Each myrng In Range(Cells(2, "A"), Cells(rowCnt, "A"))
            myrng.Select
            mergeRow = myrng.MergeArea.Count
            
            myrng.UnMerge
            myrng.Resize(mergeRow, 1) = myrng
        Next
        
        Sheets(shtIdx).Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
Next
End Sub
