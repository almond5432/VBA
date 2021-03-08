Attribute VB_Name = "Module1"
Option Explicit

Sub 口罩排序()
Attribute 口罩排序.VB_Description = "數量由大到小排序"
Attribute 口罩排序.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' 口罩排序 巨集
' 數量由大到小排序
'
' 快速鍵: Ctrl+q
'
    Columns("A:B").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub 計算總和()
Attribute 計算總和.VB_Description = "計算口罩數量總和"
Attribute 計算總和.VB_ProcData.VB_Invoke_Func = "s\n14"
'
' 計算總和 巨集
' 計算口罩數量總和
'
' 快速鍵: Ctrl+s
'
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R2C[-3]:R1048576C[-3])"
    Range("G1").Select
End Sub
Sub 計算平均()
Attribute 計算平均.VB_Description = "計算口罩數量平均"
Attribute 計算平均.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' 計算平均 巨集
' 計算口罩數量平均
'
' 快速鍵: Ctrl+w
'
    Range("G2").Select
    ActiveCell.FormulaR1C1 = ""
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R2C[-5]:R1048576C[-5])"
    Range("G2").Select
End Sub
