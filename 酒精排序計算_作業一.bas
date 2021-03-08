Attribute VB_Name = "Module1"
Option Explicit
Sub 遞增排序計算()
Attribute 遞增排序計算.VB_Description = "由小到大排序+計算總和,最大值,最小值\n"
Attribute 遞增排序計算.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' 遞增排序計算 巨集
' 由小到大排序+計算總和,最大值,最小值
'
' 快速鍵: Ctrl+a
'
    Columns("A:B").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B553"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B553")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R2C[-3]:R1048576C[-3])"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=MAX(R2C[-5]:R1048576C[-5])"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "=MIN(R2C[-7]:R1048576C[-7])"
    Range("I2").Select
End Sub
Sub 遞減排序計算()
Attribute 遞減排序計算.VB_Description = "遞減排序+計算總和,最大值,最小值\n"
Attribute 遞減排序計算.VB_ProcData.VB_Invoke_Func = "b\n14"
'
' 遞減排序計算 巨集
' 遞減排序+計算總和,最大值,最小值
'
' 快速鍵: Ctrl+b
'
    Columns("A:B").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B553"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B553")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R2C[-3]:R1048576C[-3])"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=MAX(R2C[-5]:R1048576C[-5])"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "=MIN(R2C[-7]:R1048576C[-7])"
    Range("I2").Select
End Sub
