Attribute VB_Name = "Module1"
Option Explicit
Sub ���W�Ƨǭp��()
Attribute ���W�Ƨǭp��.VB_Description = "�Ѥp��j�Ƨ�+�p���`�M,�̤j��,�̤p��\n"
Attribute ���W�Ƨǭp��.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' ���W�Ƨǭp�� ����
' �Ѥp��j�Ƨ�+�p���`�M,�̤j��,�̤p��
'
' �ֳt��: Ctrl+a
'
    Columns("A:B").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B553"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
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
Sub ����Ƨǭp��()
Attribute ����Ƨǭp��.VB_Description = "����Ƨ�+�p���`�M,�̤j��,�̤p��\n"
Attribute ����Ƨǭp��.VB_ProcData.VB_Invoke_Func = "b\n14"
'
' ����Ƨǭp�� ����
' ����Ƨ�+�p���`�M,�̤j��,�̤p��
'
' �ֳt��: Ctrl+b
'
    Columns("A:B").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B553"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
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
