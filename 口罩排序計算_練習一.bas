Attribute VB_Name = "Module1"
Option Explicit

Sub �f�n�Ƨ�()
Attribute �f�n�Ƨ�.VB_Description = "�ƶq�Ѥj��p�Ƨ�"
Attribute �f�n�Ƨ�.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' �f�n�Ƨ� ����
' �ƶq�Ѥj��p�Ƨ�
'
' �ֳt��: Ctrl+q
'
    Columns("A:B").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub �p���`�M()
Attribute �p���`�M.VB_Description = "�p��f�n�ƶq�`�M"
Attribute �p���`�M.VB_ProcData.VB_Invoke_Func = "s\n14"
'
' �p���`�M ����
' �p��f�n�ƶq�`�M
'
' �ֳt��: Ctrl+s
'
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R2C[-3]:R1048576C[-3])"
    Range("G1").Select
End Sub
Sub �p�⥭��()
Attribute �p�⥭��.VB_Description = "�p��f�n�ƶq����"
Attribute �p�⥭��.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' �p�⥭�� ����
' �p��f�n�ƶq����
'
' �ֳt��: Ctrl+w
'
    Range("G2").Select
    ActiveCell.FormulaR1C1 = ""
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R2C[-5]:R1048576C[-5])"
    Range("G2").Select
End Sub
