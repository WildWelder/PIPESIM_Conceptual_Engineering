Attribute VB_Name = "IVG"
'Option Explicit

Sub Get_Dir_Module()      '�������� ���� ���������� ������ ���������� ��������

Dim fd As FileDialog
Dim result As Integer
Dim FileName, FilePath, dirModel, pnsxFile, sumFile, gapFile As String
Dim position, AllPos As Long
With ActiveWorkbook.Sheets("���")
    'get directory Gap model file
    
    Set fd = Application.FileDialog(msoFileDialogOpen)
    
    fd.Filters.Clear
    
    With fd
        .Title = "�������� ���� � ����������� ��������� (Excel)"
        .Filters.Add "Excel file", "*.xlsx", 1
        .Filters.Add "Excel file", "*.xlsm", 1
        .AllowMultiSelect = False
        .InitialFileName = dirModel
    End With
        
    result = fd.Show

End With

Range("B1").Value = Trim(fd.SelectedItems.Item(1))
    
End Sub

Sub Add_Item_1()
Item_1.Show
End Sub

Sub Add_Item_2()
Item_2.Show
End Sub

Sub Add_Item_6()
With Item_6.Res_Type
    .AddItem "��������"
    .AddItem "��������/���"
    .AddItem "�������"
End With
Item_6.Res_Type.Value = "�������"
With Item_6.Technology
    .AddItem "��� -30��"
    .AddItem "��� -60��"
    .AddItem "���������"
    .AddItem "���������"
End With
Item_6.Show
End Sub

Sub DKS_Capacity()
Row = 3
For Row = 3 To 112
    T = ����3.Cells(Row, 3) + 273
    P = ����3.Cells(Row, 2)
    Ppk = 4.578252201
    Tpk = 216.690595
    Ppr = P / Ppk
    Tpr = T / Tpk
    A1 = -0.39 + 2.03 / Tpr - 3.16 / Tpr ^ 2 + 1.09 / Tpr ^ 3
    A2 = 0.0423 - 0.1812 / Tpr + 0.2124 / Tpr ^ 2
    Z = 1 + A1 * Ppr + A2 * Ppr ^ 2
    ����3.Cells(Row, 5) = Z
    
Next Row
End Sub

Sub test()
Dim Pred_Mod_Dir As String, Pos As String, FilePath As String, FileName As String, Source_Name As String, First_Cell As String
Dim Row As Integer
Dim Source_List As Variant
Dim GetValue As Object
    Row = 3
    Source_List = []
    Source_Name = "1"
    First_Cell = "J3"
    Do Until Source_Name = ""
        Pred_Mod_Dir = ����1.Cells(1, 2).Value
        Pos = InStrRev(Pred_Mod_Dir, "\")
        FilePath = Left(Pred_Mod_Dir, Pos)
        FileName = Right(Pred_Mod_Dir, Len(Pred_Mod_Dir) - Pos)
        Arg = "'" & FilePath & "[" & FileName & "]" & Sheet & "'!" & _
           Range(ref).Range("A1").Address(, , xlR1C1)
        Source_Name = GetValue(FilePath, FileName, "���� 3", First_Cell)
        Source_List = Source_List + Source_Name
        Row = Row + 1
    Loop
    ����1.Cells(11, 1) = Source_List
        
End Sub

Sub test111()
Tg = 283
Pg = 7.5

End Sub
