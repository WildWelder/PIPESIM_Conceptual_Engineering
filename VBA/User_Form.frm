VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Item_6 
   Caption         =   "Объекты подготовки"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10350
   OleObjectBlob   =   "User_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Item_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Start_Item_6_Click()
Dim Prediction_Path As String

Prediction_Path = Лист1.Cells(1, 2)
Set FOO_Sheet = ThisWorkbook.Worksheets("ФОО")
Workbooks.Open (Prediction_Path)
Prediction_Name = ActiveWorkbook.Name

If UKPG.Value = True Then Call UKPG_Capacity(FOO_Sheet, Prediction_Name)
If UDK.Value = True Then Call UDK_Capacity(FOO_Sheet, Prediction_Name)
If URM.Value = True Then Call URM_Capacity(FOO_Sheet, Prediction_Name)
If UPN.Value = True Then Call UPN_Capacity(FOO_Sheet, Prediction_Name)
If SKS.Value = True Then Call SKS_Capacity(FOO_Sheet, Prediction_Name)
If TKS.Value = True Then Call TKS_Capacity(FOO_Sheet, Prediction_Name)
If UPPG.Value = True Then Call UPPG_Capacity(FOO_Sheet, Prediction_Name)

Workbooks(Prediction_Name).Close (False)
Unload Item_6
End Sub

Private Sub UKPG_Capacity(FOO_Sheet, Prediction_Name)
Dim GasRate_List() As Variant

Set Prediction_Sheet = Workbooks(Prediction_Name).Worksheets("Газ")
Last_Row = Prediction_Sheet.Cells(1, 1).CurrentRegion.Rows.Count
Last_Column = Prediction_Sheet.Cells(1, 1).CurrentRegion.Columns.Count
GasRate_List = Prediction_Sheet.Range(Prediction_Sheet.Cells(3, 2), Prediction_Sheet.Cells(Last_Row, Last_Column)).Value

Max_Rate = 0
For Row = 1 To UBound(GasRate_List, 1)
    Sum_Rate = 0
    For Column = 1 To UBound(GasRate_List, 2)
        Sum_Rate = Sum_Rate + GasRate_List(Row, Column)
    Next Column
    If Max_Rate < Sum_Rate Then Max_Rate = Sum_Rate
Next Row
Capacity = Max_Rate * SOG.Value / 1000000 * 365 * 0.95
Capacity = Left(Str(Capacity), InStr(1, Str(Capacity), ".") + 2)

Set Found_Cell = FOO_Sheet.Columns(2).Find("Объекты подготовки", LookAt:=1)
Target_Row = Found_Cell.Row
FOO_Sheet.Rows(Target_Row + 1).Insert
FOO_Sheet.Cells(Target_Row + 1, 2) = "Установка комплексной подготовки газа (" & Res_Type.Value & ", " & Technology.Value & ")"
FOO_Sheet.Cells(Target_Row + 1, 3) = Capacity
FOO_Sheet.Cells(Target_Row + 1, 3).NumberFormat = "0.00"
FOO_Sheet.Cells(Target_Row + 1, 4) = "млрд. м3 / год"
FOO_Sheet.Cells(Target_Row + 1, 6) = Capacity
FOO_Sheet.Cells(Target_Row + 1, 6).NumberFormat = "0.00"
With FOO_Sheet.Range(FOO_Sheet.Cells(Target_Row + 1, 1), FOO_Sheet.Cells(Target_Row + 1, Last_Row + 4))
    .Interior.Color = RGB(221, 235, 247)
    .Font.Bold = False
    .Font.Size = 10
    .EntireRow.AutoFit
End With

End Sub

Private Sub UDK_Capacity(FOO_Sheet, Prediction_Name)
Dim CondRate_List() As Variant

Set Prediction_Sheet = Workbooks(Prediction_Name).Worksheets("Нефть")
Last_Row = Prediction_Sheet.Cells(1, 1).CurrentRegion.Rows.Count
Last_Column = Prediction_Sheet.Cells(1, 1).CurrentRegion.Columns.Count
CondRate_List = Prediction_Sheet.Range(Prediction_Sheet.Cells(3, 2), Prediction_Sheet.Cells(Last_Row, Last_Column)).Value

Max_Rate = 0
For Row = 1 To UBound(CondRate_List, 1)
    Sum_Rate = 0
    For Column = 1 To UBound(CondRate_List, 2)
        Sum_Rate = Sum_Rate + CondRate_List(Row, Column)
    Next Column
    If Max_Rate < Sum_Rate Then Max_Rate = Sum_Rate
Next Row
Capacity = Max_Rate * DEK.Value / 1000000 * 365 * 0.95 * DEK_Den.Value
Capacity = Left(Str(Capacity), InStr(1, Str(Capacity), ".") + 2)

Set Found_Cell = FOO_Sheet.Columns(2).Find("Объекты подготовки", LookAt:=1)
Target_Row = Found_Cell.Row
FOO_Sheet.Rows(Target_Row + 1).Insert
FOO_Sheet.Cells(Target_Row + 1, 2) = "Установка деэтанизации конденсата"
FOO_Sheet.Cells(Target_Row + 1, 3) = Capacity
FOO_Sheet.Cells(Target_Row + 1, 3).NumberFormat = "0.00"
FOO_Sheet.Cells(Target_Row + 1, 4) = "млн. т / год"
FOO_Sheet.Cells(Target_Row + 1, 6) = Capacity
FOO_Sheet.Cells(Target_Row + 1, 6).NumberFormat = "0.00"
With FOO_Sheet.Range(FOO_Sheet.Cells(Target_Row + 1, 1), FOO_Sheet.Cells(Target_Row + 1, Last_Row + 4))
    .Interior.Color = RGB(221, 235, 247)
    .Font.Bold = False
    .Font.Size = 10
    .EntireRow.AutoFit
End With

End Sub

Private Sub UPN_Capacity(FOO_Sheet, Prediction_Name)
Dim OilRate_List() As Variant

Set Prediction_Sheet = Workbooks(Prediction_Name).Worksheets("Нефть")
Last_Row = Prediction_Sheet.Cells(1, 1).CurrentRegion.Rows.Count
Last_Column = Prediction_Sheet.Cells(1, 1).CurrentRegion.Columns.Count
OilRate_List = Prediction_Sheet.Range(Prediction_Sheet.Cells(3, 2), Prediction_Sheet.Cells(Last_Row, Last_Column)).Value

Max_Rate = 0
For Row = 1 To UBound(OilRate_List, 1)
    Sum_Rate = 0
    For Column = 1 To UBound(OilRate_List, 2)
        Sum_Rate = Sum_Rate + OilRate_List(Row, Column)
    Next Column
    If Max_Rate < Sum_Rate Then Max_Rate = Sum_Rate
Next Row
Capacity = Max_Rate / 1000000 * 365 * 0.95 * DOD.Value
Capacity = Left(Str(Capacity), InStr(1, Str(Capacity), ".") + 2)

Set Found_Cell = FOO_Sheet.Columns(2).Find("Объекты подготовки", LookAt:=1)
Target_Row = Found_Cell.Row
FOO_Sheet.Rows(Target_Row + 1).Insert
FOO_Sheet.Cells(Target_Row + 1, 2) = "Установка подготовки нефти"
FOO_Sheet.Cells(Target_Row + 1, 3) = Capacity
FOO_Sheet.Cells(Target_Row + 1, 3).NumberFormat = "0.00"
FOO_Sheet.Cells(Target_Row + 1, 4) = "млн. т / год"
FOO_Sheet.Cells(Target_Row + 1, 6) = Capacity
FOO_Sheet.Cells(Target_Row + 1, 6).NumberFormat = "0.00"
With FOO_Sheet.Range(FOO_Sheet.Cells(Target_Row + 1, 1), FOO_Sheet.Cells(Target_Row + 1, Last_Row + 4))
    .Interior.Color = RGB(221, 235, 247)
    .Font.Bold = False
    .Font.Size = 10
    .EntireRow.AutoFit
End With

End Sub

Private Sub URM_Capacity(FOO_Sheet, Prediction_Name)
Dim WaterRate_List() As Variant

Set Prediction_Sheet = Workbooks(Prediction_Name).Worksheets("Вода")
Last_Row = Prediction_Sheet.Cells(1, 1).CurrentRegion.Rows.Count
Last_Column = Prediction_Sheet.Cells(1, 1).CurrentRegion.Columns.Count
WaterRate_List = Prediction_Sheet.Range(Prediction_Sheet.Cells(3, 2), Prediction_Sheet.Cells(Last_Row, Last_Column)).Value

Max_Rate = 0
For Row = 1 To UBound(WaterRate_List, 1)
    Sum_Rate = 0
    For Column = 1 To UBound(WaterRate_List, 2)
        Sum_Rate = Sum_Rate + WaterRate_List(Row, Column)
    Next Column
    If Max_Rate < Sum_Rate Then Max_Rate = Sum_Rate
Next Row
Capacity = Max_Rate * 1.15 / 1000000 * 365 * 0.95
Capacity = Left(Str(Capacity), InStr(1, Str(Capacity), ".") + 2)

Set Found_Cell = FOO_Sheet.Columns(2).Find("Объекты подготовки", LookAt:=1)
Target_Row = Found_Cell.Row
FOO_Sheet.Rows(Target_Row + 1).Insert
FOO_Sheet.Cells(Target_Row + 1, 2) = "Установка регенерации метанола"
FOO_Sheet.Cells(Target_Row + 1, 3) = Capacity
FOO_Sheet.Cells(Target_Row + 1, 3).NumberFormat = "0.00"
FOO_Sheet.Cells(Target_Row + 1, 4) = "млн. т / год"
FOO_Sheet.Cells(Target_Row + 1, 6) = Capacity
FOO_Sheet.Cells(Target_Row + 1, 6).NumberFormat = "0.00"
With FOO_Sheet.Range(FOO_Sheet.Cells(Target_Row + 1, 1), FOO_Sheet.Cells(Target_Row + 1, Last_Row + 4))
    .Interior.Color = RGB(221, 235, 247)
    .Font.Bold = False
    .Font.Size = 10
    .EntireRow.AutoFit
End With

End Sub

Private Sub UPPG_Capacity(FOO_Sheet, Prediction_Name)
Dim GasRate_List() As Variant

Set Prediction_Sheet = Workbooks(Prediction_Name).Worksheets("Газ")
Last_Row = Prediction_Sheet.Cells(1, 1).CurrentRegion.Rows.Count
Last_Column = Prediction_Sheet.Cells(1, 1).CurrentRegion.Columns.Count
GasRate_List = Prediction_Sheet.Range(Prediction_Sheet.Cells(3, 2), Prediction_Sheet.Cells(Last_Row, Last_Column)).Value

Max_Rate = 0
For Row = 1 To UBound(GasRate_List, 1)
    Sum_Rate = 0
    For Column = 1 To UBound(GasRate_List, 2)
        Sum_Rate = Sum_Rate + GasRate_List(Row, Column)
    Next Column
    If Max_Rate < Sum_Rate Then Max_Rate = Sum_Rate
Next Row
Capacity = Max_Rate / 1000000 * 365 * 0.95
Capacity = Left(Str(Capacity), InStr(1, Str(Capacity), ".") + 2)

Set Found_Cell = FOO_Sheet.Columns(2).Find("Объекты подготовки", LookAt:=1)
Target_Row = Found_Cell.Row
FOO_Sheet.Rows(Target_Row + 1).Insert
FOO_Sheet.Cells(Target_Row + 1, 2) = "Установка предварительной подготовки газа"
FOO_Sheet.Cells(Target_Row + 1, 3) = Capacity
FOO_Sheet.Cells(Target_Row + 1, 3).NumberFormat = "0.00"
FOO_Sheet.Cells(Target_Row + 1, 4) = "млрд. м3 / год"
FOO_Sheet.Cells(Target_Row + 1, 6) = Capacity
FOO_Sheet.Cells(Target_Row + 1, 6).NumberFormat = "0.00"
With FOO_Sheet.Range(FOO_Sheet.Cells(Target_Row + 1, 1), FOO_Sheet.Cells(Target_Row + 1, Last_Row + 4))
    .Interior.Color = RGB(221, 235, 247)
    .Font.Bold = False
    .Font.Size = 10
    .EntireRow.AutoFit
End With

End Sub

Private Sub SKS_Capacity(FOO_Sheet, Prediction_Name)

Set Prediction_Sheet = Workbooks(Prediction_Name).Worksheets("Газ")
Last_Row = Prediction_Sheet.Cells(1, 1).CurrentRegion.Rows.Count
Last_Column = Prediction_Sheet.Cells(1, 1).CurrentRegion.Columns.Count
GasRate_List = Prediction_Sheet.Range(Prediction_Sheet.Cells(3, 2), Prediction_Sheet.Cells(Last_Row, Last_Column)).Value

Set Prediction_Sheet = Workbooks(Prediction_Name).Worksheets("Давл.(Сеп.)")
Last_Row = Prediction_Sheet.Cells(1, 1).CurrentRegion.Rows.Count
InPres_List = Prediction_Sheet.Range(Prediction_Sheet.Cells(3, 2), Prediction_Sheet.Cells(Last_Row, 2)).Value
Date_List = Prediction_Sheet.Range(Prediction_Sheet.Cells(3, 1), Prediction_Sheet.Cells(Last_Row, 1)).Value

ReDim SKS_List(Last_Row - 2, 10) As Variant
SKS_List(0, 0) = ""
SKS_List(0, 1) = "Расход газа, млн.м3/сут"
SKS_List(0, 2) = "Давление на входе в КС, Бар(а)"
SKS_List(0, 3) = "Температура на входе в КС, Бар(а)"
SKS_List(0, 4) = "Давление на выходе из 1-ой ступени сепарации, Бар(а)"
SKS_List(0, 5) = "Температура на выходе из 1-ой ступени сепарации, Бар(а)"
SKS_List(0, 6) = "Давление на выходе из 2-ой ступени сепарации, Бар(а)"
SKS_List(0, 7) = "Температура на выходе из 2-ой ступени сепарации, Бар(а)"
SKS_List(0, 8) = "Мощность 1-ой ступени, МВт"
SKS_List(0, 9) = "Мощность 2-ой ступени, МВт"
SKS_List(0, 10) = "Мощность КС, МВт"

OutPres = SKS_Pres.Value + 0.5
KPD = SKS_Target_KPD.Value
Ppk = 4.63 * 10
Tpk = 190.5
InTemp = 283

Max_Capacity = 0
For Row = 1 To UBound(GasRate_List, 1)
    Sum_Rate = 0
    For Column = 1 To UBound(GasRate_List, 2)
        Sum_Rate = Sum_Rate + GasRate_List(Row, Column) / 1000
    Next Column
    InPres = InPres_List(Row, 1) - 2
    SKS_List(Row, 0) = Date_List(Row, 1)
    SKS_List(Row, 1) = Sum_Rate
    SKS_List(Row, 2) = InPres
    SKS_List(Row, 3) = InTemp - 273
    Tpr = InTemp / Tpk
    Ppr = InPres / Ppk
    A1 = -0.39 + 2.03 / Tpr - 3.16 / Tpr ^ 2 + 1.09 / Tpr ^ 3
    A2 = 0.0423 - 0.1812 / Tpr + 0.2124 / Tpr ^ 2
    Z = 1 + A1 * Ppr + A2 * Ppr ^ 2
    OutTemp = InTemp * OutPres / InPres ^ (0.235 / KPD)
    If OutTemp > 423 Then         ' Проверка на необходимость второй ступени
        Temp_AVO = 298
        OutPres1 = (OutPres / InPres) ^ (1 / 2) * InPres
        OutTemp1 = InTemp * (OutPres1 / InPres) ^ (0.235 / KPD)
        OutTemp2 = Temp_AVO * (OutPres / OutPres1) ^ (0.235 / KPD)
        InPres2 = OutPres1 - 0.5
        Tpr_2 = Temp_AVO / Tpk
        Ppr_2 = InPres2 / Ppk
        A1_2 = -0.39 + 2.03 / Tpr_2 - 3.16 / Tpr_2 ^ 2 + 1.09 / Tpr_2 ^ 3
        A2_2 = 0.0423 - 0.1812 / Tpr_2 + 0.2124 / Tpr_2 ^ 2
        Z_2 = 1 + A1_2 * Ppr_2 + A2_2 * Ppr_2 ^ 2
        Capacity1 = 13.34 * Z * InTemp * Sum_Rate / KPD * (((OutPres1 / InPres) ^ 0.3) - 1) / 1000
        Capacity2 = 13.34 * Z_2 * Temp_AVO * Sum_Rate / KPD * (((OutPres / InPres2) ^ 0.3) - 1) / 1000
        SKS_List(Row, 4) = OutPres1
        SKS_List(Row, 5) = OutTemp1 - 273
        SKS_List(Row, 6) = OutPres
        SKS_List(Row, 7) = OutTemp2 - 273
        SKS_List(Row, 8) = Capacity1
        SKS_List(Row, 9) = Capacity2
        SKS_List(Row, 10) = Capacity1 + Capacity2
        GoTo NextRow
    End If
    Capacity1 = 13.34 * Z * InTemp * Sum_Rate / KPD * (((OutPres1 / InPres) ^ 0.3) - 1) / 1000
    SKS_List(Row, 4) = OutPres
    SKS_List(Row, 5) = OutTemp - 273
    SKS_List(Row, 8) = Capacity1
    SKS_List(Row, 10) = Capacity1
    If Max_Capacity < Capacity Then Max_Capacity = Capacity
NextRow:
Next Row
SKS_File_Path = ThisWorkbook.Path & "\Профиль загрузки КС.xlsx"
If Dir(SKS_File_Path) = "" Then
    Set SKS_Book = Workbooks.Add
    SKS_Book.Worksheets("Лист1").Name = "СКС"
    SKS_Book.Worksheets("СКС").Range(Worksheets("СКС").Cells(1, 1), Worksheets("СКС").Cells(Last_Row - 1, 11)).Value = SKS_List
    SKS_Book.SaveAs FileName:=SKS_File_Path
Else:
    Set SKS_Book = Workbooks.Open(SKS_File_Path)
    SKS_Book.Worksheets("Лист2").Name = "СКС"
    SKS_Book.Worksheets("СКС").Range(Worksheets("СКС").Cells(1, 1), Worksheets("СКС").Cells(Last_Row - 1, 11)).Value = SKS_List
    SKS_Book.Save
End If

End Sub

Private Sub TKS_Capacity(FOO_Sheet, Prediction_Name)

Set Prediction_Sheet = Workbooks(Prediction_Name).Worksheets("Газ")
Last_Row = Prediction_Sheet.Cells(1, 1).CurrentRegion.Rows.Count
Last_Column = Prediction_Sheet.Cells(1, 1).CurrentRegion.Columns.Count
GasRate_List = Prediction_Sheet.Range(Prediction_Sheet.Cells(3, 2), Prediction_Sheet.Cells(Last_Row, Last_Column)).Value
Date_List = Prediction_Sheet.Range(Prediction_Sheet.Cells(3, 1), Prediction_Sheet.Cells(Last_Row, 1)).Value

ReDim TKS_List(Last_Row - 2, 10) As Variant
TKS_List(0, 0) = ""
TKS_List(0, 1) = "Расход газа, млн.м3/сут"
TKS_List(0, 2) = "Давление на входе в КС, Бар(а)"
TKS_List(0, 3) = "Температура на входе в КС, Бар(а)"
TKS_List(0, 4) = "Давление на выходе из 1-ой ступени сепарации, Бар(а)"
TKS_List(0, 5) = "Температура на выходе из 1-ой ступени сепарации, Бар(а)"
TKS_List(0, 6) = "Давление на выходе из 2-ой ступени сепарации, Бар(а)"
TKS_List(0, 7) = "Температура на выходе из 2-ой ступени сепарации, Бар(а)"
TKS_List(0, 8) = "Мощность 1-ой ступени, МВт"
TKS_List(0, 9) = "Мощность 2-ой ступени, МВт"
TKS_List(0, 10) = "Мощность КС, МВт"

OutPres = TKS_Pres.Value + 0.5
KPD = TKS_Target_KPD.Value
If Res_Type.Value = "Сеноман" Then          'Аналог - Береговое сеноманский промысел
    Ppk = 4.63 * 10
    Tpk = 190.5
    InTemp = 283
ElseIf Res_Type.Value = "Валанжин" Then         'Аналог - Северо-Уренгойское восточный купол
    Ppk = 4.65 * 10
    Tpk = 205.5
    InTemp = 288
Else                                        'Аналог - Самбургское
    Ppk = 4.59 * 10
    Tpk = 214.5
    InTemp = 290
End If

Max_Capacity = 0
For Row = 1 To UBound(GasRate_List, 1)
    Sum_Rate = 0
    For Column = 1 To UBound(GasRate_List, 2)
        Sum_Rate = Sum_Rate + GasRate_List(Row, Column) / 1000
    Next Column
    InPres = TKS_InPres.Value - 2
    TKS_List(Row, 0) = Date_List(Row, 1)
    TKS_List(Row, 1) = Sum_Rate
    TKS_List(Row, 2) = InPres
    TKS_List(Row, 3) = InTemp - 273
    Tpr = InTemp / Tpk
    Ppr = InPres / Ppk
    A1 = -0.39 + 2.03 / Tpr - 3.16 / Tpr ^ 2 + 1.09 / Tpr ^ 3
    A2 = 0.0423 - 0.1812 / Tpr + 0.2124 / Tpr ^ 2
    Z = 1 + A1 * Ppr + A2 * Ppr ^ 2
    OutTemp = InTemp * (OutPres / InPres) ^ (0.235 / KPD)
    If OutTemp > 423 Then         ' Проверка на необходимость второй ступени
        Temp_AVO = 298
        OutPres1 = (OutPres / InPres) ^ (1 / 2) * InPres
        OutTemp1 = InTemp * (OutPres1 / InPres) ^ (0.235 / KPD)
        OutTemp2 = Temp_AVO * (OutPres / OutPres1) ^ (0.235 / KPD)
        InPres2 = OutPres1 - 0.5
        Tpr_2 = Temp_AVO / Tpk
        Ppr_2 = InPres2 / Ppk
        A1_2 = -0.39 + 2.03 / Tpr_2 - 3.16 / Tpr_2 ^ 2 + 1.09 / Tpr_2 ^ 3
        A2_2 = 0.0423 - 0.1812 / Tpr_2 + 0.2124 / Tpr_2 ^ 2
        Z_2 = 1 + A1_2 * Ppr_2 + A2_2 * Ppr_2 ^ 2
        Capacity1 = 13.34 * Z * InTemp * Sum_Rate / KPD * (((OutPres1 / InPres) ^ 0.3) - 1) / 1000
        Capacity2 = 13.34 * Z_2 * Temp_AVO * Sum_Rate / KPD * (((OutPres / InPres2) ^ 0.3) - 1) / 1000
        TKS_List(Row, 4) = OutPres1
        TKS_List(Row, 5) = OutTemp1 - 273
        TKS_List(Row, 6) = OutPres
        TKS_List(Row, 7) = OutTemp2 - 273
        TKS_List(Row, 8) = Capacity1
        TKS_List(Row, 9) = Capacity2
        TKS_List(Row, 10) = Capacity1 + Capacity2
        GoTo NextRow
    End If
    Capacity1 = 13.34 * Z * InTemp * Sum_Rate / KPD * (((OutPres / InPres) ^ 0.3) - 1) / 1000
    TKS_List(Row, 4) = OutPres
    TKS_List(Row, 5) = OutTemp - 273
    TKS_List(Row, 8) = Capacity1
    TKS_List(Row, 10) = Capacity1
    If Max_Capacity < Capacity Then Max_Capacity = Capacity
NextRow:
Next Row
TKS_File_Path = ThisWorkbook.Path & "\Профиль загрузки КС.xlsx"
If Dir(TKS_File_Path) = "" Then
    Set TKS_Book = Workbooks.Add
    TKS_Book.Worksheets("Лист1").Name = "ТКС"
    TKS_Book.Worksheets("ТКС").Range(Worksheets("ТКС").Cells(1, 1), Worksheets("ТКС").Cells(Last_Row - 1, 11)).Value = TKS_List
    TKS_Book.SaveAs FileName:=TKS_File_Path
Else:
    Set TKS_Book = Workbooks.Open(TKS_File_Path)
    TKS_Book.Worksheets("Лист2").Name = "ТКС"
    TKS_Book.Worksheets("ТКС").Range(Worksheets("ТКС").Cells(1, 1), Worksheets("ТКС").Cells(Last_Row - 1, 11)).Value = TKS_List
    TKS_Book.Save
End If

End Sub

Private Sub Res_Type_Change()
Select Case Res_Type.Value & ", " & Technology.Value
    Case Is = "Валанжин, НТС -30°С"
        SOG.Value = 0.98
        DEK.Value = 1.2
        DEK_Den.Value = 0.7
    Case Is = "Ачимовка, НТС -30°С"
        SOG.Value = 0.96
        DEK.Value = 1.3
        DEK_Den.Value = 0.7
    Case Is = "Валанжин, НТС -60°С"
        SOG.Value = 0.9
        DEK.Value = 1.5
        DEK_Den.Value = 0.65
    Case Is = "Ачимовка, НТС -60°С"
        SOG.Value = 0.88
        DEK.Value = 1.6
        DEK_Den.Value = 0.65
    Case Is = "Сеноман, НТС"
        SOG.Value = 1
        DEK.Value = ""
        DEK_Den.Value = ""
    Case Is = "Сеноман, Адсорбция"
        SOG.Value = 1
        DEK.Value = ""
        DEK_Den.Value = ""
    Case Is = "Сеноман, Абсорбция"
        SOG.Value = 1
        DEK.Value = ""
        DEK_Den.Value = ""
End Select
End Sub
Private Sub Technology_Change()
Select Case Res_Type.Value & ", " & Technology.Value
    Case Is = "Валанжин, НТС -30°С"
        SOG.Value = 0.98
        DEK.Value = 1.2
        DEK_Den.Value = 0.7
    Case Is = "Ачимовка/Юра, НТС -30°С"
        SOG.Value = 0.96
        DEK.Value = 1.3
        DEK_Den.Value = 0.7
    Case Is = "Валанжин, НТС -60°С"
        SOG.Value = 0.9
        DEK.Value = 1.5
        DEK_Den.Value = 0.65
    Case Is = "Ачимовка/Юра, НТС -60°С"
        SOG.Value = 0.88
        DEK.Value = 1.6
        DEK_Den.Value = 0.65
    Case Is = "Сеноман, НТС"
        SOG.Value = 1
        DEK.Value = ""
        DEK_Den.Value = ""
    Case Is = "Сеноман, Адсорбция"
        SOG.Value = 1
        DEK.Value = ""
        DEK_Den.Value = ""
    Case Is = "Сеноман, Абсорбция"
        SOG.Value = 1
        DEK.Value = ""
        DEK_Den.Value = ""
End Select
End Sub

Private Sub UDK_Click()
If UDK.Value = True Then
    DEK_Den.Enabled = True
    DEK_Den.BackStyle = 1
    If UKPG.Value = False Then
        Technology.Enabled = True
        Technology.BackStyle = 1
        SOG.Enabled = True
        SOG.BackStyle = 1
        DEK.Enabled = True
        DEK.BackStyle = 1
    End If
Else:
    DEK_Den.Enabled = False
    DEK_Den.BackStyle = 0
    If UKPG.Value = False Then
        Technology.Enabled = False
        Technology.BackStyle = 0
        SOG.Enabled = False
        SOG.BackStyle = 0
        DEK.Enabled = False
        DEK.BackStyle = 0
    End If
End If
End Sub

Private Sub UKPG_Click()
If UKPG.Value = True Then
    If UDK.Value = False Then
        Technology.Enabled = True
        Technology.BackStyle = 1
        SOG.Enabled = True
        SOG.BackStyle = 1
        DEK.Enabled = True
        DEK.BackStyle = 1
    End If
Else:
    If UDK.Value = False Then
        Technology.Enabled = False
        Technology.BackStyle = 0
        SOG.Enabled = False
        SOG.BackStyle = 0
        DEK.Enabled = False
        DEK.BackStyle = 0
    End If
End If
End Sub

Private Sub UPN_Click()
If UPN.Value = True Then
    DOD.Enabled = True
    DOD.BackStyle = 1
Else:
    DOD.Enabled = False
    DOD.BackStyle = 0
End If
End Sub

Private Sub TKS_Click()
If TKS.Value = True Then
    TKS_Pres.Enabled = True
    TKS_Pres.BackStyle = 1
    TKS_Target_KPD.Enabled = True
    TKS_Target_KPD.BackStyle = 1
    TKS_InPres.Enabled = True
    TKS_InPres.BackStyle = 1
Else:
    TKS_Pres.Enabled = False
    TKS_Pres.BackStyle = 0
    TKS_Target_KPD.Enabled = False
    TKS_Target_KPD.BackStyle = 0
    TKS_InPres.Enabled = False
    TKS_InPres.BackStyle = 0
End If
End Sub

Private Sub SKS_Click()
If SKS.Value = True Then
    SKS_Pres.Enabled = True
    SKS_Pres.BackStyle = 1
    SKS_Target_KPD.Enabled = True
    SKS_Target_KPD.BackStyle = 1
Else:
    SKS_Pres.Enabled = False
    SKS_Pres.BackStyle = 0
    SKS_Target_KPD.Enabled = False
    SKS_Target_KPD.BackStyle = 0
End If
End Sub
