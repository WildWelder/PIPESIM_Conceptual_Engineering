VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub OilRate_Click()

If OilRate Then

    CGR.Value = False
    If Лист3.OilRate.Caption = "Oil Rate" Then Лист5.Cells(1, 1).Value = "Дебит нефти, м3/сут" Else Лист5.Cells(1, 1).Value = "Дебит конденсата, м3/сут"
    
End If

End Sub

Sub CGR_Click()

If CGR Then

    OilRate.Value = False
    Лист5.Cells(1, 1).Value = "КГФ, м3/м3"
    
End If

End Sub

Sub GasRate_Click()

If GasRate Then

    GOR.Value = False
    Лист1.Cells(1, 1).Value = "Дебит газа, тыс.м3/сут"
    
End If

End Sub

Sub GOR_Click()

If GOR Then

    GasRate.Value = False
    Лист1.Cells(1, 1).Value = "ГНФ, м3/м3"
    
End If

End Sub

Sub WaterRate_Click()

If WaterRate Then

    WGR.Value = False
    WCT.Value = False
    Лист6.Cells(1, 1).Value = "Дебит воды, м3/сут"
    
End If

End Sub

Sub WGR_Click()

If WGR Then

    WaterRate.Value = False
    WCT.Value = False
    Лист6.Cells(1, 1).Value = "ВГФ, м3/м3"
    
End If

End Sub

Sub WCT_Click()

If WCT Then

    WaterRate.Value = False
    WGR.Value = False
    Лист6.Cells(1, 1).Value = "Обводненность, %"
    
End If

End Sub

Sub Choke_Change_Click()

If Choke_Change.Value = True Then
    
    Лист13.Visible = xlSheetVisible
    Choke_Change.Caption = "ON"
    Лист4.Cells(1, 4).Value = "ИСТИНА"
    
End If

If Choke_Change.Value = False Then
    Лист13.Range("B1:OO1000").Clear
    Лист13.Range("A3:A1000").Clear
    Лист13.Visible = xlSheetHidden

    Choke_Change.Caption = "OFF"
    Лист4.Cells(1, 4).Value = "ЛОЖЬ"
    
End If

End Sub

Sub Init_Vel_Gas_Check_Click()

If Лист3.CheckBoxes("Init_Vel_Gas_Check").Value = 1 Then
    Лист8.Visible = xlSheetVisible
End If

If Лист3.CheckBoxes("Init_Vel_Gas_Check").Value = -4146 Then
    Лист8.Range("B1:OO1000").Clear
    Лист8.Range("A3:A1000").Clear
    Лист8.Visible = xlSheetHidden
    
End If

End Sub

Sub Final_Vel_Gas_Check_Click()

If Лист3.CheckBoxes("Final_Vel_Gas_Check").Value = 1 Then
    Лист9.Visible = xlSheetVisible
End If

If Лист3.CheckBoxes("Final_Vel_Gas_Check").Value = -4146 Then
    Лист9.Range("B1:OO1000").Clear
    Лист9.Range("A3:A1000").Clear
    Лист9.Visible = xlSheetHidden
End If

End Sub

Sub Init_Vel_Liq_Check_Click()

If Лист3.CheckBoxes("Init_Vel_Liq_Check").Value = 1 Then
    Лист15.Visible = xlSheetVisible
End If

If Лист3.CheckBoxes("Init_Vel_Liq_Check").Value = -4146 Then
    Лист15.Range("B1:OO1000").Clear
    Лист15.Range("A3:A1000").Clear
    Лист15.Visible = xlSheetHidden
    
End If

End Sub

Sub Final_Vel_Liq_Check_Click()

If Лист3.CheckBoxes("Final_Vel_Liq_Check").Value = 1 Then
    Лист16.Visible = xlSheetVisible
End If

If Лист3.CheckBoxes("Final_Vel_Liq_Check").Value = -4146 Then
    Лист16.Range("B1:OO1000").Clear
    Лист16.Range("A3:A1000").Clear
    Лист16.Visible = xlSheetHidden
End If

End Sub

Sub Spec_DP_Check_Click()

If Лист3.CheckBoxes("Spec_DP_Check").Value = 1 Then
    Лист10.Visible = xlSheetVisible
End If

If Лист3.CheckBoxes("Spec_DP_Check").Value = -4146 Then
    Лист10.Range("B1:OO1000").Clear
    Лист10.Range("A3:A1000").Clear
    Лист10.Visible = xlSheetHidden
End If

End Sub

Sub EVR_Check_Click()

If Лист3.CheckBoxes("EVR_Check").Value = 1 Then
    Лист12.Visible = xlSheetVisible
End If

If Лист3.CheckBoxes("EVR_Check").Value = -4146 Then
    Лист12.Range("B1:OO1000").Clear
    Лист12.Range("A3:A1000").Clear
    Лист12.Visible = xlSheetHidden
End If

End Sub

Sub THP_Check_Click()

If Лист3.CheckBoxes("THP_Check").Value = 1 Then
    Лист11.Visible = xlSheetVisible
End If

If Лист3.CheckBoxes("THP_Check").Value = -4146 Then
    Лист11.Range("B1:OO1000").Clear
    Лист11.Range("A3:A1000").Clear
    Лист11.Visible = xlSheetHidden
End If

End Sub
