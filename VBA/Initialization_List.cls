VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "����3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub OilRate_Click()

If OilRate Then

    CGR.Value = False
    If ����3.OilRate.Caption = "Oil Rate" Then ����5.Cells(1, 1).Value = "����� �����, �3/���" Else ����5.Cells(1, 1).Value = "����� ����������, �3/���"
    
End If

End Sub

Sub CGR_Click()

If CGR Then

    OilRate.Value = False
    ����5.Cells(1, 1).Value = "���, �3/�3"
    
End If

End Sub

Sub GasRate_Click()

If GasRate Then

    GOR.Value = False
    ����1.Cells(1, 1).Value = "����� ����, ���.�3/���"
    
End If

End Sub

Sub GOR_Click()

If GOR Then

    GasRate.Value = False
    ����1.Cells(1, 1).Value = "���, �3/�3"
    
End If

End Sub

Sub WaterRate_Click()

If WaterRate Then

    WGR.Value = False
    WCT.Value = False
    ����6.Cells(1, 1).Value = "����� ����, �3/���"
    
End If

End Sub

Sub WGR_Click()

If WGR Then

    WaterRate.Value = False
    WCT.Value = False
    ����6.Cells(1, 1).Value = "���, �3/�3"
    
End If

End Sub

Sub WCT_Click()

If WCT Then

    WaterRate.Value = False
    WGR.Value = False
    ����6.Cells(1, 1).Value = "�������������, %"
    
End If

End Sub

Sub Choke_Change_Click()

If Choke_Change.Value = True Then
    
    ����13.Visible = xlSheetVisible
    Choke_Change.Caption = "ON"
    ����4.Cells(1, 4).Value = "������"
    
End If

If Choke_Change.Value = False Then
    ����13.Range("B1:OO1000").Clear
    ����13.Range("A3:A1000").Clear
    ����13.Visible = xlSheetHidden

    Choke_Change.Caption = "OFF"
    ����4.Cells(1, 4).Value = "����"
    
End If

End Sub

Sub Init_Vel_Gas_Check_Click()

If ����3.CheckBoxes("Init_Vel_Gas_Check").Value = 1 Then
    ����8.Visible = xlSheetVisible
End If

If ����3.CheckBoxes("Init_Vel_Gas_Check").Value = -4146 Then
    ����8.Range("B1:OO1000").Clear
    ����8.Range("A3:A1000").Clear
    ����8.Visible = xlSheetHidden
    
End If

End Sub

Sub Final_Vel_Gas_Check_Click()

If ����3.CheckBoxes("Final_Vel_Gas_Check").Value = 1 Then
    ����9.Visible = xlSheetVisible
End If

If ����3.CheckBoxes("Final_Vel_Gas_Check").Value = -4146 Then
    ����9.Range("B1:OO1000").Clear
    ����9.Range("A3:A1000").Clear
    ����9.Visible = xlSheetHidden
End If

End Sub

Sub Init_Vel_Liq_Check_Click()

If ����3.CheckBoxes("Init_Vel_Liq_Check").Value = 1 Then
    ����15.Visible = xlSheetVisible
End If

If ����3.CheckBoxes("Init_Vel_Liq_Check").Value = -4146 Then
    ����15.Range("B1:OO1000").Clear
    ����15.Range("A3:A1000").Clear
    ����15.Visible = xlSheetHidden
    
End If

End Sub

Sub Final_Vel_Liq_Check_Click()

If ����3.CheckBoxes("Final_Vel_Liq_Check").Value = 1 Then
    ����16.Visible = xlSheetVisible
End If

If ����3.CheckBoxes("Final_Vel_Liq_Check").Value = -4146 Then
    ����16.Range("B1:OO1000").Clear
    ����16.Range("A3:A1000").Clear
    ����16.Visible = xlSheetHidden
End If

End Sub

Sub Spec_DP_Check_Click()

If ����3.CheckBoxes("Spec_DP_Check").Value = 1 Then
    ����10.Visible = xlSheetVisible
End If

If ����3.CheckBoxes("Spec_DP_Check").Value = -4146 Then
    ����10.Range("B1:OO1000").Clear
    ����10.Range("A3:A1000").Clear
    ����10.Visible = xlSheetHidden
End If

End Sub

Sub EVR_Check_Click()

If ����3.CheckBoxes("EVR_Check").Value = 1 Then
    ����12.Visible = xlSheetVisible
End If

If ����3.CheckBoxes("EVR_Check").Value = -4146 Then
    ����12.Range("B1:OO1000").Clear
    ����12.Range("A3:A1000").Clear
    ����12.Visible = xlSheetHidden
End If

End Sub

Sub THP_Check_Click()

If ����3.CheckBoxes("THP_Check").Value = 1 Then
    ����11.Visible = xlSheetVisible
End If

If ����3.CheckBoxes("THP_Check").Value = -4146 Then
    ����11.Range("B1:OO1000").Clear
    ����11.Range("A3:A1000").Clear
    ����11.Visible = xlSheetHidden
End If

End Sub
