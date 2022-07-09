Attribute VB_Name = "Module1"
Sub Gas_Click()

����4.Cells(10, 4) = "���"
����3.Shapes("Group 27").Select
    With Selection.ShapeRange.Glow
        .Color.ObjectThemeColor = msoThemeColorAccent1
        .Color.TintAndShade = 0
        .Color.Brightness = 0
        .Transparency = 0.6000000238
        .Radius = 18
    End With
    
����3.Shapes("Group 30").Select
    With Selection.ShapeRange.Glow
        .Radius = 0
    End With
    
����3.GOR.Enabled = False
����3.GOR.Caption = "GOR"
����3.GOR.Value = False
����3.WCT.Enabled = False
����3.WCT.Caption = "Watercut"
����3.WCT.Value = False

����3.CGR.Enabled = True
����3.WGR.Enabled = True
����3.OilRate.Caption = "Condensate Rate"
����3.Cells(5, 2).Value = "������ ������� ������ ����������:"
    With ����3.Cells(5, 2).Characters(Start:=16, Length:=18).Font
        .Name = "Calibri"
        .FontStyle = "����������"
    End With

����3.GasRate.Value = False
����3.OilRate.Value = False
����3.CGR.Value = False
����3.WaterRate.Value = False
����3.WGR.Value = False

Cells(1, 2).Select


'With ����11.Cells(2, 1).Validation
 '   .Delete
  '  .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
   ' xlBetween, Formula1:="=Source_Var_Gas"
    '.ShowInput = True
    '.ShowError = True
'End With

'����11.Cells(2, 1) = "������ ����� ��������"

End Sub

Sub Oil_Click()

����4.Cells(10, 4) = "�����"
����3.Shapes("Group 30").Select
    With Selection.ShapeRange.Glow
        .Color.ObjectThemeColor = msoThemeColorAccent2
        .Color.TintAndShade = 0
        .Color.Brightness = 0
        .Transparency = 0.6000000238
        .Radius = 18
    End With
    
����3.Shapes("Group 27").Select
    With Selection.ShapeRange.Glow
        .Radius = 0
    End With

����3.CGR.Enabled = False
����3.CGR.Caption = "CGR"
����3.CGR.Value = False
����3.WGR.Enabled = False
����3.WGR.Caption = "WGR"
����3.WGR.Value = False

����3.GOR.Enabled = True
����3.WCT.Enabled = True
����3.OilRate.Caption = "Oil Rate"
����3.Cells(5, 2).Value = "������ ������� ������ �����:"
    With ����3.Cells(5, 2).Characters(Start:=16, Length:=12).Font
        .Name = "Calibri"
        .FontStyle = "����������"
    End With

����3.GasRate.Value = False
����3.GOR.Value = False
����3.OilRate.Value = False
����3.WaterRate.Value = False
����3.WCT.Value = False

Cells(1, 2).Select

'With ����11.Cells(2, 1).Validation
'    .Delete
 '   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'    xlBetween, Formula1:="=Source_Var_Oil"
'    .ShowInput = True
'    .ShowError = True
'End With

'����11.Cells(2, 1) = "������ ����� ��������"

End Sub

Sub Get_Dir_Model()      '�������� ���� ���������� ������ PIPESIM

Dim fd As FileDialog
Dim result As Integer
Dim fileName, FilePath, dirModel, pnsxFile, sumFile, gapFile As String
Dim position, AllPos As Long
With ActiveWorkbook.Sheets("�������.")
    'get directory Gap model file
    
    Set fd = Application.FileDialog(msoFileDialogOpen)
    
    fd.Filters.Clear
    
    With fd
        .Title = "Select PIPESIM model file"
        .Filters.Add "PIPESIM model file", "*.pips", 1
        .AllowMultiSelect = False
        .InitialFileName = dirModel
    End With
        
    result = fd.Show

End With

Range("A21").Value = Trim(fd.SelectedItems.Item(1))
    
End Sub

Sub Transfer_Dates_List()

Dim Date_Num As Integer

Date_Num = 3

Do Until IsEmpty(����3.Cells(Date_Num, 8).Value)

    Date_Value = ����3.Cells(Date_Num, 8).Value

    ����1.Cells(Date_Num, 1).Value = Date_Value
    ����1.Cells(Date_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
    ����1.Cells(Date_Num, 1).Interior.TintAndShade = -0.499984740745262
    ����1.Cells(Date_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����1.Cells(Date_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    ����1.Cells(Date_Num, 1).Font.Bold = True
    ����1.Cells(Date_Num, 1).Font.ThemeColor = xlThemeColorDark1
    ����1.Cells(Date_Num, 1).HorizontalAlignment = xlCenter
    ����1.Cells(Date_Num, 1).VerticalAlignment = xlCenter
    
    ����5.Cells(Date_Num, 1).Value = Date_Value
    ����5.Cells(Date_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
    ����5.Cells(Date_Num, 1).Interior.TintAndShade = -0.499984740745262
    ����5.Cells(Date_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����5.Cells(Date_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    ����5.Cells(Date_Num, 1).Font.Bold = True
    ����5.Cells(Date_Num, 1).Font.ThemeColor = xlThemeColorDark1
    ����5.Cells(Date_Num, 1).HorizontalAlignment = xlCenter
    ����5.Cells(Date_Num, 1).VerticalAlignment = xlCenter
    
    ����6.Cells(Date_Num, 1).Value = Date_Value
    ����6.Cells(Date_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
    ����6.Cells(Date_Num, 1).Interior.TintAndShade = -0.499984740745262
    ����6.Cells(Date_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����6.Cells(Date_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    ����6.Cells(Date_Num, 1).Font.Bold = True
    ����6.Cells(Date_Num, 1).Font.ThemeColor = xlThemeColorDark1
    ����6.Cells(Date_Num, 1).HorizontalAlignment = xlCenter
    ����6.Cells(Date_Num, 1).VerticalAlignment = xlCenter
    
    ����2.Cells(Date_Num, 1).Value = Date_Value
    ����2.Cells(Date_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
    ����2.Cells(Date_Num, 1).Interior.TintAndShade = -0.499984740745262
    ����2.Cells(Date_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����2.Cells(Date_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    ����2.Cells(Date_Num, 1).Font.Bold = True
    ����2.Cells(Date_Num, 1).Font.ThemeColor = xlThemeColorDark1
    ����2.Cells(Date_Num, 1).HorizontalAlignment = xlCenter
    ����2.Cells(Date_Num, 1).VerticalAlignment = xlCenter
    
    ����14.Cells(Date_Num, 1).Value = Date_Value
    ����14.Cells(Date_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
    ����14.Cells(Date_Num, 1).Interior.TintAndShade = -0.499984740745262
    ����14.Cells(Date_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����14.Cells(Date_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    ����14.Cells(Date_Num, 1).Font.Bold = True
    ����14.Cells(Date_Num, 1).Font.ThemeColor = xlThemeColorDark1
    ����14.Cells(Date_Num, 1).HorizontalAlignment = xlCenter
    ����14.Cells(Date_Num, 1).VerticalAlignment = xlCenter
    
    ����7.Cells(Date_Num, 1).Value = Date_Value
    ����7.Cells(Date_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
    ����7.Cells(Date_Num, 1).Interior.TintAndShade = -0.499984740745262
    ����7.Cells(Date_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����7.Cells(Date_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    ����7.Cells(Date_Num, 1).Font.Bold = True
    ����7.Cells(Date_Num, 1).Font.ThemeColor = xlThemeColorDark1
    ����7.Cells(Date_Num, 1).HorizontalAlignment = xlCenter
    ����7.Cells(Date_Num, 1).VerticalAlignment = xlCenter

    If ����3.Choke_Change.Value = True Then
        ����13.Cells(Date_Num, 1).Value = Date_Value
        ����13.Cells(Date_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
        ����13.Cells(Date_Num, 1).Interior.TintAndShade = -0.499984740745262
        ����13.Cells(Date_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
        ����13.Cells(Date_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        ����13.Cells(Date_Num, 1).Font.Bold = True
        ����13.Cells(Date_Num, 1).Font.ThemeColor = xlThemeColorDark1
        ����13.Cells(Date_Num, 1).HorizontalAlignment = xlCenter
        ����13.Cells(Date_Num, 1).VerticalAlignment = xlCenter
    End If
    If ����3.CheckBoxes("Init_Vel_Gas_Check").Value = 1 Then
        ����8.Cells(2, Date_Num - 1).Value = Date_Value
        ����8.Cells(2, Date_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
        ����8.Cells(2, Date_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
        ����8.Cells(2, Date_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        ����8.Cells(2, Date_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
        ����8.Cells(2, Date_Num - 1).Font.Bold = True
        ����8.Cells(2, Date_Num - 1).Font.ThemeColor = xlThemeColorDark1
        ����8.Cells(2, Date_Num - 1).HorizontalAlignment = xlCenter
        ����8.Cells(2, Date_Num - 1).VerticalAlignment = xlCenter
    End If
    If ����3.CheckBoxes("Init_Vel_Liq_Check").Value = 1 Then
        ����15.Cells(2, Date_Num - 1).Value = Date_Value
        ����15.Cells(2, Date_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
        ����15.Cells(2, Date_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
        ����15.Cells(2, Date_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        ����15.Cells(2, Date_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
        ����15.Cells(2, Date_Num - 1).Font.Bold = True
        ����15.Cells(2, Date_Num - 1).Font.ThemeColor = xlThemeColorDark1
        ����15.Cells(2, Date_Num - 1).HorizontalAlignment = xlCenter
        ����15.Cells(2, Date_Num - 1).VerticalAlignment = xlCenter
    End If
    If ����3.CheckBoxes("Init_Vel_Liq_Check").Value = 1 Then
        ����16.Cells(2, Date_Num - 1).Value = Date_Value
        ����16.Cells(2, Date_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
        ����16.Cells(2, Date_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
        ����16.Cells(2, Date_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        ����16.Cells(2, Date_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
        ����16.Cells(2, Date_Num - 1).Font.Bold = True
        ����16.Cells(2, Date_Num - 1).Font.ThemeColor = xlThemeColorDark1
        ����16.Cells(2, Date_Num - 1).HorizontalAlignment = xlCenter
        ����16.Cells(2, Date_Num - 1).VerticalAlignment = xlCenter
    End If
    If ����3.CheckBoxes("Final_Vel_Gas_Check").Value = 1 Then
        ����9.Cells(2, Date_Num - 1).Value = Date_Value
        ����9.Cells(2, Date_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
        ����9.Cells(2, Date_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
        ����9.Cells(2, Date_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        ����9.Cells(2, Date_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
        ����9.Cells(2, Date_Num - 1).Font.Bold = True
        ����9.Cells(2, Date_Num - 1).Font.ThemeColor = xlThemeColorDark1
        ����9.Cells(2, Date_Num - 1).HorizontalAlignment = xlCenter
        ����9.Cells(2, Date_Num - 1).VerticalAlignment = xlCenter
    End If
    If ����3.CheckBoxes("Spec_DP_Check").Value = 1 Then
        ����10.Cells(2, Date_Num - 1).Value = Date_Value
        ����10.Cells(2, Date_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
        ����10.Cells(2, Date_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
        ����10.Cells(2, Date_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        ����10.Cells(2, Date_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
        ����10.Cells(2, Date_Num - 1).Font.Bold = True
        ����10.Cells(2, Date_Num - 1).Font.ThemeColor = xlThemeColorDark1
        ����10.Cells(2, Date_Num - 1).HorizontalAlignment = xlCenter
        ����10.Cells(2, Date_Num - 1).VerticalAlignment = xlCenter
    End If
    If ����3.CheckBoxes("THP_Check").Value = 1 Then
        ����11.Cells(Date_Num, 1).Value = Date_Value
        ����11.Cells(Date_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
        ����11.Cells(Date_Num, 1).Interior.TintAndShade = -0.499984740745262
        ����11.Cells(Date_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
        ����11.Cells(Date_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        ����11.Cells(Date_Num, 1).Font.Bold = True
        ����11.Cells(Date_Num, 1).Font.ThemeColor = xlThemeColorDark1
        ����11.Cells(Date_Num, 1).HorizontalAlignment = xlCenter
        ����11.Cells(Date_Num, 1).VerticalAlignment = xlCenter
    End If
    If ����3.CheckBoxes("EVR_Check").Value = 1 Then
        ����12.Cells(2, Date_Num - 1).Value = Date_Value
        ����12.Cells(2, Date_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
        ����12.Cells(2, Date_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
        ����12.Cells(2, Date_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        ����12.Cells(2, Date_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
        ����12.Cells(2, Date_Num - 1).Font.Bold = True
        ����12.Cells(2, Date_Num - 1).Font.ThemeColor = xlThemeColorDark1
        ����12.Cells(2, Date_Num - 1).HorizontalAlignment = xlCenter
        ����12.Cells(2, Date_Num - 1).VerticalAlignment = xlCenter
    End If

    Date_Num = Date_Num + 1
    
Loop

End Sub

Sub Transfer_Pipes_List()

Dim Pipe_Num As Integer, Pipe_Value As String

Pipe_Num = 3

Do Until IsEmpty(����3.Cells(Pipe_Num, 9).Value)

    Pipe_Value = ����3.Cells(Pipe_Num, 9).Value
    
    ����17.Cells(Pipe_Num, 1).Value = Pipe_Value
    ����17.Cells(Pipe_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
    ����17.Cells(Pipe_Num, 1).Interior.TintAndShade = -0.499984740745262
    ����17.Cells(Pipe_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����17.Cells(Pipe_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    ����17.Cells(Pipe_Num, 1).Font.Bold = True
    ����17.Cells(Pipe_Num, 1).Font.ThemeColor = xlThemeColorDark1
    ����17.Cells(Pipe_Num, 1).HorizontalAlignment = xlCenter
    ����17.Cells(Pipe_Num, 1).VerticalAlignment = xlCenter
    ����17.Cells(Pipe_Num, 2).Style = "20% � ������3"
    ����17.Cells(Pipe_Num, 2).HorizontalAlignment = xlCenter
    ����17.Cells(Pipe_Num, 2).VerticalAlignment = xlCenter
    ����17.Cells(Pipe_Num, 2).borders(xlEdgeBottom).LineStyle = xlContinuous
    ����17.Cells(Pipe_Num, 2).borders(xlEdgeRight).LineStyle = xlContinuous
    
    If ����3.CheckBoxes("Init_Vel_Gas_Check").Value = 1 Then
        ����8.Cells(Pipe_Num, 1).Value = Pipe_Value
        ����8.Cells(Pipe_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
        ����8.Cells(Pipe_Num, 1).Interior.TintAndShade = -0.499984740745262
        ����8.Cells(Pipe_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
        ����8.Cells(Pipe_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        ����8.Cells(Pipe_Num, 1).Font.Bold = True
        ����8.Cells(Pipe_Num, 1).Font.ThemeColor = xlThemeColorDark1
        ����8.Cells(Pipe_Num, 1).HorizontalAlignment = xlCenter
        ����8.Cells(Pipe_Num, 1).VerticalAlignment = xlCenter
    End If
    If ����3.CheckBoxes("Final_Vel_Gas_Check").Value = 1 Then
        ����9.Cells(Pipe_Num, 1).Value = Pipe_Value
        ����9.Cells(Pipe_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
        ����9.Cells(Pipe_Num, 1).Interior.TintAndShade = -0.499984740745262
        ����9.Cells(Pipe_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
        ����9.Cells(Pipe_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        ����9.Cells(Pipe_Num, 1).Font.Bold = True
        ����9.Cells(Pipe_Num, 1).Font.ThemeColor = xlThemeColorDark1
        ����9.Cells(Pipe_Num, 1).HorizontalAlignment = xlCenter
        ����9.Cells(Pipe_Num, 1).VerticalAlignment = xlCenter
    End If
    If ����3.CheckBoxes("Init_Vel_Liq_Check").Value = 1 Then
        ����15.Cells(Pipe_Num, 1).Value = Pipe_Value
        ����15.Cells(Pipe_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
        ����15.Cells(Pipe_Num, 1).Interior.TintAndShade = -0.499984740745262
        ����15.Cells(Pipe_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
        ����15.Cells(Pipe_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        ����15.Cells(Pipe_Num, 1).Font.Bold = True
        ����15.Cells(Pipe_Num, 1).Font.ThemeColor = xlThemeColorDark1
        ����15.Cells(Pipe_Num, 1).HorizontalAlignment = xlCenter
        ����15.Cells(Pipe_Num, 1).VerticalAlignment = xlCenter
    End If
    If ����3.CheckBoxes("Final_Vel_Liq_Check").Value = 1 Then
        ����16.Cells(Pipe_Num, 1).Value = Pipe_Value
        ����16.Cells(Pipe_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
        ����16.Cells(Pipe_Num, 1).Interior.TintAndShade = -0.499984740745262
        ����16.Cells(Pipe_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
        ����16.Cells(Pipe_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        ����16.Cells(Pipe_Num, 1).Font.Bold = True
        ����16.Cells(Pipe_Num, 1).Font.ThemeColor = xlThemeColorDark1
        ����16.Cells(Pipe_Num, 1).HorizontalAlignment = xlCenter
        ����16.Cells(Pipe_Num, 1).VerticalAlignment = xlCenter
    End If
    If ����3.CheckBoxes("Spec_DP_Check").Value = 1 Then
        ����10.Cells(Pipe_Num, 1).Value = Pipe_Value
        ����10.Cells(Pipe_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
        ����10.Cells(Pipe_Num, 1).Interior.TintAndShade = -0.499984740745262
        ����10.Cells(Pipe_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
        ����10.Cells(Pipe_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        ����10.Cells(Pipe_Num, 1).Font.Bold = True
        ����10.Cells(Pipe_Num, 1).Font.ThemeColor = xlThemeColorDark1
        ����10.Cells(Pipe_Num, 1).HorizontalAlignment = xlCenter
        ����10.Cells(Pipe_Num, 1).VerticalAlignment = xlCenter
    End If
    If ����3.CheckBoxes("EVR_Check").Value = 1 Then
        ����12.Cells(Pipe_Num, 1).Value = Pipe_Value
        ����12.Cells(Pipe_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
        ����12.Cells(Pipe_Num, 1).Interior.TintAndShade = -0.499984740745262
        ����12.Cells(Pipe_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
        ����12.Cells(Pipe_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        ����12.Cells(Pipe_Num, 1).Font.Bold = True
        ����12.Cells(Pipe_Num, 1).Font.ThemeColor = xlThemeColorDark1
        ����12.Cells(Pipe_Num, 1).HorizontalAlignment = xlCenter
        ����12.Cells(Pipe_Num, 1).VerticalAlignment = xlCenter
    End If

    Pipe_Num = Pipe_Num + 1
    
Loop

End Sub

Sub Transfer_Source_List()

Dim Source_Num As Integer, Source_Value As String

Source_Num = 3

Do Until IsEmpty(����3.Cells(Source_Num, 10).Value)

    Source_Value = ����3.Cells(Source_Num, 10).Value
    
    If ����3.Cells(Source_Num, 11).Value = "" Then GoTo NoneWP
    ����1.Cells(1, Source_Num - 1) = ����3.Range(Left(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
    ����1.Cells(1, Source_Num - 1).Style = "20% � ������1"
    ����1.Cells(1, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����1.Cells(1, Source_Num - 1).Font.Bold = True
    ����1.Cells(1, Source_Num - 1).HorizontalAlignment = xlCenter
    ����1.Cells(1, Source_Num - 1).VerticalAlignment = xlCenter
          
    ����5.Cells(1, Source_Num - 1).Value = ����3.Range(Left(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
    ����5.Cells(1, Source_Num - 1).Style = "20% � ������1"
    ����5.Cells(1, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����5.Cells(1, Source_Num - 1).Font.Bold = True
    ����5.Cells(1, Source_Num - 1).HorizontalAlignment = xlCenter
    ����5.Cells(1, Source_Num - 1).VerticalAlignment = xlCenter
        
    ����6.Cells(1, Source_Num - 1).Value = ����3.Range(Left(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
    ����6.Cells(1, Source_Num - 1).Style = "20% � ������1"
    ����6.Cells(1, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����6.Cells(1, Source_Num - 1).Font.Bold = True
    ����6.Cells(1, Source_Num - 1).HorizontalAlignment = xlCenter
    ����6.Cells(1, Source_Num - 1).VerticalAlignment = xlCenter
        
    ����2.Cells(1, Source_Num - 1).Value = ����3.Range(Left(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
    ����2.Cells(1, Source_Num - 1).Style = "20% � ������1"
    ����2.Cells(1, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����2.Cells(1, Source_Num - 1).Font.Bold = True
    ����2.Cells(1, Source_Num - 1).HorizontalAlignment = xlCenter
    ����2.Cells(1, Source_Num - 1).VerticalAlignment = xlCenter
        
    ����14.Cells(1, Source_Num - 1).Value = ����3.Range(Left(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
    ����14.Cells(1, Source_Num - 1).Style = "20% � ������1"
    ����14.Cells(1, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����14.Cells(1, Source_Num - 1).Font.Bold = True
    ����14.Cells(1, Source_Num - 1).HorizontalAlignment = xlCenter
    ����14.Cells(1, Source_Num - 1).VerticalAlignment = xlCenter
NoneWP:

    ����1.Cells(2, Source_Num - 1).Value = Source_Value
    ����1.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
    ����1.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    ����1.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����1.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
    ����1.Cells(2, Source_Num - 1).Font.Bold = True
    ����1.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
    ����1.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
    ����1.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter
      
    ����5.Cells(2, Source_Num - 1).Value = Source_Value
    ����5.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
    ����5.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    ����5.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����5.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
    ����5.Cells(2, Source_Num - 1).Font.Bold = True
    ����5.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
    ����5.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
    ����5.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter
    
    ����6.Cells(2, Source_Num - 1).Value = Source_Value
    ����6.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
    ����6.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    ����6.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����6.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
    ����6.Cells(2, Source_Num - 1).Font.Bold = True
    ����6.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
    ����6.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
    ����6.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter
    
    ����2.Cells(2, Source_Num - 1).Value = Source_Value
    ����2.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
    ����2.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    ����2.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����2.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
    ����2.Cells(2, Source_Num - 1).Font.Bold = True
    ����2.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
    ����2.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
    ����2.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter
    
    ����14.Cells(2, Source_Num - 1).Value = Source_Value
    ����14.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
    ����14.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    ����14.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����14.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
    ����14.Cells(2, Source_Num - 1).Font.Bold = True
    ����14.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
    ����14.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
    ����14.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter
    
    If ����3.CheckBoxes("THP_Check").Value = 1 Then
        ����11.Cells(2, Source_Num - 1).Value = Source_Value
        ����11.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
        ����11.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        ����11.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
        ����11.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
        ����11.Cells(2, Source_Num - 1).Font.Bold = True
        ����11.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
        ����11.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
        ����11.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter
    End If
    
    Source_Num = Source_Num + 1
    
Loop

End Sub

Sub Transfer_Source_Filter_List()

Dim Source_Num As Integer, Source_Num_g_g As Integer, Source_Num_g_o As Integer, Source_Num_o_o As Integer, Source_Num_o_g As Integer
Dim Source_Num_w_w As Integer, Source_Num_w_g As Integer, Source_Num_w_o As Integer
Dim Source_Value As String

Source_Num = 3
Source_Num_g_g = 3
Source_Num_g_o = 3
Source_Num_o_o = 3
Source_Num_o_g = 3
Source_Num_w_w = 3
Source_Num_w_g = 3
Source_Num_w_o = 3

Do Until IsEmpty(����3.Cells(Source_Num, 10).Value)
    Source_Value = ����3.Cells(Source_Num, 10).Value
    
    If ����3.GasRate.Value = True Then
        If Right(Source_Value, 2) = "_g" Then
            ����1.Cells(1, Source_Num_g_g - 1).Value = ����3.Range(Left(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
            ����1.Cells(1, Source_Num_g_g - 1).Style = "20% � ������1"
            ����1.Cells(1, Source_Num_g_g - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            ����1.Cells(1, Source_Num_g_g - 1).Font.Bold = True
            ����1.Cells(1, Source_Num_g_g - 1).HorizontalAlignment = xlCenter
            ����1.Cells(1, Source_Num_g_g - 1).VerticalAlignment = xlCenter
            
            ����1.Cells(2, Source_Num_g_g - 1).Value = Source_Value
            ����1.Cells(2, Source_Num_g_g - 1).Interior.ThemeColor = xlThemeColorAccent1
            ����1.Cells(2, Source_Num_g_g - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
            ����1.Cells(2, Source_Num_g_g - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            ����1.Cells(2, Source_Num_g_g - 1).borders(xlEdgeTop).LineStyle = xlContinuous
            ����1.Cells(2, Source_Num_g_g - 1).Font.Bold = True
            ����1.Cells(2, Source_Num_g_g - 1).Font.ThemeColor = xlThemeColorDark1
            ����1.Cells(2, Source_Num_g_g - 1).HorizontalAlignment = xlCenter
            ����1.Cells(2, Source_Num_g_g - 1).VerticalAlignment = xlCenter
            Source_Num_g_g = Source_Num_g_g + 1
        End If
      Else
        If Right(Source_Value, 2) = "_o" Then
            ����1.Cells(1, Source_Num_g_o - 1).Value = ����3.Range(Left(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
            ����1.Cells(1, Source_Num_g_o - 1).Style = "20% � ������1"
            ����1.Cells(1, Source_Num_g_o - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            ����1.Cells(1, Source_Num_g_o - 1).Font.Bold = True
            ����1.Cells(1, Source_Num_g_o - 1).HorizontalAlignment = xlCenter
            ����1.Cells(1, Source_Num_g_o - 1).VerticalAlignment = xlCenter
  
            ����1.Cells(2, Source_Num_g_o - 1).Value = Source_Value
            ����1.Cells(2, Source_Num_g_o - 1).Interior.ThemeColor = xlThemeColorAccent1
            ����1.Cells(2, Source_Num_g_o - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
            ����1.Cells(2, Source_Num_g_o - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            ����1.Cells(2, Source_Num_g_o - 1).borders(xlEdgeTop).LineStyle = xlContinuous
            ����1.Cells(2, Source_Num_g_o - 1).Font.Bold = True
            ����1.Cells(2, Source_Num_g_o - 1).Font.ThemeColor = xlThemeColorDark1
            ����1.Cells(2, Source_Num_g_o - 1).HorizontalAlignment = xlCenter
            ����1.Cells(2, Source_Num_g_o - 1).VerticalAlignment = xlCenter
            Source_Num_g_o = Source_Num_g_o + 1
        End If
    End If
          
    If ����3.OilRate.Value = True Then
        If Right(Source_Value, 2) = "_o" Or Right(Source_Value, 2) = "_c" Then
            ����5.Cells(1, Source_Num_o_o - 1).Value = ����3.Range(Left(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
            ����5.Cells(1, Source_Num_o_o - 1).Style = "20% � ������1"
            ����5.Cells(1, Source_Num_o_o - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            ����5.Cells(1, Source_Num_o_o - 1).Font.Bold = True
            ����5.Cells(1, Source_Num_o_o - 1).HorizontalAlignment = xlCenter
            ����5.Cells(1, Source_Num_o_o - 1).VerticalAlignment = xlCenter
            
            ����5.Cells(2, Source_Num_o_o - 1).Value = Source_Value
            ����5.Cells(2, Source_Num_o_o - 1).Interior.ThemeColor = xlThemeColorAccent1
            ����5.Cells(2, Source_Num_o_o - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
            ����5.Cells(2, Source_Num_o_o - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            ����5.Cells(2, Source_Num_o_o - 1).borders(xlEdgeTop).LineStyle = xlContinuous
            ����5.Cells(2, Source_Num_o_o - 1).Font.Bold = True
            ����5.Cells(2, Source_Num_o_o - 1).Font.ThemeColor = xlThemeColorDark1
            ����5.Cells(2, Source_Num_o_o - 1).HorizontalAlignment = xlCenter
            ����5.Cells(2, Source_Num_o_o - 1).VerticalAlignment = xlCenter
            Source_Num_o_o = Source_Num_o_o + 1
        End If
    Else
        If Right(Source_Value, 2) = "_g" Then
            ����5.Cells(1, Source_Num_o_g - 1).Value = ����3.Range(Left(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
            ����5.Cells(1, Source_Num_o_g - 1).Style = "20% � ������1"
            ����5.Cells(1, Source_Num_o_g - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            ����5.Cells(1, Source_Num_o_g - 1).Font.Bold = True
            ����5.Cells(1, Source_Num_o_g - 1).HorizontalAlignment = xlCenter
            ����5.Cells(1, Source_Num_o_g - 1).VerticalAlignment = xlCenter
            
            ����5.Cells(2, Source_Num_o_g - 1).Value = Source_Value
            ����5.Cells(2, Source_Num_o_g - 1).Interior.ThemeColor = xlThemeColorAccent1
            ����5.Cells(2, Source_Num_o_g - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
            ����5.Cells(2, Source_Num_o_g - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            ����5.Cells(2, Source_Num_o_g - 1).borders(xlEdgeTop).LineStyle = xlContinuous
            ����5.Cells(2, Source_Num_o_g - 1).Font.Bold = True
            ����5.Cells(2, Source_Num_o_g - 1).Font.ThemeColor = xlThemeColorDark1
            ����5.Cells(2, Source_Num_o_g - 1).HorizontalAlignment = xlCenter
            ����5.Cells(2, Source_Num_o_g - 1).VerticalAlignment = xlCenter
            Source_Num_o_g = Source_Num_o_g + 1
        End If
    End If
    
    If ����3.WaterRate.Value = True Then
        If Right(Source_Value, 2) = "_w" Then
            ����6.Cells(1, Source_Num_w_w - 1).Value = ����3.Range(Left(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
            ����6.Cells(1, Source_Num_w_w - 1).Style = "20% � ������1"
            ����6.Cells(1, Source_Num_w_w - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            ����6.Cells(1, Source_Num_w_w - 1).Font.Bold = True
            ����6.Cells(1, Source_Num_w_w - 1).HorizontalAlignment = xlCenter
            ����6.Cells(1, Source_Num_w_w - 1).VerticalAlignment = xlCenter
            
            ����6.Cells(2, Source_Num_w_w - 1).Value = Source_Value
            ����6.Cells(2, Source_Num_w_w - 1).Interior.ThemeColor = xlThemeColorAccent1
            ����6.Cells(2, Source_Num_w_w - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
            ����6.Cells(2, Source_Num_w_w - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            ����6.Cells(2, Source_Num_w_w - 1).borders(xlEdgeTop).LineStyle = xlContinuous
            ����6.Cells(2, Source_Num_w_w - 1).Font.Bold = True
            ����6.Cells(2, Source_Num_w_w - 1).Font.ThemeColor = xlThemeColorDark1
            ����6.Cells(2, Source_Num_w_w - 1).HorizontalAlignment = xlCenter
            ����6.Cells(2, Source_Num_w_w - 1).VerticalAlignment = xlCenter
            Source_Num_w_w = Source_Num_w_w + 1
        End If
    ElseIf ����3.WGR.Value = True Then
        If Right(Source_Value, 2) = "_g" Then
            ����6.Cells(1, Source_Num_w_g - 1).Value = ����3.Range(Left(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
            ����6.Cells(1, Source_Num_w_g - 1).Style = "20% � ������1"
            ����6.Cells(1, Source_Num_w_g - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            ����6.Cells(1, Source_Num_w_g - 1).Font.Bold = True
            ����6.Cells(1, Source_Num_w_g - 1).HorizontalAlignment = xlCenter
            ����6.Cells(1, Source_Num_w_g - 1).VerticalAlignment = xlCenter
            
            ����6.Cells(2, Source_Num_w_g - 1).Value = Source_Value
            ����6.Cells(2, Source_Num_w_g - 1).Interior.ThemeColor = xlThemeColorAccent1
            ����6.Cells(2, Source_Num_w_g - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
            ����6.Cells(2, Source_Num_w_g - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            ����6.Cells(2, Source_Num_w_g - 1).borders(xlEdgeTop).LineStyle = xlContinuous
            ����6.Cells(2, Source_Num_w_g - 1).Font.Bold = True
            ����6.Cells(2, Source_Num_w_g - 1).Font.ThemeColor = xlThemeColorDark1
            ����6.Cells(2, Source_Num_w_g - 1).HorizontalAlignment = xlCenter
            ����6.Cells(2, Source_Num_w_g - 1).VerticalAlignment = xlCenter
            Source_Num_w_g = Source_Num_w_g + 1
    Else
        If Right(Source_Value, 2) = "_o" Then
            ����6.Cells(1, Source_Num_w_o - 1).Value = ����3.Range(Left(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
            ����6.Cells(1, Source_Num_w_o - 1).Style = "20% � ������1"
            ����6.Cells(1, Source_Num_w_o - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            ����6.Cells(1, Source_Num_w_o - 1).Font.Bold = True
            ����6.Cells(1, Source_Num_w_o - 1).HorizontalAlignment = xlCenter
            ����6.Cells(1, Source_Num_w_o - 1).VerticalAlignment = xlCenter
    
            ����6.Cells(2, Source_Num_w_o - 1).Value = Source_Value
            ����6.Cells(2, Source_Num_w_o - 1).Interior.ThemeColor = xlThemeColorAccent1
            ����6.Cells(2, Source_Num_w_o - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
            ����6.Cells(2, Source_Num_w_o - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            ����6.Cells(2, Source_Num_w_o - 1).borders(xlEdgeTop).LineStyle = xlContinuous
            ����6.Cells(2, Source_Num_w_o - 1).Font.Bold = True
            ����6.Cells(2, Source_Num_w_o - 1).Font.ThemeColor = xlThemeColorDark1
            ����6.Cells(2, Source_Num_w_o - 1).HorizontalAlignment = xlCenter
            ����6.Cells(2, Source_Num_w_o - 1).VerticalAlignment = xlCenter
            Source_Num_w_o = Source_Num_w_o + 1
            End If
        End If
    End If
   
    ����2.Cells(1, Source_Num - 1).Value = ����3.Range(Left(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
    ����2.Cells(1, Source_Num - 1).Style = "20% � ������1"
    ����2.Cells(1, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����2.Cells(1, Source_Num - 1).Font.Bold = True
    ����2.Cells(1, Source_Num - 1).HorizontalAlignment = xlCenter
    ����2.Cells(1, Source_Num - 1).VerticalAlignment = xlCenter

    ����2.Cells(2, Source_Num - 1).Value = Source_Value
    ����2.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
    ����2.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    ����2.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����2.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
    ����2.Cells(2, Source_Num - 1).Font.Bold = True
    ����2.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
    ����2.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
    ����2.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter

    ����14.Cells(1, Source_Num - 1).Value = ����3.Range(Left(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(����3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
    ����14.Cells(1, Source_Num - 1).Style = "20% � ������1"
    ����14.Cells(1, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����14.Cells(1, Source_Num - 1).Font.Bold = True
    ����14.Cells(1, Source_Num - 1).HorizontalAlignment = xlCenter
    ����14.Cells(1, Source_Num - 1).VerticalAlignment = xlCenter
    
    ����14.Cells(2, Source_Num - 1).Value = Source_Value
    ����14.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
    ����14.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    ����14.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����14.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
    ����14.Cells(2, Source_Num - 1).Font.Bold = True
    ����14.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
    ����14.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
    ����14.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter

    If ����3.CheckBoxes("THP_Check").Value = 1 Then
        ����11.Cells(2, Source_Num - 1).Value = Source_Value
        ����11.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
        ����11.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        ����11.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
        ����11.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
        ����11.Cells(2, Source_Num - 1).Font.Bold = True
        ����11.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
        ����11.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
        ����11.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter
    End If

    Source_Num = Source_Num + 1
    
Loop

End Sub

Sub Transfer_Sink_List()

Dim Sink_Num As Integer, Sink_Value As String

Sink_Num = 3

Do Until IsEmpty(����3.Cells(Sink_Num, 12).Value)

    Sink_Value = ����3.Cells(Sink_Num, 12).Value
    
    ����7.Cells(2, Sink_Num - 1).Value = Sink_Value
    ����7.Cells(2, Sink_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
    ����7.Cells(2, Sink_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    ����7.Cells(2, Sink_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    ����7.Cells(2, Sink_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
    ����7.Cells(2, Sink_Num - 1).Font.Bold = True
    ����7.Cells(2, Sink_Num - 1).Font.ThemeColor = xlThemeColorDark1
    ����7.Cells(2, Sink_Num - 1).HorizontalAlignment = xlCenter
    ����7.Cells(2, Sink_Num - 1).VerticalAlignment = xlCenter
    
    Sink_Num = Sink_Num + 1
    
Loop

End Sub

Sub Transfer_All_Lists()

����1.Activate
����1.Range(Cells(3, 1), Cells(1000, 1)).Clear
����1.Range(Cells(1, 2), Cells(2, 1000)).Clear
����2.Activate
����2.Range(Cells(3, 1), Cells(1000, 1)).Clear
����2.Range(Cells(1, 2), Cells(2, 1000)).Clear
����5.Activate
����5.Range(Cells(3, 1), Cells(1000, 1)).Clear
����5.Range(Cells(1, 2), Cells(2, 1000)).Clear
����6.Activate
����6.Range(Cells(3, 1), Cells(1000, 1)).Clear
����6.Range(Cells(1, 2), Cells(2, 1000)).Clear
����7.Activate
����7.Range(Cells(3, 1), Cells(1000, 1)).Clear
����7.Range(Cells(2, 2), Cells(2, 1000)).Clear
����8.Activate
����8.Range(Cells(3, 1), Cells(1000, 1)).Clear
����8.Range(Cells(2, 2), Cells(2, 1000)).Clear
����9.Activate
����9.Range(Cells(3, 1), Cells(1000, 1)).Clear
����9.Range(Cells(2, 2), Cells(2, 1000)).Clear
����10.Activate
����10.Range(Cells(3, 1), Cells(1000, 1)).Clear
����10.Range(Cells(2, 2), Cells(2, 1000)).Clear
����11.Activate
����11.Range(Cells(3, 1), Cells(1000, 1)).Clear
����11.Range(Cells(2, 2), Cells(2, 1000)).Clear
����12.Activate
����12.Range(Cells(3, 1), Cells(1000, 1)).Clear
����12.Range(Cells(2, 2), Cells(2, 1000)).Clear
����13.Activate
����13.Range(Cells(3, 1), Cells(1000, 1)).Clear
����14.Activate
����14.Range(Cells(3, 1), Cells(1000, 1)).Clear
����14.Range(Cells(1, 2), Cells(2, 1000)).Clear
����3.Activate

Call Transfer_Dates_List
Call Transfer_Pipes_List
If ����4.Cells(9, 4).Value = True Then Call Transfer_Source_Filter_List Else Call Transfer_Source_List
Call Transfer_Sink_List

End Sub




Sub borders()

����1.Activate
����1.Range(Cells(3, 2), Cells(����1.Cells(3, 1).End(xlDown).Row, ����1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
����2.Activate
����2.Range(Cells(3, 2), Cells(����1.Cells(3, 1).End(xlDown).Row, ����1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
����5.Activate
����5.Range(Cells(3, 2), Cells(����1.Cells(3, 1).End(xlDown).Row, ����1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
����6.Activate
����6.Range(Cells(3, 2), Cells(����1.Cells(3, 1).End(xlDown).Row, ����1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
����7.Activate
����7.Range(Cells(3, 2), Cells(����1.Cells(3, 1).End(xlDown).Row, ����1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
����14.Activate
����14.Range(Cells(3, 2), Cells(����1.Cells(3, 1).End(xlDown).Row, ����1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True

If ����3.CheckBoxes("Init_Vel_Gas_Check").Value = 1 Then
    ����8.Activate
    ����8.Range(Cells(3, 2), Cells(����1.Cells(3, 1).End(xlDown).Row, ����1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
End If

If ����3.CheckBoxes("Final_Vel_Gas_Check").Value = 1 Then
    ����9.Activate
    ����9.Range(Cells(3, 2), Cells(����1.Cells(3, 1).End(xlDown).Row, ����1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
End If

If ����3.CheckBoxes("Spec_DP_Check").Value = 1 Then
    ����10.Activate
    ����10.Range(Cells(3, 2), Cells(����1.Cells(3, 1).End(xlDown).Row, ����1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
End If

If ����3.CheckBoxes("THP_Check").Value = 1 Then
    ����11.Activate
    ����11.Range(Cells(3, 2), Cells(����1.Cells(3, 1).End(xlDown).Row, ����1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
End If

If ����3.CheckBoxes("EVR_Check").Value = 1 Then
    ����12.Activate
    ����12.Range(Cells(3, 2), Cells(����1.Cells(3, 1).End(xlDown).Row, ����1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
End If

If ����3.Choke_Change.Value = True Then
    ����13.Activate
    ����13.Range(Cells(3, 2), Cells(����1.Cells(3, 1).End(xlDown).Row, ����1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
End If

If ����3.CheckBoxes("Init_Vel_Liq_Check").Value = 1 Then
    ����15.Activate
    ����15.Range(Cells(3, 2), Cells(����1.Cells(3, 1).End(xlDown).Row, ����1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
End If

If ����3.CheckBoxes("Final_Vel_Liq_Check").Value = 1 Then
    ����16.Activate
    ����16.Range(Cells(3, 2), Cells(����1.Cells(3, 1).End(xlDown).Row, ����1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
End If

End Sub
