Attribute VB_Name = "Module1"
Sub Gas_Click()

Лист4.Cells(10, 4) = "Газ"
Лист3.Shapes("Group 27").Select
    With Selection.ShapeRange.Glow
        .Color.ObjectThemeColor = msoThemeColorAccent1
        .Color.TintAndShade = 0
        .Color.Brightness = 0
        .Transparency = 0.6000000238
        .Radius = 18
    End With
    
Лист3.Shapes("Group 30").Select
    With Selection.ShapeRange.Glow
        .Radius = 0
    End With
    
Лист3.GOR.Enabled = False
Лист3.GOR.Caption = "GOR"
Лист3.GOR.Value = False
Лист3.WCT.Enabled = False
Лист3.WCT.Caption = "Watercut"
Лист3.WCT.Value = False

Лист3.CGR.Enabled = True
Лист3.WGR.Enabled = True
Лист3.OilRate.Caption = "Condensate Rate"
Лист3.Cells(5, 2).Value = "Способ задания дебита конденсата:"
    With Лист3.Cells(5, 2).Characters(Start:=16, Length:=18).Font
        .Name = "Calibri"
        .FontStyle = "полужирный"
    End With

Лист3.GasRate.Value = False
Лист3.OilRate.Value = False
Лист3.CGR.Value = False
Лист3.WaterRate.Value = False
Лист3.WGR.Value = False

Cells(1, 2).Select


'With Лист11.Cells(2, 1).Validation
 '   .Delete
  '  .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
   ' xlBetween, Formula1:="=Source_Var_Gas"
    '.ShowInput = True
    '.ShowError = True
'End With

'Лист11.Cells(2, 1) = "Единый общий источник"

End Sub

Sub Oil_Click()

Лист4.Cells(10, 4) = "Нефть"
Лист3.Shapes("Group 30").Select
    With Selection.ShapeRange.Glow
        .Color.ObjectThemeColor = msoThemeColorAccent2
        .Color.TintAndShade = 0
        .Color.Brightness = 0
        .Transparency = 0.6000000238
        .Radius = 18
    End With
    
Лист3.Shapes("Group 27").Select
    With Selection.ShapeRange.Glow
        .Radius = 0
    End With

Лист3.CGR.Enabled = False
Лист3.CGR.Caption = "CGR"
Лист3.CGR.Value = False
Лист3.WGR.Enabled = False
Лист3.WGR.Caption = "WGR"
Лист3.WGR.Value = False

Лист3.GOR.Enabled = True
Лист3.WCT.Enabled = True
Лист3.OilRate.Caption = "Oil Rate"
Лист3.Cells(5, 2).Value = "Способ задания дебита нефти:"
    With Лист3.Cells(5, 2).Characters(Start:=16, Length:=12).Font
        .Name = "Calibri"
        .FontStyle = "полужирный"
    End With

Лист3.GasRate.Value = False
Лист3.GOR.Value = False
Лист3.OilRate.Value = False
Лист3.WaterRate.Value = False
Лист3.WCT.Value = False

Cells(1, 2).Select

'With Лист11.Cells(2, 1).Validation
'    .Delete
 '   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'    xlBetween, Formula1:="=Source_Var_Oil"
'    .ShowInput = True
'    .ShowError = True
'End With

'Лист11.Cells(2, 1) = "Единый общий источник"

End Sub

Sub Get_Dir_Model()      'Указание пути нахождения модели PIPESIM

Dim fd As FileDialog
Dim result As Integer
Dim fileName, FilePath, dirModel, pnsxFile, sumFile, gapFile As String
Dim position, AllPos As Long
With ActiveWorkbook.Sheets("Инициал.")
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

Do Until IsEmpty(Лист3.Cells(Date_Num, 8).Value)

    Date_Value = Лист3.Cells(Date_Num, 8).Value

    Лист1.Cells(Date_Num, 1).Value = Date_Value
    Лист1.Cells(Date_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
    Лист1.Cells(Date_Num, 1).Interior.TintAndShade = -0.499984740745262
    Лист1.Cells(Date_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист1.Cells(Date_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    Лист1.Cells(Date_Num, 1).Font.Bold = True
    Лист1.Cells(Date_Num, 1).Font.ThemeColor = xlThemeColorDark1
    Лист1.Cells(Date_Num, 1).HorizontalAlignment = xlCenter
    Лист1.Cells(Date_Num, 1).VerticalAlignment = xlCenter
    
    Лист5.Cells(Date_Num, 1).Value = Date_Value
    Лист5.Cells(Date_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
    Лист5.Cells(Date_Num, 1).Interior.TintAndShade = -0.499984740745262
    Лист5.Cells(Date_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист5.Cells(Date_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    Лист5.Cells(Date_Num, 1).Font.Bold = True
    Лист5.Cells(Date_Num, 1).Font.ThemeColor = xlThemeColorDark1
    Лист5.Cells(Date_Num, 1).HorizontalAlignment = xlCenter
    Лист5.Cells(Date_Num, 1).VerticalAlignment = xlCenter
    
    Лист6.Cells(Date_Num, 1).Value = Date_Value
    Лист6.Cells(Date_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
    Лист6.Cells(Date_Num, 1).Interior.TintAndShade = -0.499984740745262
    Лист6.Cells(Date_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист6.Cells(Date_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    Лист6.Cells(Date_Num, 1).Font.Bold = True
    Лист6.Cells(Date_Num, 1).Font.ThemeColor = xlThemeColorDark1
    Лист6.Cells(Date_Num, 1).HorizontalAlignment = xlCenter
    Лист6.Cells(Date_Num, 1).VerticalAlignment = xlCenter
    
    Лист2.Cells(Date_Num, 1).Value = Date_Value
    Лист2.Cells(Date_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
    Лист2.Cells(Date_Num, 1).Interior.TintAndShade = -0.499984740745262
    Лист2.Cells(Date_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист2.Cells(Date_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    Лист2.Cells(Date_Num, 1).Font.Bold = True
    Лист2.Cells(Date_Num, 1).Font.ThemeColor = xlThemeColorDark1
    Лист2.Cells(Date_Num, 1).HorizontalAlignment = xlCenter
    Лист2.Cells(Date_Num, 1).VerticalAlignment = xlCenter
    
    Лист14.Cells(Date_Num, 1).Value = Date_Value
    Лист14.Cells(Date_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
    Лист14.Cells(Date_Num, 1).Interior.TintAndShade = -0.499984740745262
    Лист14.Cells(Date_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист14.Cells(Date_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    Лист14.Cells(Date_Num, 1).Font.Bold = True
    Лист14.Cells(Date_Num, 1).Font.ThemeColor = xlThemeColorDark1
    Лист14.Cells(Date_Num, 1).HorizontalAlignment = xlCenter
    Лист14.Cells(Date_Num, 1).VerticalAlignment = xlCenter
    
    Лист7.Cells(Date_Num, 1).Value = Date_Value
    Лист7.Cells(Date_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
    Лист7.Cells(Date_Num, 1).Interior.TintAndShade = -0.499984740745262
    Лист7.Cells(Date_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист7.Cells(Date_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    Лист7.Cells(Date_Num, 1).Font.Bold = True
    Лист7.Cells(Date_Num, 1).Font.ThemeColor = xlThemeColorDark1
    Лист7.Cells(Date_Num, 1).HorizontalAlignment = xlCenter
    Лист7.Cells(Date_Num, 1).VerticalAlignment = xlCenter

    If Лист3.Choke_Change.Value = True Then
        Лист13.Cells(Date_Num, 1).Value = Date_Value
        Лист13.Cells(Date_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
        Лист13.Cells(Date_Num, 1).Interior.TintAndShade = -0.499984740745262
        Лист13.Cells(Date_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
        Лист13.Cells(Date_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        Лист13.Cells(Date_Num, 1).Font.Bold = True
        Лист13.Cells(Date_Num, 1).Font.ThemeColor = xlThemeColorDark1
        Лист13.Cells(Date_Num, 1).HorizontalAlignment = xlCenter
        Лист13.Cells(Date_Num, 1).VerticalAlignment = xlCenter
    End If
    If Лист3.CheckBoxes("Init_Vel_Gas_Check").Value = 1 Then
        Лист8.Cells(2, Date_Num - 1).Value = Date_Value
        Лист8.Cells(2, Date_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
        Лист8.Cells(2, Date_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
        Лист8.Cells(2, Date_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        Лист8.Cells(2, Date_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
        Лист8.Cells(2, Date_Num - 1).Font.Bold = True
        Лист8.Cells(2, Date_Num - 1).Font.ThemeColor = xlThemeColorDark1
        Лист8.Cells(2, Date_Num - 1).HorizontalAlignment = xlCenter
        Лист8.Cells(2, Date_Num - 1).VerticalAlignment = xlCenter
    End If
    If Лист3.CheckBoxes("Init_Vel_Liq_Check").Value = 1 Then
        Лист15.Cells(2, Date_Num - 1).Value = Date_Value
        Лист15.Cells(2, Date_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
        Лист15.Cells(2, Date_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
        Лист15.Cells(2, Date_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        Лист15.Cells(2, Date_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
        Лист15.Cells(2, Date_Num - 1).Font.Bold = True
        Лист15.Cells(2, Date_Num - 1).Font.ThemeColor = xlThemeColorDark1
        Лист15.Cells(2, Date_Num - 1).HorizontalAlignment = xlCenter
        Лист15.Cells(2, Date_Num - 1).VerticalAlignment = xlCenter
    End If
    If Лист3.CheckBoxes("Init_Vel_Liq_Check").Value = 1 Then
        Лист16.Cells(2, Date_Num - 1).Value = Date_Value
        Лист16.Cells(2, Date_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
        Лист16.Cells(2, Date_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
        Лист16.Cells(2, Date_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        Лист16.Cells(2, Date_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
        Лист16.Cells(2, Date_Num - 1).Font.Bold = True
        Лист16.Cells(2, Date_Num - 1).Font.ThemeColor = xlThemeColorDark1
        Лист16.Cells(2, Date_Num - 1).HorizontalAlignment = xlCenter
        Лист16.Cells(2, Date_Num - 1).VerticalAlignment = xlCenter
    End If
    If Лист3.CheckBoxes("Final_Vel_Gas_Check").Value = 1 Then
        Лист9.Cells(2, Date_Num - 1).Value = Date_Value
        Лист9.Cells(2, Date_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
        Лист9.Cells(2, Date_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
        Лист9.Cells(2, Date_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        Лист9.Cells(2, Date_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
        Лист9.Cells(2, Date_Num - 1).Font.Bold = True
        Лист9.Cells(2, Date_Num - 1).Font.ThemeColor = xlThemeColorDark1
        Лист9.Cells(2, Date_Num - 1).HorizontalAlignment = xlCenter
        Лист9.Cells(2, Date_Num - 1).VerticalAlignment = xlCenter
    End If
    If Лист3.CheckBoxes("Spec_DP_Check").Value = 1 Then
        Лист10.Cells(2, Date_Num - 1).Value = Date_Value
        Лист10.Cells(2, Date_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
        Лист10.Cells(2, Date_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
        Лист10.Cells(2, Date_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        Лист10.Cells(2, Date_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
        Лист10.Cells(2, Date_Num - 1).Font.Bold = True
        Лист10.Cells(2, Date_Num - 1).Font.ThemeColor = xlThemeColorDark1
        Лист10.Cells(2, Date_Num - 1).HorizontalAlignment = xlCenter
        Лист10.Cells(2, Date_Num - 1).VerticalAlignment = xlCenter
    End If
    If Лист3.CheckBoxes("THP_Check").Value = 1 Then
        Лист11.Cells(Date_Num, 1).Value = Date_Value
        Лист11.Cells(Date_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
        Лист11.Cells(Date_Num, 1).Interior.TintAndShade = -0.499984740745262
        Лист11.Cells(Date_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
        Лист11.Cells(Date_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        Лист11.Cells(Date_Num, 1).Font.Bold = True
        Лист11.Cells(Date_Num, 1).Font.ThemeColor = xlThemeColorDark1
        Лист11.Cells(Date_Num, 1).HorizontalAlignment = xlCenter
        Лист11.Cells(Date_Num, 1).VerticalAlignment = xlCenter
    End If
    If Лист3.CheckBoxes("EVR_Check").Value = 1 Then
        Лист12.Cells(2, Date_Num - 1).Value = Date_Value
        Лист12.Cells(2, Date_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
        Лист12.Cells(2, Date_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
        Лист12.Cells(2, Date_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        Лист12.Cells(2, Date_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
        Лист12.Cells(2, Date_Num - 1).Font.Bold = True
        Лист12.Cells(2, Date_Num - 1).Font.ThemeColor = xlThemeColorDark1
        Лист12.Cells(2, Date_Num - 1).HorizontalAlignment = xlCenter
        Лист12.Cells(2, Date_Num - 1).VerticalAlignment = xlCenter
    End If

    Date_Num = Date_Num + 1
    
Loop

End Sub

Sub Transfer_Pipes_List()

Dim Pipe_Num As Integer, Pipe_Value As String

Pipe_Num = 3

Do Until IsEmpty(Лист3.Cells(Pipe_Num, 9).Value)

    Pipe_Value = Лист3.Cells(Pipe_Num, 9).Value
    
    Лист17.Cells(Pipe_Num, 1).Value = Pipe_Value
    Лист17.Cells(Pipe_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
    Лист17.Cells(Pipe_Num, 1).Interior.TintAndShade = -0.499984740745262
    Лист17.Cells(Pipe_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист17.Cells(Pipe_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    Лист17.Cells(Pipe_Num, 1).Font.Bold = True
    Лист17.Cells(Pipe_Num, 1).Font.ThemeColor = xlThemeColorDark1
    Лист17.Cells(Pipe_Num, 1).HorizontalAlignment = xlCenter
    Лист17.Cells(Pipe_Num, 1).VerticalAlignment = xlCenter
    Лист17.Cells(Pipe_Num, 2).Style = "20% — акцент3"
    Лист17.Cells(Pipe_Num, 2).HorizontalAlignment = xlCenter
    Лист17.Cells(Pipe_Num, 2).VerticalAlignment = xlCenter
    Лист17.Cells(Pipe_Num, 2).borders(xlEdgeBottom).LineStyle = xlContinuous
    Лист17.Cells(Pipe_Num, 2).borders(xlEdgeRight).LineStyle = xlContinuous
    
    If Лист3.CheckBoxes("Init_Vel_Gas_Check").Value = 1 Then
        Лист8.Cells(Pipe_Num, 1).Value = Pipe_Value
        Лист8.Cells(Pipe_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
        Лист8.Cells(Pipe_Num, 1).Interior.TintAndShade = -0.499984740745262
        Лист8.Cells(Pipe_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
        Лист8.Cells(Pipe_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        Лист8.Cells(Pipe_Num, 1).Font.Bold = True
        Лист8.Cells(Pipe_Num, 1).Font.ThemeColor = xlThemeColorDark1
        Лист8.Cells(Pipe_Num, 1).HorizontalAlignment = xlCenter
        Лист8.Cells(Pipe_Num, 1).VerticalAlignment = xlCenter
    End If
    If Лист3.CheckBoxes("Final_Vel_Gas_Check").Value = 1 Then
        Лист9.Cells(Pipe_Num, 1).Value = Pipe_Value
        Лист9.Cells(Pipe_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
        Лист9.Cells(Pipe_Num, 1).Interior.TintAndShade = -0.499984740745262
        Лист9.Cells(Pipe_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
        Лист9.Cells(Pipe_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        Лист9.Cells(Pipe_Num, 1).Font.Bold = True
        Лист9.Cells(Pipe_Num, 1).Font.ThemeColor = xlThemeColorDark1
        Лист9.Cells(Pipe_Num, 1).HorizontalAlignment = xlCenter
        Лист9.Cells(Pipe_Num, 1).VerticalAlignment = xlCenter
    End If
    If Лист3.CheckBoxes("Init_Vel_Liq_Check").Value = 1 Then
        Лист15.Cells(Pipe_Num, 1).Value = Pipe_Value
        Лист15.Cells(Pipe_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
        Лист15.Cells(Pipe_Num, 1).Interior.TintAndShade = -0.499984740745262
        Лист15.Cells(Pipe_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
        Лист15.Cells(Pipe_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        Лист15.Cells(Pipe_Num, 1).Font.Bold = True
        Лист15.Cells(Pipe_Num, 1).Font.ThemeColor = xlThemeColorDark1
        Лист15.Cells(Pipe_Num, 1).HorizontalAlignment = xlCenter
        Лист15.Cells(Pipe_Num, 1).VerticalAlignment = xlCenter
    End If
    If Лист3.CheckBoxes("Final_Vel_Liq_Check").Value = 1 Then
        Лист16.Cells(Pipe_Num, 1).Value = Pipe_Value
        Лист16.Cells(Pipe_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
        Лист16.Cells(Pipe_Num, 1).Interior.TintAndShade = -0.499984740745262
        Лист16.Cells(Pipe_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
        Лист16.Cells(Pipe_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        Лист16.Cells(Pipe_Num, 1).Font.Bold = True
        Лист16.Cells(Pipe_Num, 1).Font.ThemeColor = xlThemeColorDark1
        Лист16.Cells(Pipe_Num, 1).HorizontalAlignment = xlCenter
        Лист16.Cells(Pipe_Num, 1).VerticalAlignment = xlCenter
    End If
    If Лист3.CheckBoxes("Spec_DP_Check").Value = 1 Then
        Лист10.Cells(Pipe_Num, 1).Value = Pipe_Value
        Лист10.Cells(Pipe_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
        Лист10.Cells(Pipe_Num, 1).Interior.TintAndShade = -0.499984740745262
        Лист10.Cells(Pipe_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
        Лист10.Cells(Pipe_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        Лист10.Cells(Pipe_Num, 1).Font.Bold = True
        Лист10.Cells(Pipe_Num, 1).Font.ThemeColor = xlThemeColorDark1
        Лист10.Cells(Pipe_Num, 1).HorizontalAlignment = xlCenter
        Лист10.Cells(Pipe_Num, 1).VerticalAlignment = xlCenter
    End If
    If Лист3.CheckBoxes("EVR_Check").Value = 1 Then
        Лист12.Cells(Pipe_Num, 1).Value = Pipe_Value
        Лист12.Cells(Pipe_Num, 1).Interior.ThemeColor = xlThemeColorAccent3
        Лист12.Cells(Pipe_Num, 1).Interior.TintAndShade = -0.499984740745262
        Лист12.Cells(Pipe_Num, 1).borders(xlEdgeRight).LineStyle = xlContinuous
        Лист12.Cells(Pipe_Num, 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        Лист12.Cells(Pipe_Num, 1).Font.Bold = True
        Лист12.Cells(Pipe_Num, 1).Font.ThemeColor = xlThemeColorDark1
        Лист12.Cells(Pipe_Num, 1).HorizontalAlignment = xlCenter
        Лист12.Cells(Pipe_Num, 1).VerticalAlignment = xlCenter
    End If

    Pipe_Num = Pipe_Num + 1
    
Loop

End Sub

Sub Transfer_Source_List()

Dim Source_Num As Integer, Source_Value As String

Source_Num = 3

Do Until IsEmpty(Лист3.Cells(Source_Num, 10).Value)

    Source_Value = Лист3.Cells(Source_Num, 10).Value
    
    If Лист3.Cells(Source_Num, 11).Value = "" Then GoTo NoneWP
    Лист1.Cells(1, Source_Num - 1) = Лист3.Range(Left(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
    Лист1.Cells(1, Source_Num - 1).Style = "20% — акцент1"
    Лист1.Cells(1, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист1.Cells(1, Source_Num - 1).Font.Bold = True
    Лист1.Cells(1, Source_Num - 1).HorizontalAlignment = xlCenter
    Лист1.Cells(1, Source_Num - 1).VerticalAlignment = xlCenter
          
    Лист5.Cells(1, Source_Num - 1).Value = Лист3.Range(Left(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
    Лист5.Cells(1, Source_Num - 1).Style = "20% — акцент1"
    Лист5.Cells(1, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист5.Cells(1, Source_Num - 1).Font.Bold = True
    Лист5.Cells(1, Source_Num - 1).HorizontalAlignment = xlCenter
    Лист5.Cells(1, Source_Num - 1).VerticalAlignment = xlCenter
        
    Лист6.Cells(1, Source_Num - 1).Value = Лист3.Range(Left(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
    Лист6.Cells(1, Source_Num - 1).Style = "20% — акцент1"
    Лист6.Cells(1, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист6.Cells(1, Source_Num - 1).Font.Bold = True
    Лист6.Cells(1, Source_Num - 1).HorizontalAlignment = xlCenter
    Лист6.Cells(1, Source_Num - 1).VerticalAlignment = xlCenter
        
    Лист2.Cells(1, Source_Num - 1).Value = Лист3.Range(Left(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
    Лист2.Cells(1, Source_Num - 1).Style = "20% — акцент1"
    Лист2.Cells(1, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист2.Cells(1, Source_Num - 1).Font.Bold = True
    Лист2.Cells(1, Source_Num - 1).HorizontalAlignment = xlCenter
    Лист2.Cells(1, Source_Num - 1).VerticalAlignment = xlCenter
        
    Лист14.Cells(1, Source_Num - 1).Value = Лист3.Range(Left(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
    Лист14.Cells(1, Source_Num - 1).Style = "20% — акцент1"
    Лист14.Cells(1, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист14.Cells(1, Source_Num - 1).Font.Bold = True
    Лист14.Cells(1, Source_Num - 1).HorizontalAlignment = xlCenter
    Лист14.Cells(1, Source_Num - 1).VerticalAlignment = xlCenter
NoneWP:

    Лист1.Cells(2, Source_Num - 1).Value = Source_Value
    Лист1.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
    Лист1.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    Лист1.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист1.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
    Лист1.Cells(2, Source_Num - 1).Font.Bold = True
    Лист1.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
    Лист1.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
    Лист1.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter
      
    Лист5.Cells(2, Source_Num - 1).Value = Source_Value
    Лист5.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
    Лист5.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    Лист5.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист5.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
    Лист5.Cells(2, Source_Num - 1).Font.Bold = True
    Лист5.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
    Лист5.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
    Лист5.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter
    
    Лист6.Cells(2, Source_Num - 1).Value = Source_Value
    Лист6.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
    Лист6.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    Лист6.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист6.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
    Лист6.Cells(2, Source_Num - 1).Font.Bold = True
    Лист6.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
    Лист6.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
    Лист6.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter
    
    Лист2.Cells(2, Source_Num - 1).Value = Source_Value
    Лист2.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
    Лист2.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    Лист2.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист2.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
    Лист2.Cells(2, Source_Num - 1).Font.Bold = True
    Лист2.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
    Лист2.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
    Лист2.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter
    
    Лист14.Cells(2, Source_Num - 1).Value = Source_Value
    Лист14.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
    Лист14.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    Лист14.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист14.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
    Лист14.Cells(2, Source_Num - 1).Font.Bold = True
    Лист14.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
    Лист14.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
    Лист14.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter
    
    If Лист3.CheckBoxes("THP_Check").Value = 1 Then
        Лист11.Cells(2, Source_Num - 1).Value = Source_Value
        Лист11.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
        Лист11.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        Лист11.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
        Лист11.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
        Лист11.Cells(2, Source_Num - 1).Font.Bold = True
        Лист11.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
        Лист11.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
        Лист11.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter
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

Do Until IsEmpty(Лист3.Cells(Source_Num, 10).Value)
    Source_Value = Лист3.Cells(Source_Num, 10).Value
    
    If Лист3.GasRate.Value = True Then
        If Right(Source_Value, 2) = "_g" Then
            Лист1.Cells(1, Source_Num_g_g - 1).Value = Лист3.Range(Left(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
            Лист1.Cells(1, Source_Num_g_g - 1).Style = "20% — акцент1"
            Лист1.Cells(1, Source_Num_g_g - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            Лист1.Cells(1, Source_Num_g_g - 1).Font.Bold = True
            Лист1.Cells(1, Source_Num_g_g - 1).HorizontalAlignment = xlCenter
            Лист1.Cells(1, Source_Num_g_g - 1).VerticalAlignment = xlCenter
            
            Лист1.Cells(2, Source_Num_g_g - 1).Value = Source_Value
            Лист1.Cells(2, Source_Num_g_g - 1).Interior.ThemeColor = xlThemeColorAccent1
            Лист1.Cells(2, Source_Num_g_g - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
            Лист1.Cells(2, Source_Num_g_g - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            Лист1.Cells(2, Source_Num_g_g - 1).borders(xlEdgeTop).LineStyle = xlContinuous
            Лист1.Cells(2, Source_Num_g_g - 1).Font.Bold = True
            Лист1.Cells(2, Source_Num_g_g - 1).Font.ThemeColor = xlThemeColorDark1
            Лист1.Cells(2, Source_Num_g_g - 1).HorizontalAlignment = xlCenter
            Лист1.Cells(2, Source_Num_g_g - 1).VerticalAlignment = xlCenter
            Source_Num_g_g = Source_Num_g_g + 1
        End If
      Else
        If Right(Source_Value, 2) = "_o" Then
            Лист1.Cells(1, Source_Num_g_o - 1).Value = Лист3.Range(Left(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
            Лист1.Cells(1, Source_Num_g_o - 1).Style = "20% — акцент1"
            Лист1.Cells(1, Source_Num_g_o - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            Лист1.Cells(1, Source_Num_g_o - 1).Font.Bold = True
            Лист1.Cells(1, Source_Num_g_o - 1).HorizontalAlignment = xlCenter
            Лист1.Cells(1, Source_Num_g_o - 1).VerticalAlignment = xlCenter
  
            Лист1.Cells(2, Source_Num_g_o - 1).Value = Source_Value
            Лист1.Cells(2, Source_Num_g_o - 1).Interior.ThemeColor = xlThemeColorAccent1
            Лист1.Cells(2, Source_Num_g_o - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
            Лист1.Cells(2, Source_Num_g_o - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            Лист1.Cells(2, Source_Num_g_o - 1).borders(xlEdgeTop).LineStyle = xlContinuous
            Лист1.Cells(2, Source_Num_g_o - 1).Font.Bold = True
            Лист1.Cells(2, Source_Num_g_o - 1).Font.ThemeColor = xlThemeColorDark1
            Лист1.Cells(2, Source_Num_g_o - 1).HorizontalAlignment = xlCenter
            Лист1.Cells(2, Source_Num_g_o - 1).VerticalAlignment = xlCenter
            Source_Num_g_o = Source_Num_g_o + 1
        End If
    End If
          
    If Лист3.OilRate.Value = True Then
        If Right(Source_Value, 2) = "_o" Or Right(Source_Value, 2) = "_c" Then
            Лист5.Cells(1, Source_Num_o_o - 1).Value = Лист3.Range(Left(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
            Лист5.Cells(1, Source_Num_o_o - 1).Style = "20% — акцент1"
            Лист5.Cells(1, Source_Num_o_o - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            Лист5.Cells(1, Source_Num_o_o - 1).Font.Bold = True
            Лист5.Cells(1, Source_Num_o_o - 1).HorizontalAlignment = xlCenter
            Лист5.Cells(1, Source_Num_o_o - 1).VerticalAlignment = xlCenter
            
            Лист5.Cells(2, Source_Num_o_o - 1).Value = Source_Value
            Лист5.Cells(2, Source_Num_o_o - 1).Interior.ThemeColor = xlThemeColorAccent1
            Лист5.Cells(2, Source_Num_o_o - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
            Лист5.Cells(2, Source_Num_o_o - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            Лист5.Cells(2, Source_Num_o_o - 1).borders(xlEdgeTop).LineStyle = xlContinuous
            Лист5.Cells(2, Source_Num_o_o - 1).Font.Bold = True
            Лист5.Cells(2, Source_Num_o_o - 1).Font.ThemeColor = xlThemeColorDark1
            Лист5.Cells(2, Source_Num_o_o - 1).HorizontalAlignment = xlCenter
            Лист5.Cells(2, Source_Num_o_o - 1).VerticalAlignment = xlCenter
            Source_Num_o_o = Source_Num_o_o + 1
        End If
    Else
        If Right(Source_Value, 2) = "_g" Then
            Лист5.Cells(1, Source_Num_o_g - 1).Value = Лист3.Range(Left(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
            Лист5.Cells(1, Source_Num_o_g - 1).Style = "20% — акцент1"
            Лист5.Cells(1, Source_Num_o_g - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            Лист5.Cells(1, Source_Num_o_g - 1).Font.Bold = True
            Лист5.Cells(1, Source_Num_o_g - 1).HorizontalAlignment = xlCenter
            Лист5.Cells(1, Source_Num_o_g - 1).VerticalAlignment = xlCenter
            
            Лист5.Cells(2, Source_Num_o_g - 1).Value = Source_Value
            Лист5.Cells(2, Source_Num_o_g - 1).Interior.ThemeColor = xlThemeColorAccent1
            Лист5.Cells(2, Source_Num_o_g - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
            Лист5.Cells(2, Source_Num_o_g - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            Лист5.Cells(2, Source_Num_o_g - 1).borders(xlEdgeTop).LineStyle = xlContinuous
            Лист5.Cells(2, Source_Num_o_g - 1).Font.Bold = True
            Лист5.Cells(2, Source_Num_o_g - 1).Font.ThemeColor = xlThemeColorDark1
            Лист5.Cells(2, Source_Num_o_g - 1).HorizontalAlignment = xlCenter
            Лист5.Cells(2, Source_Num_o_g - 1).VerticalAlignment = xlCenter
            Source_Num_o_g = Source_Num_o_g + 1
        End If
    End If
    
    If Лист3.WaterRate.Value = True Then
        If Right(Source_Value, 2) = "_w" Then
            Лист6.Cells(1, Source_Num_w_w - 1).Value = Лист3.Range(Left(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
            Лист6.Cells(1, Source_Num_w_w - 1).Style = "20% — акцент1"
            Лист6.Cells(1, Source_Num_w_w - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            Лист6.Cells(1, Source_Num_w_w - 1).Font.Bold = True
            Лист6.Cells(1, Source_Num_w_w - 1).HorizontalAlignment = xlCenter
            Лист6.Cells(1, Source_Num_w_w - 1).VerticalAlignment = xlCenter
            
            Лист6.Cells(2, Source_Num_w_w - 1).Value = Source_Value
            Лист6.Cells(2, Source_Num_w_w - 1).Interior.ThemeColor = xlThemeColorAccent1
            Лист6.Cells(2, Source_Num_w_w - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
            Лист6.Cells(2, Source_Num_w_w - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            Лист6.Cells(2, Source_Num_w_w - 1).borders(xlEdgeTop).LineStyle = xlContinuous
            Лист6.Cells(2, Source_Num_w_w - 1).Font.Bold = True
            Лист6.Cells(2, Source_Num_w_w - 1).Font.ThemeColor = xlThemeColorDark1
            Лист6.Cells(2, Source_Num_w_w - 1).HorizontalAlignment = xlCenter
            Лист6.Cells(2, Source_Num_w_w - 1).VerticalAlignment = xlCenter
            Source_Num_w_w = Source_Num_w_w + 1
        End If
    ElseIf Лист3.WGR.Value = True Then
        If Right(Source_Value, 2) = "_g" Then
            Лист6.Cells(1, Source_Num_w_g - 1).Value = Лист3.Range(Left(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
            Лист6.Cells(1, Source_Num_w_g - 1).Style = "20% — акцент1"
            Лист6.Cells(1, Source_Num_w_g - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            Лист6.Cells(1, Source_Num_w_g - 1).Font.Bold = True
            Лист6.Cells(1, Source_Num_w_g - 1).HorizontalAlignment = xlCenter
            Лист6.Cells(1, Source_Num_w_g - 1).VerticalAlignment = xlCenter
            
            Лист6.Cells(2, Source_Num_w_g - 1).Value = Source_Value
            Лист6.Cells(2, Source_Num_w_g - 1).Interior.ThemeColor = xlThemeColorAccent1
            Лист6.Cells(2, Source_Num_w_g - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
            Лист6.Cells(2, Source_Num_w_g - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            Лист6.Cells(2, Source_Num_w_g - 1).borders(xlEdgeTop).LineStyle = xlContinuous
            Лист6.Cells(2, Source_Num_w_g - 1).Font.Bold = True
            Лист6.Cells(2, Source_Num_w_g - 1).Font.ThemeColor = xlThemeColorDark1
            Лист6.Cells(2, Source_Num_w_g - 1).HorizontalAlignment = xlCenter
            Лист6.Cells(2, Source_Num_w_g - 1).VerticalAlignment = xlCenter
            Source_Num_w_g = Source_Num_w_g + 1
    Else
        If Right(Source_Value, 2) = "_o" Then
            Лист6.Cells(1, Source_Num_w_o - 1).Value = Лист3.Range(Left(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
            Лист6.Cells(1, Source_Num_w_o - 1).Style = "20% — акцент1"
            Лист6.Cells(1, Source_Num_w_o - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            Лист6.Cells(1, Source_Num_w_o - 1).Font.Bold = True
            Лист6.Cells(1, Source_Num_w_o - 1).HorizontalAlignment = xlCenter
            Лист6.Cells(1, Source_Num_w_o - 1).VerticalAlignment = xlCenter
    
            Лист6.Cells(2, Source_Num_w_o - 1).Value = Source_Value
            Лист6.Cells(2, Source_Num_w_o - 1).Interior.ThemeColor = xlThemeColorAccent1
            Лист6.Cells(2, Source_Num_w_o - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
            Лист6.Cells(2, Source_Num_w_o - 1).borders(xlEdgeRight).LineStyle = xlContinuous
            Лист6.Cells(2, Source_Num_w_o - 1).borders(xlEdgeTop).LineStyle = xlContinuous
            Лист6.Cells(2, Source_Num_w_o - 1).Font.Bold = True
            Лист6.Cells(2, Source_Num_w_o - 1).Font.ThemeColor = xlThemeColorDark1
            Лист6.Cells(2, Source_Num_w_o - 1).HorizontalAlignment = xlCenter
            Лист6.Cells(2, Source_Num_w_o - 1).VerticalAlignment = xlCenter
            Source_Num_w_o = Source_Num_w_o + 1
            End If
        End If
    End If
   
    Лист2.Cells(1, Source_Num - 1).Value = Лист3.Range(Left(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
    Лист2.Cells(1, Source_Num - 1).Style = "20% — акцент1"
    Лист2.Cells(1, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист2.Cells(1, Source_Num - 1).Font.Bold = True
    Лист2.Cells(1, Source_Num - 1).HorizontalAlignment = xlCenter
    Лист2.Cells(1, Source_Num - 1).VerticalAlignment = xlCenter

    Лист2.Cells(2, Source_Num - 1).Value = Source_Value
    Лист2.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
    Лист2.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    Лист2.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист2.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
    Лист2.Cells(2, Source_Num - 1).Font.Bold = True
    Лист2.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
    Лист2.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
    Лист2.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter

    Лист14.Cells(1, Source_Num - 1).Value = Лист3.Range(Left(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), InStr(Лист3.Cells(Source_Num, 11).MergeArea.Address(0, 0), ":") - 1)).Value
    Лист14.Cells(1, Source_Num - 1).Style = "20% — акцент1"
    Лист14.Cells(1, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист14.Cells(1, Source_Num - 1).Font.Bold = True
    Лист14.Cells(1, Source_Num - 1).HorizontalAlignment = xlCenter
    Лист14.Cells(1, Source_Num - 1).VerticalAlignment = xlCenter
    
    Лист14.Cells(2, Source_Num - 1).Value = Source_Value
    Лист14.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
    Лист14.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    Лист14.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист14.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
    Лист14.Cells(2, Source_Num - 1).Font.Bold = True
    Лист14.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
    Лист14.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
    Лист14.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter

    If Лист3.CheckBoxes("THP_Check").Value = 1 Then
        Лист11.Cells(2, Source_Num - 1).Value = Source_Value
        Лист11.Cells(2, Source_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
        Лист11.Cells(2, Source_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
        Лист11.Cells(2, Source_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
        Лист11.Cells(2, Source_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
        Лист11.Cells(2, Source_Num - 1).Font.Bold = True
        Лист11.Cells(2, Source_Num - 1).Font.ThemeColor = xlThemeColorDark1
        Лист11.Cells(2, Source_Num - 1).HorizontalAlignment = xlCenter
        Лист11.Cells(2, Source_Num - 1).VerticalAlignment = xlCenter
    End If

    Source_Num = Source_Num + 1
    
Loop

End Sub

Sub Transfer_Sink_List()

Dim Sink_Num As Integer, Sink_Value As String

Sink_Num = 3

Do Until IsEmpty(Лист3.Cells(Sink_Num, 12).Value)

    Sink_Value = Лист3.Cells(Sink_Num, 12).Value
    
    Лист7.Cells(2, Sink_Num - 1).Value = Sink_Value
    Лист7.Cells(2, Sink_Num - 1).Interior.ThemeColor = xlThemeColorAccent1
    Лист7.Cells(2, Sink_Num - 1).borders(xlEdgeBottom).LineStyle = xlContinuous
    Лист7.Cells(2, Sink_Num - 1).borders(xlEdgeRight).LineStyle = xlContinuous
    Лист7.Cells(2, Sink_Num - 1).borders(xlEdgeTop).LineStyle = xlContinuous
    Лист7.Cells(2, Sink_Num - 1).Font.Bold = True
    Лист7.Cells(2, Sink_Num - 1).Font.ThemeColor = xlThemeColorDark1
    Лист7.Cells(2, Sink_Num - 1).HorizontalAlignment = xlCenter
    Лист7.Cells(2, Sink_Num - 1).VerticalAlignment = xlCenter
    
    Sink_Num = Sink_Num + 1
    
Loop

End Sub

Sub Transfer_All_Lists()

Лист1.Activate
Лист1.Range(Cells(3, 1), Cells(1000, 1)).Clear
Лист1.Range(Cells(1, 2), Cells(2, 1000)).Clear
Лист2.Activate
Лист2.Range(Cells(3, 1), Cells(1000, 1)).Clear
Лист2.Range(Cells(1, 2), Cells(2, 1000)).Clear
Лист5.Activate
Лист5.Range(Cells(3, 1), Cells(1000, 1)).Clear
Лист5.Range(Cells(1, 2), Cells(2, 1000)).Clear
Лист6.Activate
Лист6.Range(Cells(3, 1), Cells(1000, 1)).Clear
Лист6.Range(Cells(1, 2), Cells(2, 1000)).Clear
Лист7.Activate
Лист7.Range(Cells(3, 1), Cells(1000, 1)).Clear
Лист7.Range(Cells(2, 2), Cells(2, 1000)).Clear
Лист8.Activate
Лист8.Range(Cells(3, 1), Cells(1000, 1)).Clear
Лист8.Range(Cells(2, 2), Cells(2, 1000)).Clear
Лист9.Activate
Лист9.Range(Cells(3, 1), Cells(1000, 1)).Clear
Лист9.Range(Cells(2, 2), Cells(2, 1000)).Clear
Лист10.Activate
Лист10.Range(Cells(3, 1), Cells(1000, 1)).Clear
Лист10.Range(Cells(2, 2), Cells(2, 1000)).Clear
Лист11.Activate
Лист11.Range(Cells(3, 1), Cells(1000, 1)).Clear
Лист11.Range(Cells(2, 2), Cells(2, 1000)).Clear
Лист12.Activate
Лист12.Range(Cells(3, 1), Cells(1000, 1)).Clear
Лист12.Range(Cells(2, 2), Cells(2, 1000)).Clear
Лист13.Activate
Лист13.Range(Cells(3, 1), Cells(1000, 1)).Clear
Лист14.Activate
Лист14.Range(Cells(3, 1), Cells(1000, 1)).Clear
Лист14.Range(Cells(1, 2), Cells(2, 1000)).Clear
Лист3.Activate

Call Transfer_Dates_List
Call Transfer_Pipes_List
If Лист4.Cells(9, 4).Value = True Then Call Transfer_Source_Filter_List Else Call Transfer_Source_List
Call Transfer_Sink_List

End Sub




Sub borders()

Лист1.Activate
Лист1.Range(Cells(3, 2), Cells(Лист1.Cells(3, 1).End(xlDown).Row, Лист1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
Лист2.Activate
Лист2.Range(Cells(3, 2), Cells(Лист1.Cells(3, 1).End(xlDown).Row, Лист1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
Лист5.Activate
Лист5.Range(Cells(3, 2), Cells(Лист1.Cells(3, 1).End(xlDown).Row, Лист1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
Лист6.Activate
Лист6.Range(Cells(3, 2), Cells(Лист1.Cells(3, 1).End(xlDown).Row, Лист1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
Лист7.Activate
Лист7.Range(Cells(3, 2), Cells(Лист1.Cells(3, 1).End(xlDown).Row, Лист1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
Лист14.Activate
Лист14.Range(Cells(3, 2), Cells(Лист1.Cells(3, 1).End(xlDown).Row, Лист1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True

If Лист3.CheckBoxes("Init_Vel_Gas_Check").Value = 1 Then
    Лист8.Activate
    Лист8.Range(Cells(3, 2), Cells(Лист1.Cells(3, 1).End(xlDown).Row, Лист1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
End If

If Лист3.CheckBoxes("Final_Vel_Gas_Check").Value = 1 Then
    Лист9.Activate
    Лист9.Range(Cells(3, 2), Cells(Лист1.Cells(3, 1).End(xlDown).Row, Лист1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
End If

If Лист3.CheckBoxes("Spec_DP_Check").Value = 1 Then
    Лист10.Activate
    Лист10.Range(Cells(3, 2), Cells(Лист1.Cells(3, 1).End(xlDown).Row, Лист1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
End If

If Лист3.CheckBoxes("THP_Check").Value = 1 Then
    Лист11.Activate
    Лист11.Range(Cells(3, 2), Cells(Лист1.Cells(3, 1).End(xlDown).Row, Лист1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
End If

If Лист3.CheckBoxes("EVR_Check").Value = 1 Then
    Лист12.Activate
    Лист12.Range(Cells(3, 2), Cells(Лист1.Cells(3, 1).End(xlDown).Row, Лист1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
End If

If Лист3.Choke_Change.Value = True Then
    Лист13.Activate
    Лист13.Range(Cells(3, 2), Cells(Лист1.Cells(3, 1).End(xlDown).Row, Лист1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
End If

If Лист3.CheckBoxes("Init_Vel_Liq_Check").Value = 1 Then
    Лист15.Activate
    Лист15.Range(Cells(3, 2), Cells(Лист1.Cells(3, 1).End(xlDown).Row, Лист1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
End If

If Лист3.CheckBoxes("Final_Vel_Liq_Check").Value = 1 Then
    Лист16.Activate
    Лист16.Range(Cells(3, 2), Cells(Лист1.Cells(3, 1).End(xlDown).Row, Лист1.Cells(2, 2).End(xlToRight).Column)).borders.LineStyle = True
End If

End Sub
