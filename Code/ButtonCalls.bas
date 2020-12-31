Attribute VB_Name = "ButtonCalls"
Public SelectionShapeID As String
Public SelectionShapeSlideIndex As Long
Public SelectionShapeIDNumber As Long
Public PrimaryShape As shape
Public PrimaryHeight As Long
Public PrimaryWidth As Long
Dim CusDataLabel As New CustomDataLabel
'To determine if selection is a valid shape
Public Function GetShape() As shape

  Dim shp
  Dim Selected_Shape
  
  'Determine Which Shape is Active
  If True Then 'ActiveWindow.Selection.Type = ppSelectionShapes Then
    'Loop in case multiples shapes selected
       For Each shp In ActiveWindow.Selection.ShapeRange
         'ActiveShape is first shape selected
            SelectionShapeSlideIndex = Application.ActiveWindow.View.Slide.SlideIndex
            SelectionShapeID = shp.Name
            SelectionShapeIDNumber = shp.Id
            Exit For
       Next shp
  Else
    MsgBox "Only shapes are supported.", vbCritical, "Sorry"
    End
  End If
  
  Set GetShape = shp
      
End Function

Public Sub Error_On_Apply()

    MsgBox "The secondary shape you have chosen has a property that cannot be applied from the primary shape. Some restrictions " & _
           "apply for VBA where attributes cannot be copied as well. If anything has been applied incorrectly, undo it. " & vbNewLine & vbNewLine & _
           "The current primary shape information has also been wiped. Please re-select it.", vbCritical, "Note"
    
    End

End Sub

Public Sub NoPrimaryShape()

    MsgBox "It seems like you haven't chosen a primary shape. Please do so. If this is in error, please file an issue " & _
           "on Github with an example where it can be reproduced.", vbExclamation, "Note"
           
    End

End Sub

'Callback for LockShape onAction
Sub LockShape(control As IRibbonControl)
    
  Set PrimaryShape = GetShape
  
  Exit Sub
    
End Sub

'Callback for SetHeight onAction
Sub SetHeight(control As IRibbonControl)

  On Error GoTo ErrHand

  If Not PrimaryShape Is Nothing Then
    With GetShape
        .Height = PrimaryShape.Height
    End With
  Else
    NoPrimaryShape
  End If
  
  Exit Sub
  
ErrHand:
  Error_On_Apply

End Sub

'Callback for SetWidth onAction
Sub SetWidth(control As IRibbonControl)
  
  On Error GoTo ErrHand

  With GetShape
    .Width = PrimaryShape.Width
  End With

  
  Exit Sub
  
ErrHand:
  Error_On_Apply

End Sub

'Callback for SetDimesion onAction
Sub SetDimension(control As IRibbonControl)

  On Error GoTo ErrHand

  If Not PrimaryShape Is Nothing Then
    With GetShape
        .Width = PrimaryShape.Width
        .Height = PrimaryShape.Height
    End With
  Else
    NoPrimaryShape
  End If
   
  Exit Sub
  
ErrHand:
  Error_On_Apply

End Sub

'Callback for SetPosition onAction
Sub SetPosition(control As IRibbonControl)
  
  On Error GoTo ErrHand
  
  If Not PrimaryShape Is Nothing Then
    With GetShape
        .Left = PrimaryShape.Left
        .Top = PrimaryShape.Top
    End With
  Else
    NoPrimaryShape
  End If
  
  Exit Sub
  
ErrHand:
  Error_On_Apply

End Sub

'Callback for SetColour onAction
Sub SetFill(control As IRibbonControl)
  
  On Error GoTo ErrHand
  
  If Not PrimaryShape Is Nothing Then
    With GetShape
        .Fill.BackColor.RGB = PrimaryShape.Fill.BackColor.RGB
        .Fill.ForeColor.RGB = PrimaryShape.Fill.ForeColor.RGB
        .Fill.Transparency = PrimaryShape.Fill.Transparency
        If PrimaryShape.Fill.GradientColorType = msoGradientOneColor Or PrimaryShape.Fill.GradientColorType = msoGradientMultiColor Then
            .Fill.OneColorGradient PrimaryShape.Fill.GradientStyle, _
                PrimaryShape.Fill.GradientVariant, _
                PrimaryShape.Fill.GradientDegree
        ElseIf PrimaryShape.Fill.GradientColorType = msoGradientPresetColors Then
            .Fill.PresetGradient PrimaryShape.Fill.GradientStyle, _
                PrimaryShape.Fill.GradientVariant, _
                PrimaryShape.Fill.PresetGradientType
        ElseIf PrimaryShape.Fill.GradientColorType = msoGradientTwoColors Then
            .Fill.TwoColorGradient PrimaryShape.Fill.GradientStyle, _
                PrimaryShape.Fill.GradientVariant
        End If
    End With
  Else
    NoPrimaryShape
  End If
  
  Exit Sub
  
ErrHand:
  Error_On_Apply

End Sub

'Callback for SetOutline onAction
Sub SetOutline(control As IRibbonControl)

  On Error GoTo ErrHand
  
  If Not PrimaryShape Is Nothing Then
    With GetShape
        .Line.DashStyle = PrimaryShape.Line.DashStyle
        'Weight must be set first, then colours, otherwise the colour does not change
        'As noted by a not-very-upvoted answer on https://stackoverflow.com/questions/15624199/vba-power-point-changing-image-border-color-on-click
        .Line.Weight = PrimaryShape.Line.Weight
        .Line.Style = PrimaryShape.Line.Style
        .Line.Transparency = PrimaryShape.Line.Transparency
        .Line.ForeColor.RGB = PrimaryShape.Line.ForeColor.RGB
        .Line.BackColor.RGB = PrimaryShape.Line.BackColor.RGB
    End With
  Else
    NoPrimaryShape
  End If
  
  Exit Sub
  
ErrHand:
  Error_On_Apply

End Sub


'Callback for SetDimCol onAction
Sub SetDimCol(control As IRibbonControl)
    
  If Not PrimaryShape Is Nothing Then
    With GetShape
    
    On Error GoTo ErrHand
        
        .Width = PrimaryShape.Width
        .Height = PrimaryShape.Height
        
        .Fill.BackColor.RGB = PrimaryShape.Fill.BackColor.RGB
        .Fill.ForeColor.RGB = PrimaryShape.Fill.ForeColor.RGB
        .Fill.Transparency = PrimaryShape.Fill.Transparency
        If PrimaryShape.Fill.GradientColorType = msoGradientOneColor Or PrimaryShape.Fill.GradientColorType = msoGradientMultiColor Then
            .Fill.OneColorGradient PrimaryShape.Fill.GradientStyle, _
                PrimaryShape.Fill.GradientVariant, _
                PrimaryShape.Fill.GradientDegree
        ElseIf PrimaryShape.Fill.GradientColorType = msoGradientPresetColors Then
            .Fill.PresetGradient PrimaryShape.Fill.GradientStyle, _
                PrimaryShape.Fill.GradientVariant, _
                PrimaryShape.Fill.PresetGradientType
        ElseIf PrimaryShape.Fill.GradientColorType = msoGradientTwoColors Then
            .Fill.TwoColorGradient PrimaryShape.Fill.GradientStyle, _
                PrimaryShape.Fill.GradientVariant
        End If
        
        .Line.DashStyle = PrimaryShape.Line.DashStyle
        'Weight must be set first, then colours, otherwise the colour does not change
        'As noted by a not-very-upvoted answer on https://stackoverflow.com/questions/15624199/vba-power-point-changing-image-border-color-on-click
        .Line.Weight = PrimaryShape.Line.Weight
        .Line.Style = PrimaryShape.Line.Style
        .Line.Transparency = PrimaryShape.Line.Transparency
        .Line.ForeColor.RGB = PrimaryShape.Line.ForeColor.RGB
        .Line.BackColor.RGB = PrimaryShape.Line.BackColor.RGB
    
    End With
  Else
    NoPrimaryShape
  End If
  
  Exit Sub

ErrHand:
  Error_On_Apply

End Sub

'Callback for SetDimPos onAction
Sub SetDimPos(control As IRibbonControl)
    
  On Error GoTo ErrHand
  
  If Not PrimaryShape Is Nothing Then
    With GetShape
        .Left = PrimaryShape.Left
        .Top = PrimaryShape.Top
        .Width = PrimaryShape.Width
        .Height = PrimaryShape.Height
    End With
  Else
    NoPrimaryShape
  End If
  
  Exit Sub
  
ErrHand:
  Error_On_Apply
    

End Sub

'Callback for SetAll onAction
Sub SetAll(control As IRibbonControl)
    
  If Not PrimaryShape Is Nothing Then
    With GetShape
    
    On Error GoTo ErrHand
        
        .Width = PrimaryShape.Width
        .Height = PrimaryShape.Height
        
        .Fill.BackColor.RGB = PrimaryShape.Fill.BackColor.RGB
        .Fill.ForeColor.RGB = PrimaryShape.Fill.ForeColor.RGB
        .Fill.Transparency = PrimaryShape.Fill.Transparency
        If PrimaryShape.Fill.GradientColorType = msoGradientOneColor Or PrimaryShape.Fill.GradientColorType = msoGradientMultiColor Then
            .Fill.OneColorGradient PrimaryShape.Fill.GradientStyle, _
                PrimaryShape.Fill.GradientVariant, _
                PrimaryShape.Fill.GradientDegree
        ElseIf PrimaryShape.Fill.GradientColorType = msoGradientPresetColors Then
            .Fill.PresetGradient PrimaryShape.Fill.GradientStyle, _
                PrimaryShape.Fill.GradientVariant, _
                PrimaryShape.Fill.PresetGradientType
        ElseIf PrimaryShape.Fill.GradientColorType = msoGradientTwoColors Then
            .Fill.TwoColorGradient PrimaryShape.Fill.GradientStyle, _
                PrimaryShape.Fill.GradientVariant
        End If
        
        .Line.DashStyle = PrimaryShape.Line.DashStyle
        'Weight must be set first, then colours, otherwise the colour does not change
        'As noted by a not-very-upvoted answer on https://stackoverflow.com/questions/15624199/vba-power-point-changing-image-border-color-on-click
        .Line.Weight = PrimaryShape.Line.Weight
        .Line.Style = PrimaryShape.Line.Style
        .Line.Transparency = PrimaryShape.Line.Transparency
        .Line.ForeColor.RGB = PrimaryShape.Line.ForeColor.RGB
        .Line.BackColor.RGB = PrimaryShape.Line.BackColor.RGB
        
        .Left = PrimaryShape.Left
        .Top = PrimaryShape.Top
    
    End With
  Else
    NoPrimaryShape
  End If
  
  Exit Sub
  
ErrHand:
  Error_On_Apply
        
End Sub
'Callback for SyncValueAxis onAction
Sub SyncValueAxis(control As IRibbonControl)

  Dim y_primary As Boolean
  Dim y_secondary As Boolean
  
  y_primary = False
  y_secondary = False

  If Not PrimaryShape Is Nothing Then
    With GetShape
        
        If .HasChart = msoFalse Or PrimaryShape.HasChart = msoFalse Then
            MsgBox "This is not a chart. Please select charts.", vbCritical, "Note"
            Exit Sub
        
        Else
        
            If PrimaryShape.Chart.HasAxis(xlValue, xlPrimary) Then y_primary = True
        
            With .Chart
                .HasAxis(xlValue, xlPrimary) = True
                PrimaryShape.Chart.HasAxis(xlValue, xlPrimary) = True
                .Axes(xlValue, xlPrimary).MinimumScale = _
                  PrimaryShape.Chart.Axes(xlValue, xlPrimary).MinimumScale
                .Axes(xlValue, xlPrimary).MaximumScale = _
                  PrimaryShape.Chart.Axes(xlValue, xlPrimary).MaximumScale
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = _
                  PrimaryShape.Chart.Axes(xlValue, xlPrimary).TickLabels.NumberFormat
            
                If y_primary = False Then
                    .HasAxis(xlValue, xlPrimary) = False
                    PrimaryShape.Chart.HasAxis(xlValue, xlPrimary) = False
                End If
            
            End With
        End If
    End With
  Else
    NoPrimaryShape
  End If
  
  Exit Sub

no_value_based_y_axes:
  MsgBox "This chart does have a value based axis, and so cannot be synced.", vbCritical, "Note"

End Sub
'Callback for SyncDateAxis onAction
Sub SyncDateAxis(control As IRibbonControl)

  Dim x_primary As Boolean
  Dim x_secondary As Boolean
  
  x_primary = False
  x_secondary = False
  
  On Error GoTo not_possible_x_axis

  If Not PrimaryShape Is Nothing Then
    With GetShape
        
        If .HasChart = msoFalse Or PrimaryShape.HasChart = msoFalse Then
            MsgBox "This is not a chart. Please select charts.", vbCritical, "Note"
            Exit Sub
        
        Else
        
            If PrimaryShape.Chart.HasAxis(xlCategory, xlPrimary) Then x_primary = True
        
            With .Chart
                .HasAxis(xlCategory, xlPrimary) = True
                PrimaryShape.Chart.HasAxis(xlCategory, xlPrimary) = True
                .Axes(xlCategory, xlPrimary).MinimumScale = _
                  PrimaryShape.Chart.Axes(xlCategory, xlPrimary).MinimumScale
                .Axes(xlCategory, xlPrimary).MaximumScale = _
                  PrimaryShape.Chart.Axes(xlCategory, xlPrimary).MaximumScale
                .Axes(xlCategory, xlPrimary).TickLabels.NumberFormat = _
                  PrimaryShape.Chart.Axes(xlCategory, xlPrimary).TickLabels.NumberFormat
            
                If x_primary = False Then
                    .HasAxis(xlValue, xlPrimary) = False
                    PrimaryShape.Chart.HasAxis(xlCategory, xlPrimary) = True
                End If
            
            End With
        End If
    End With
  Else
    NoPrimaryShape
  End If
  
  Exit Sub

not_possible_x_axis:
  MsgBox "This chart does not have a value or timeline based axis.", vbCritical, "Note"

End Sub
'Callback for SyncDateAxis onAction
Sub SyncPlotArea(control As IRibbonControl)

  If Not PrimaryShape Is Nothing Then
    With GetShape
        
        If .HasChart = msoFalse Or PrimaryShape.HasChart = msoFalse Then
            
            MsgBox "Either the primary or secondary shape is not a chart. Please select charts.", vbCritical, "Note"
            Exit Sub
        
        Else
                
            .Height = PrimaryShape.Height
            .Width = PrimaryShape.Width
            .Chart.PlotArea.Height = PrimaryShape.Chart.PlotArea.Height
            .Chart.PlotArea.Width = PrimaryShape.Chart.PlotArea.Width
            .Chart.PlotArea.Left = PrimaryShape.Chart.PlotArea.Left
            .Chart.PlotArea.Top = PrimaryShape.Chart.PlotArea.Top
            
        End If
    End With
  Else
    NoPrimaryShape
  End If
  
  Exit Sub

End Sub

'Callback for SyncTitleArea onAction
Sub SyncTitleArea(control As IRibbonControl)

  If Not PrimaryShape Is Nothing Then
    With GetShape
        
        If .HasChart = msoFalse Or PrimaryShape.HasChart = msoFalse Then
            
            MsgBox "This is not a chart. Please select charts.", vbCritical, "Note"
            Exit Sub
            
        ElseIf PrimaryShape.Chart.HasTitle = False Then
            
            .Chart.HasTitle = False
        
        Else
                
            .Chart.ChartTitle.Top = PrimaryShape.Chart.ChartTitle.Top + _
                                    (PrimaryShape.Chart.ChartTitle.Height / 2) - _
                                    (.Chart.ChartTitle.Height / 2)
            .Chart.ChartTitle.Left = PrimaryShape.Chart.ChartTitle.Left + _
                                    (PrimaryShape.Chart.ChartTitle.Width / 2) - _
                                    (.Chart.ChartTitle.Width / 2)
            
        End If
    End With
  Else
    NoPrimaryShape
  End If
  
  Exit Sub

End Sub

Sub SyncLegendArea(control As IRibbonControl)

  If Not PrimaryShape Is Nothing Then
    With GetShape
        
        If .HasChart = msoFalse Or PrimaryShape.HasChart = msoFalse Then
            
            MsgBox "This is not a chart. Please select charts.", vbCritical, "Note"
            Exit Sub
            
        ElseIf PrimaryShape.Chart.HasLegend = False Then
            
            .Chart.HasLegend = False
        
        Else
            
            .Chart.HasLegend = True
            .Chart.Legend.Position = PrimaryShape.Chart.Legend.Position
            .Chart.Legend.Height = PrimaryShape.Chart.Legend.Height
            .Chart.Legend.Width = PrimaryShape.Chart.Legend.Width
            .Chart.Legend.Top = PrimaryShape.Chart.Legend.Top
            .Chart.Legend.Left = PrimaryShape.Chart.Legend.Left
            
        End If
    End With
  Else
    NoPrimaryShape
  End If
  
  Exit Sub

End Sub


'Callback for SetColWidths onAction
Sub SyncColumnWidth(control As IRibbonControl)

  If Not PrimaryShape Is Nothing Then
    
    With GetShape
    
      If .HasTable = msoFalse Or PrimaryShape.HasTable = msoFalse Then
        MsgBox "The shape(s) you have selected is not a table.", vbCritical, "Note"
      
      Else
        
        If .Table.Columns.Count <> PrimaryShape.Table.Columns.Count Then
            
            MsgBox "The two tables you have chosen do not have identical column counts. " & _
                   "No width adjustment will be made.", vbCritical, "Note"
            
        Else
        
            For i = 1 To .Table.Columns.Count
                
                .Table.Columns(i).Width = PrimaryShape.Table.Columns(i).Width
            
            Next i
        
        End If
        
      End If
    
    End With
    
  End If

End Sub

'Callback for SetRowHeight onAction
Sub SyncRowHeight(control As IRibbonControl)

  If Not PrimaryShape Is Nothing Then
    
    With GetShape
    
      If .HasTable = msoFalse Or PrimaryShape.HasTable = msoFalse Then
        MsgBox "The shape(s) you have selected is not a table.", vbCritical, "Note"
      
      Else
        
        If .Table.Rows.Count <> PrimaryShape.Table.Rows.Count Then
            
            MsgBox "The two tables you have chosen do not have identical row counts. " & _
                   "No height adjustment will be made.", vbCritical, "Note"
            
        Else
        
            For i = 1 To .Table.Rows.Count
                
                .Table.Rows(i).Height = PrimaryShape.Table.Rows(i).Height
            
            Next i
        
        End If
        
      End If
    
    End With
    
  End If

End Sub

'Callback for SetTableDims onAction
Sub SyncTableDims(control As IRibbonControl)

  If Not PrimaryShape Is Nothing Then
    
    With GetShape
    
      If .HasTable = msoFalse Or PrimaryShape.HasTable = msoFalse Then
        MsgBox "The shape(s) you have selected is not a table.", vbCritical, "Note"
      
      Else
        
        If .Table.Rows.Count <> PrimaryShape.Table.Rows.Count Or _
           .Table.Columns.Count <> PrimaryShape.Table.Columns.Count Then
            
            MsgBox "The two tables you have chosen do not have identical row and/or columns counts. " & _
                   "No height/width adjustment will be made.", vbCritical, "Note"
            
        Else
        
            For i = 1 To .Table.Rows.Count
                
                .Table.Rows(i).Height = PrimaryShape.Table.Rows(i).Height
            
            Next i
            
            For i = 1 To .Table.Columns.Count
                
                .Table.Columns(i).Width = PrimaryShape.Table.Columns(i).Width
            
            Next i
        
        End If
        
      End If
    
    End With
    
  End If

End Sub

Sub CustomizeDataLabels(control As IRibbonControl)

    Set PrimaryShape = GetShape()
    
    If PrimaryShape.HasChart = msoFalse Then
        MsgBox "This feature requires a chart to be selected. Please select a chart.", vbCritical, "Note"
        Exit Sub
    End If
    
    CusDataLabel.UserForm_Activate
    CusDataLabel.show vbModeless

End Sub

Sub RerunCustomLabels(control As IRibbonControl)

    Set PrimaryShape = GetShape()
    
    If PrimaryShape.HasChart = msoFalse Then
        MsgBox "This feature requires a chart to be selected. Please select a chart.", vbCritical, "Note"
        Exit Sub
    End If

    Call CusDataLabel.RunWithoutSave_Click
    Call CusDataLabel.UserForm_Activate

End Sub


Sub FormatPainter(control As IRibbonControl)

    If Not PrimaryShape Is Nothing Then
        PrimaryShape.PickUp
        GetShape.Apply
    End If
    
End Sub


Sub ResetAxis(control As IRibbonControl)
    
    Set PrimaryShape = GetShape()
    
    With PrimaryShape.Chart
    
        Dim HadAxis As Boolean
        HadAxis = False
        
        If .HasAxis(xlValue, xlPrimary) = False Then
            .HasAxis(xlValue, xlPrimary) = True
            HadAxis = True
        End If
            
        .HasAxis(xlValue, xlPrimary) = True
        .Axes(xlValue).MinimumScaleIsAuto = True
        .Axes(xlValue).MaximumScaleIsAuto = True
        
        
        If HadAxis = True Then
            .HasAxis(xlValue, xlPrimary) = True
        End If
    
    End With
    
End Sub

