Attribute VB_Name = "CustomDataLabel"
Attribute VB_Base = "0{FB1A5CF6-B322-4663-8D98-1EA550257031}{5DE1A105-7A12-40DC-A8A7-EB3034AE1800}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub HideUserForm_Click()

    Me.Hide

End Sub

Private Sub FillEveryNValuesCheckbox_AfterUpdate()

    If FillEveryNValuesCheckbox.Value = True Then
        FillEveryNValuesTextBox.Enabled = True
        FillEveryNValueHelpLabel.Enabled = True
        FillEveryNValuesTextBox.Text = 2
        FillEveryNValuesTextBoxOffset.Enabled = True
        FillEveryNValuesTextBoxOffset.Text = 0
    Else
        FillEveryNValuesTextBox.Enabled = False
        FillEveryNValueHelpLabel.Enabled = False
        FillEveryNValuesTextBoxOffset.Enabled = False
    End If

End Sub
Private Sub FillEveryNValuesTextBoxOffset_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If FillEveryNValuesTextBoxOffset.Text = "" Then
       FillEveryNValuesTextBoxOffset.Text = 0
    ElseIf Not IsNumeric(FillEveryNValuesTextBoxOffset.Value) Then
        MsgBox "Please input only integers here.", vbCritical, "Note"
        Cancel = True
    End If

End Sub

Private Sub FillEveryNValuesTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If FillEveryNValuesTextBox.Text = "" Then
       FillEveryNValuesTextBox.Text = 2
    ElseIf Not IsNumeric(FillEveryNValuesTextBox.Value) Then
        MsgBox "Please input only integers here.", vbCritical, "Note"
        Cancel = True
    End If

End Sub

Private Sub FilterBetweenCheckBox_AfterUpdate()

    If FilterBetweenCheckBox.Value = True Then
        
        FilterBetweenLowerBound.Enabled = True
        FilterBetweenUpperBound.Enabled = True
        FilterBetweenLabel.Enabled = True
    
    Else
    
        FilterBetweenLowerBound.Enabled = False
        FilterBetweenUpperBound.Enabled = False
        FilterBetweenLabel.Enabled = True
    
    End If

End Sub

Private Sub FilterBetweenLowerBound_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsNumeric(FilterBetweenLowerBound.Value) Then
        MsgBox "Please input a number here, or disable the option to filter between values after you enter a number here.", vbCritical, "Note"
        Cancel = True
    End If

End Sub

Private Sub FilterBetweenUpperBound_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsNumeric(FilterBetweenUpperBound.Value) Then
        MsgBox "Please input a number here, or disable the option to filter between values after you enter a number here.", vbCritical, "Note"
        Cancel = True
    End If

End Sub

Private Sub FirstValueCheckBox_AfterUpdate()

    If FirstValueCheckBox.Value = True Or LastValueCheckBox = True Then
        AlignFirstLastCheckBox.Enabled = True
    Else
        AlignFirstLastCheckBox.Enabled = False
    End If
        
End Sub

Private Sub LastValueCheckBox_AfterUpdate()

    If FirstValueCheckBox.Value = True Or LastValueCheckBox = True Then
        AlignFirstLastCheckBox.Enabled = True
    Else
        AlignFirstLastCheckBox.Enabled = False
    End If
        
End Sub

Private Sub MaxValueCheckBox_AfterUpdate()

    If MaxValueCheckBox.Value = True Or MinValueCheckBox.Value = True Then
        ApplyBySeriesCheckBox.Enabled = True
    Else
        ApplyBySeriesCheckBox.Enabled = False
    End If
        
End Sub

Private Sub MinValueCheckBox_AfterUpdate()

    If MaxValueCheckBox.Value = True Or MinValueCheckBox.Value = True Then
        ApplyBySeriesCheckBox.Enabled = True
    Else
        ApplyBySeriesCheckBox.Enabled = False
    End If
        
End Sub


Private Sub CommandButton1_Click()

    For i = 0 To ListBox1.ListCount - 1
        
        present = False
        
        If ListBox2.ListCount <> 0 Then
            
            'Identify if the item selected already exists in ListBox2
            For j = 0 To ListBox2.ListCount - 1
                If ListBox1.List(i) = ListBox2.List(j) Then present = True
            Next
            
        End If
                       
        If ListBox1.Selected(i) = True And present = False Then ListBox2.AddItem ListBox1.List(i)
        
        CommandButton2.Enabled = True
        
   Next

End Sub

Private Sub CommandButton2_Click()

    ListBox1.ListIndex = -1
    ListBox2.Clear
    CommandButton2.Enabled = False

End Sub

Public Sub RunWithoutSave_Click()

    On Error Resume Next

    If PrimaryShape.Chart.HasLegend = True Then HadLegend = True
    
    PrimaryShape.Chart.ApplyDataLabels xlDataLabelsShowNone

    For Each mysrs In PrimaryShape.Chart.SeriesCollection
    
        For i = 0 To ListBox2.ListCount - 1

        With mysrs
        
          If ListBox2.List(i) = mysrs.Name Then
            
            vYvals = .Values
            vXVals = .XValues
            
            max_value = vYvals(1)
            min_value = vYvals(1)
                       
            For j = LBound(vYvals) To UBound(vYvals)
                If max_value < vYvals(j) Then max_value = vYvals(j)
                If min_value > vYvals(j) Then min_value = vYvals(j)
            Next
            
            If IsNumeric(FillEveryNValuesTextBox.Text) Then
                fill_every = Int(FillEveryNValuesTextBox.Text)
            Else
                fill_every = 2
            End If
            
            If IsNumeric(FillEveryNValuesTextBoxOffset.Text) Then
                fill_every_offset = Int(FillEveryNValuesTextBoxOffset.Text)
            Else
                fill_every_offset = 0
            End If
            
            If .Points.Count Mod 2 = 0 Then
                median_point = .Points.Count / 2
            Else
                median_point = (.Points.Count + 1) / 2
            End If

            For ipts = .Points.Count To 1 Step -1
              
              If Not IsEmpty(vYvals(ipts)) And Not IsError(vYvals(ipts)) _
                  And Not IsEmpty(vXVals(ipts)) And Not IsError(vXVals(ipts)) Then
                
                If ((FirstValueCheckBox.Value = True And ipts = 1) Or _
                   (LastValueCheckBox.Value = True And ipts = .Points.Count) Or _
                   (MaxValueCheckBox.Value = True And vYvals(ipts) = max_value) Or _
                   (MinValueCheckBox.Value = True And vYvals(ipts) = min_value) Or _
                   (FillEveryNValuesCheckbox.Value = True And _
                    (ipts + (fill_every - 1) - fill_every_offset) Mod fill_every = 0) Or _
                   (MedianValueCheckBox.Value = True And ipts = median_point)) Then

                    mysrs.Points(ipts).ApplyDataLabels _
                        ShowValue:=True
                        
                    If FirstValueCheckBox.Value = True And ipts = 1 And AlignFirstLastCheckBox.Value = True Then
                        mysrs.Points(ipts).DataLabel.Position = xlLabelPositionLeft
                    ElseIf LastValueCheckBox.Value = True And ipts = .Points.Count And AlignFirstLastCheckBox.Value = True Then
                        mysrs.Points(ipts).DataLabel.Position = xlLabelPositionRight
                    Else
                        mysrs.Points(ipts).DataLabel.Position = xlLabelPositionAbove
                    End If
                    
                               
                Else
        
                    mysrs.Points(ipts).ApplyDataLabels _
                        ShowValue:=False
               
                End If
              
              End If
            
            Next
                 
          End If
                
        End With
          
        
        Next
    Next
    If HadLegend = False Then PrimaryShape.Chart.HasLegend = False
    
end_func:
    Exit Sub

End Sub

Private Sub RunWithSave_Click()

    RunWithoutSave_Click
    Me.Hide

End Sub
Private Sub ToggleSelectAll_Click()

    If ToggleSelectAll.Value = True Then
        
        CommandButton1.Enabled = False
        
        ListBox2.Clear
        
        For i = 0 To ListBox1.ListCount - 1
              
            ListBox2.AddItem ListBox1.List(i)
            
        Next
        
        ToggleSelectAll.Caption = "Click to remove"
            
    Else
        
        CommandButton1.Enabled = True
        ListBox2.Clear
        ToggleSelectAll.Caption = "Add all items"
    
    End If

End Sub

Public Sub UserForm_Activate()

    Call SystemButtonSettings(Me, False)
    
    ListBox1.Clear

    For Each srs In PrimaryShape.Chart.SeriesCollection
        ListBox1.AddItem srs.Name
    Next
    
    If ToggleSelectAll.Value = True Then
        
        ListBox2.Clear
    
        For i = 0 To ListBox1.ListCount - 1
              
            ListBox2.AddItem ListBox1.List(i)
            CommandButton2.Enabled = True
            
        Next
    
    End If

End Sub


