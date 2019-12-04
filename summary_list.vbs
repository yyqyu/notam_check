Public Sub Summary_List()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
T_NOTAM.AutoFilterMode = False
Dim i_1 As Integer
Dim i_2 As Integer
Dim i_used As Integer
Dim rng_display As Range
Dim str_summary As String
Dim dict_cat
Set dict_cat = CreateObject("Scripting.Dictionary")
Dim dat_today As Date

i_used = T_view.UsedRange.Rows.Count
T_view.Activate
T_view.Range("A2", "K" & i_used).ClearContents
T_view.AutoFilterMode = False
If T_MENU.Checked_OK_Box = True Then
    dict_cat.Add "Checked OK", 1
End If

If T_MENU.NOTAM_excluded_Box = True Then
    dict_cat.Add "NOTAM excluded", 1
End If

If T_MENU.AIP_SUP_uploaded_Box = True Then
    dict_cat.Add "AIP SUP uploaded", 1
End If

If T_MENU.AIP_SUP_Jeppesen_Box = True Then
    dict_cat.Add "AIP via JEPPESEN", 1
End If

If T_MENU.OCC_Briefingsheet_Box = True Then
    dict_cat.Add "OCC Briefingsheet", 1
End If

If T_MENU.Alternate_Removed_Box = True Then
    dict_cat.Add "Alternate Removed", 1
End If

If T_MENU.Other_Action_Box = True Then
    dict_cat.Add "Other Action taken", 1
End If

If T_MENU.QCODE_Filter_BOX = True Then
    dict_cat.Add "Supressed QCODE", 1
End If

i_2 = 2

With T_NOTAM

  For i_1 = 2 To .UsedRange.Rows.Count
  'Go through all NOTAMS and select the "Not checked" NOTAMS
  
    If dict_cat.Exists(.Range("V" & i_1).Text) Then
      Set rng_display = .Range("A" & i_1, "Z" & i_1)
      'Enter the NOTAM ID
      T_view.Range("A" & i_2) = rng_display.Range("Q1")
      
      'Paste NOTAM End
      T_view.Range("H" & i_2) = rng_display.Range("L1")
      'Paste NOTAM BODY
      T_view.Range("E" & i_2) = rng_display.Range("M1")
      T_view.Range("E" & i_2).WrapText = True
      'Paste Location
      T_view.Range("D" & i_2) = rng_display.Range("N1")
      'Paste Date of check
      T_view.Range("B" & i_2) = rng_display.Range("X1")
      T_view.Range("B" & i_2).NumberFormat = "DDMMM YY"
      'Paste user
      T_view.Range("C" & i_2) = rng_display.Range("W1")
      'Paste Remark
      T_view.Range("F" & i_2) = rng_display.Range("Y1")
      'Paste Action Taken
      T_view.Range("G" & i_2) = rng_display.Range("V1")
      
      With T_view.Range("G" & i_2).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
         Operator:=xlBetween, Formula1:="Checked OK,NOTAM excluded,AIP SUP uploaded,AIP via JEPPESEN,OCC Briefingsheet,Alternate Removed,Other Action taken,Not checked"
        .InCellDropdown = True
      End With
      
      i_2 = i_2 + 1
      
    End If
      
    
  Next i_1
  .Range("A2", "A" & i_2).NumberFormat = "TTMMM-JJJJ hh:mm"
  
End With
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
If i_2 > 2 Then
    With T_view
    dat_today = Now
    .AutoFilterMode = False
    .Range("A1", "H" & i_2 - 2).AutoFilter _
        Field:=8, _
        Criteria1:=">" & CDbl(Now), _
        Operator:=xlAnd
    .AutoFilter.ApplyFilter

End With
End If

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
ActiveWindow.DisplayGridlines = True
T_view.Activate
T_view.Range("A2").Activate
End Sub
