Public Sub Display()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
T_NOTAM.AutoFilterMode = False
Dim i_1 As Integer
Dim i_2 As Integer
Dim rng_display As Range
Dim str_summary As String

T_EDIT.UsedRange.Rows.ClearContents
T_EDIT.Range("F1").Value = "REMARKS"
T_EDIT.Range("F1").Font.Bold = True
i_2 = 2
With T_NOTAM

  For i_1 = 2 To .UsedRange.Rows.Count
  'Go through all NOTAMS and select the "Not checked" NOTAMS
  
    If .Range("V" & i_1) = "Not checked" And .Range("K" & i_1) < Now + 2 And .Range("L" & i_1) > Now Then
      Set rng_display = .Range("A" & i_1, "Z" & i_1)
      'Enter the NOTAM ID
      T_EDIT.Range("A" & i_2) = rng_display.Range("Q1")
      'Create the summary string
      str_summary = rng_display.Range("F1") & " /" & rng_display.Range("G1") & " /" & _
      rng_display.Range("H1") & " /" & rng_display.Range("I1")
      
      T_EDIT.Range("B" & i_2) = str_summary
      T_EDIT.Range("B" & i_2).Font.Bold = True
      'Paste NOTAM Timeframe
      T_EDIT.Range("C" & i_2 + 1) = rng_display.Range("K1")
      T_EDIT.Range("D" & i_2 + 1) = rng_display.Range("L1")
      'Paste NOTAM BODY
      T_EDIT.Range("B" & i_2 + 1) = rng_display.Range("M1")
      T_EDIT.Range("B" & i_2 + 1).WrapText = True
      
      'Paste Remark
      T_EDIT.Range("F" & i_2 + 1) = rng_display.Range("Y1")
      
      T_EDIT.Range("E" & i_2 + 1) = rng_display.Range("V1")
      With T_EDIT.Range("E" & i_2 + 1).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
         Operator:=xlBetween, Formula1:="Checked OK,NOTAM excluded,AIP SUP uploaded,AIP via JEPPESEN,OCC Briefingsheet,Alternate Removed,Other Action taken,Not checked"
        .InCellDropdown = True
      End With
      
      i_2 = i_2 + 3
      
    End If
      
    
  Next i_1
  
End With
T_EDIT.Range("C1").Value = (i_2 - 2) / 3 & " NOTAMS to Check"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
T_EDIT.Activate
T_EDIT.Range("A2").Activate
ActiveWindow.DisplayGridlines = True
End Sub
