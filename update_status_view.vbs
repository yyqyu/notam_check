Public Sub Update_Status_View()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
T_NOTAM.AutoFilterMode = False
Dim i_1 As Integer
Dim i_2 As Integer
Dim str_id As String
Dim str_oaw_cat As String
Dim int_counter As Integer
Dim int_counter2 As Integer
Dim str_remark As String

For i_1 = 1 To T_view.UsedRange.Rows.Count

  If Not IsEmpty(T_view.Range("A" & i_1)) Then
    str_id = T_view.Range("A" & i_1)
    str_oaw_cat = T_view.Range("G" & i_1)
    str_remark = T_view.Range("F" & i_1)
    
  Else
    GoTo jump2
  End If
  
  For i_2 = 2 To T_NOTAM.UsedRange.Rows.Count
    If str_id = T_NOTAM.Range("Q" & i_2) Then
      T_NOTAM.Range("V" & i_2) = str_oaw_cat
      T_NOTAM.Range("Z" & i_2) = Application.UserName
      T_NOTAM.Range("AA" & i_2) = Now
      T_NOTAM.Range("Y" & i_2) = str_remark
        
    End If
  Next i_2
jump2:
Next i_1



Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
ActiveWindow.DisplayGridlines = False
T_MENU.Activate
ActiveWorkbook.Save
End Sub
