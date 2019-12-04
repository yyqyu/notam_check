Public Sub Update_Status()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
T_NOTAM.AutoFilterMode = False
Dim i_1 As Integer
Dim i_2 As Integer
Dim i_3 As Integer
Dim str_id As String
Dim str_oaw_cat As String
Dim int_counter As Integer
Dim int_counter2 As Integer
Dim str_remark As String

int_counter = 0
int_counter2 = 0
For i_1 = 2 To T_EDIT.UsedRange.Rows.Count

  If Not IsEmpty(T_EDIT.Range("A" & i_1)) Then
    str_id = T_EDIT.Range("A" & i_1)
    str_oaw_cat = T_EDIT.Range("E" & i_1 + 1)
    str_remark = T_EDIT.Range("F" & i_1 + 1)
  Else
    GoTo jump2
  End If
  
  For i_2 = 2 To T_NOTAM.UsedRange.Rows.Count
    If str_id = T_NOTAM.Range("Q" & i_2) Then
      T_NOTAM.Range("V" & i_2) = str_oaw_cat
      T_NOTAM.Range("W" & i_2) = Application.UserName
      T_NOTAM.Range("X" & i_2) = Now
      T_NOTAM.Range("Y" & i_2) = str_remark
      
        If Not str_oaw_cat = "Not checked" Then
            int_counter = int_counter + 1
        End If
        
      

      GoTo jump2
    End If
  Next i_2
  
jump2:

Next i_1



Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
T_MENU.Activate

With T_NOTAM

  For i_3 = 2 To .UsedRange.Rows.Count
  
    If .Range("V" & i_3) = "Not checked" Then
        int_counter2 = int_counter2 + 1
    End If
    
  Next i_3
  
End With


'MsgBox (int_counter & " NOTAMS have been updated " & int_counter2 & " NOTMAS have to be checked")
MsgBox (int_counter & " checked NOTAMS have been saved")

ActiveWorkbook.Save
ActiveWindow.DisplayGridlines = False
T_MENU.Activate

End Sub
