Public Function Airport_Alert()
Dim planned_AD
Dim AD As Variant
Dim i_used_config As Integer
Dim i_1 As Integer
Dim check As Boolean
Dim i_2 As Integer
Dim str_AD As String

Dim dict_AD
Set dict_AD = CreateObject("Scripting.Dictionary")


'Create a dictionary with all Airports in T_Config
With T_MENU

  For i_1 = 6 To .UsedRange.Rows.Count
    If dict_AD.Exists(Left(.Range("K" & i_1).Text, 4)) Then
    Else
      str_AD = .Range("K" & i_1).Text
      str_AD = Left(str_AD, 4)
      dict_AD.Add str_AD, 1
    End If
    
    If dict_AD.Exists(Left(.Range("N" & i_1).Text, 4)) Then
    Else
      str_AD = .Range("N" & i_1).Text
      str_AD = Left(str_AD, 4)
      dict_AD.Add str_AD, 1
    End If
    
    If dict_AD.Exists(Left(.Range("O" & i_1).Text, 4)) Then
    Else
      str_AD = .Range("O" & i_1).Text
      str_AD = Left(str_AD, 4)
      dict_AD.Add str_AD, 1
    End If
    
    If dict_AD.Exists(Left(.Range("P" & i_1).Text, 4)) Then
    Else
      str_AD = .Range("P" & i_1).Text
      str_AD = Left(str_AD, 4)
      dict_AD.Add str_AD, 1
    End If
    
    
  Next i_1
  
End With

i_used_config = T_Config.UsedRange.Rows.Count
planned_AD = dict_AD.Keys()

For Each AD In planned_AD
    check = False
    For i_2 = 1 To i_used_config
    
        If 0 = StrComp(AD, T_Config.Range("C" & i_2), vbTextCompare) Or AD = "" Then
            check = True
            GoTo jump5
            
        
        End If
    
    Next i_2
    


    If check = False Then
        MsgBox ("The Aerodrome " & AD & " is not coverd by this Application")
    End If
jump5:
Next AD

End Function
