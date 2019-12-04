
Public Sub Show_Relevant()
Dim i_1 As Integer
Dim i_2 As Integer
Dim i_3 As Integer

Dim i_row_NOTAM As Integer

Dim dict_AD
Set dict_AD = CreateObject("Scripting.Dictionary")

Dim str_user As String
Dim rng_temp As Range
Dim str_AD As String
Dim str_Q As String
Dim int_counter As Integer
Dim dict_cat
Set dict_cat = CreateObject("Scripting.Dictionary")
Dim str_test As String


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

T_NOTAM.AutoFilterMode = False
With T_Relevant


    .AutoFilterMode = False
    .UsedRange.Rows.ClearContents
    .Range("A2") = "NOTAM ID"
    .Range("C2") = "Valid From"
    .Range("D2") = "Valid To"
    .Range("E2") = "Airport"
    .Range("F2") = "Remarks"
    .Range("A1") = "Relevant NOTAMS for all planned Aerodromes for the next 48 HOURS"
    .Range("G2") = "Action"
End With


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


'Go Trouhg all NOTMAS
With T_NOTAM

i_3 = 3

i_row_NOTAM = T_NOTAM.UsedRange.Rows.Count + 1
For i_2 = 2 To i_row_NOTAM
str_test = .Range("Q" & i_2)

'And Not .Range("V" & i_2) = "Checked OK"

'l
  If dict_AD.Exists(.Range("N" & i_2).Text) And .Range("K" & i_2) < Now + 2 And .Range("L" & i_2) > Now And _
  dict_cat.Exists(.Range("V" & i_2).Text) Then
      Set rng_display = .Range("A" & i_2, "AZ" & i_2)
      'Enter the NOTAM ID
      T_Relevant.Range("A" & i_3) = rng_display.Range("Q1")
      'Create the summary string
      str_summary = rng_display.Range("F1") & " /" & rng_display.Range("G1") & " /" & _
      rng_display.Range("H1") & " /" & rng_display.Range("I1")
      
      T_Relevant.Range("B" & i_3) = str_summary
      T_Relevant.Range("B" & i_3).Font.Bold = True
      'Paste NOTAM Timeframe
      T_Relevant.Range("C" & i_3 + 1) = rng_display.Range("K1")
      T_Relevant.Range("D" & i_3 + 1) = rng_display.Range("L1")
      'Paste NOTAM BODY
      T_Relevant.Range("B" & i_3 + 1) = rng_display.Range("M1")
      T_Relevant.Range("B" & i_3 + 1).WrapText = True
      T_Relevant.Range("B" & i_3 + 1).Font.Bold = False
      
      'Paste Remark
      T_Relevant.Range("F" & i_3 + 1) = rng_display.Range("Y1")
      
      'Paste Airport
      T_Relevant.Range("E" & i_3 + 1) = rng_display.Range("N1")
      
      'Paste Actioin
      T_Relevant.Range("G" & i_3 + 1) = rng_display.Range("V1")
      T_Relevant.Range("G" & i_3 + 1).Font.Bold = False
      
      
      i_3 = i_3 + 2
      
 Else
    GoTo jump1
  End If
  
jump1:
Next i_2
End With
T_Relevant.Range("A2", "G" & i_3).AutoFilter

T_Relevant.Activate
T_Relevant.Range("A3").Select

End Sub
