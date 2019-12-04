
Public Function NOTAM_SortNew(Q As Worksheet) As Integer
Dim i_1 As Integer
Dim dict_Q
Set dict_Q = CreateObject("Scripting.Dictionary")

Dim dict_NOTAM
Set dict_NOTAM = CreateObject("Scripting.Dictionary")

Dim dict_AD
Set dict_AD = CreateObject("Scripting.Dictionary")

Dim str_user As String
Dim rng_temp As Range
Dim i_row_NOTAM As Integer
Dim str_AD As String
Dim str_Q As String
Dim int_counter As Integer

T_NOTAM.AutoFilterMode = False


str_user = Application.UserName

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

'Create a dictionary with all QCodes in T_Config
With T_Config

  For i_1 = 1 To .UsedRange.Rows.Count
     
    str_Q = .Range("B" & i_1).Value
    str_Q = Trim(str_Q)
    dict_Q.Add str_Q, 1
    
  Next i_1
  
End With

'Create a dictionary with all NOTAMS in T_NOTAM
With T_NOTAM

  For i_1 = 2 To .UsedRange.Rows.Count
  
    If Not dict_NOTAM.Exists(.Range("Q" & i_1).Value) Then
    
        dict_NOTAM.Add .Range("Q" & i_1).Value, Range("Q" & i_1).Value
    End If
    
  Next i_1
  
End With

int_counter = 0
'Go Trouhg all NOTMAS
With Q

i_row_NOTAM = T_NOTAM.UsedRange.Rows.Count + 1
For i_1 = 2 To .UsedRange.Rows.Count
'Exclude if Notam is allready in T_NOTAM
  If dict_NOTAM.Exists(.Range("Q" & i_1).Text) Then
    GoTo jump1
  End If
'Exclude if Notam is for a not planned aerodrome
  If Not dict_AD.Exists(.Range("N" & i_1).Text) Then
    GoTo jump1
  End If


  
  'Filter by QCodes
  If dict_Q.Exists(.Range("D" & i_1).Text) Then
    .Range("V" & i_1) = "Supressed QCODE"
    .Range("W" & i_1) = str_user
    '.Range("K" & i_1) = DateValue(Left(.Range("K" & i_1), 10)) + TimeValue(Mid(.Range("K" & i_1), 12, 8))
    .Range("A" & i_1, "Z" & i_1).Copy
    T_NOTAM.Range("A" & i_row_NOTAM).PasteSpecial (xlPasteValuesAndNumberFormats)
    T_NOTAM.Range("K" & i_1) = DateValue(Left(.Range("K" & i_1), 10)) + TimeValue(Mid(.Range("K" & i_1), 12, 8))
    T_NOTAM.Range("L" & i_1) = DateValue(Left(.Range("L" & i_1), 10)) + TimeValue(Mid(.Range("L" & i_1), 12, 8))
    i_row_NOTAM = i_row_NOTAM + 1
    int_counter = int_counter + 1
  'If not allready in the list and planned and not filtered by QCode
  'Add the notam to the list with index Not checked
  Else
    .Range("V" & i_1) = "Not checked"
    .Range("W" & i_1) = str_user
    .Range("A" & i_1, "Z" & i_1).Copy
    T_NOTAM.Range("A" & i_row_NOTAM).PasteSpecial (xlPasteValuesAndNumberFormats)
    T_NOTAM.Range("K" & i_row_NOTAM) = DateValue(Left(T_NOTAM.Range("K" & i_row_NOTAM), 10)) + TimeValue(Mid(T_NOTAM.Range("K" & i_row_NOTAM), 12, 8))
    T_NOTAM.Range("L" & i_row_NOTAM) = DateValue(Left(T_NOTAM.Range("L" & i_row_NOTAM), 10)) + TimeValue(Mid(T_NOTAM.Range("L" & i_row_NOTAM), 12, 8))
    i_row_NOTAM = i_row_NOTAM + 1
    int_counter = int_counter + 1
  End If
  
jump1:
Next i_1

End With
    

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
NOTAM_SortNew = int_counter

End Function
