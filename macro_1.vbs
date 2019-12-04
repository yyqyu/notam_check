Sub Makro1()
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


With T_Config

  For i_1 = 1 To .UsedRange.Rows.Count
     
    str_Q = .Range("B" & i_1).Value
    str_Q = Trim(str_Q)
    dict_Q.Add str_Q, 1
    
  Next i_1
  
End With

For i_1 = 2 To T_NOTAM.UsedRange.Rows.Count

With T_NOTAM
'Filter by QCodes
  If dict_Q.Exists(.Range("D" & i_1).Text) And .Range("V" & i_1) = "Not checked" Then
    .Range("V" & i_1) = "Supressed QCODE"
    .Range("Z" & i_1) = str_user
    .Range("AA" & i_1) = Now
    End If
End With

Next i_1

Debug.Print "test"
    
End Sub
