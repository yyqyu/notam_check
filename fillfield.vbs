Public Function FillField(HTMLFields As Object, HTMLName As String, HTMLInput As String) As Boolean
Dim x1 As Long
Dim i_1 As Long
Dim currentField As Object
FillField = False


x1 = HTMLFields.Length - 1
For i_1 = 0 To x1
  Debug.Print HTMLFields(i_1).Name
  If HTMLFields(i_1).Name = HTMLName Then
    Set currentField = HTMLFields(i_1)
    currentField.Value = HTMLInput
    currentField.Focus
    FillField = True
    Exit For
  End If
Next i_1

End Function
