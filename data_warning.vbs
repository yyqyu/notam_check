Public Function Data_Warning()
Dim Decision
If T_MENU.Range("C28") > (Now - (3 / 24)) Then
 Decision = MsgBox("The NOTAM data has been refereshed less than 3 hours ago" & vbCrLf & _
 "the refreshing of data is NOT FOR FREE. Do you want to continue?", vbYesNo)
 
    If Decision = 6 Then
        Data_Warning = False
    Else
        Data_Warning = True
    End If
End If

End Function
