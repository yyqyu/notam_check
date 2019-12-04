Public Sub Clear_Airports()
Dim i_usedRange As Integer
Dim wkb As Workbook
Dim InputFields As Object

Set wkb = ThisWorkbook

With T_MENU
    MsgBox ("Copy Airportlist from Chrome or Firefox" & vbCrLf & vbCrLf & "Login: " & vbCrLf & "Password:")
    i_usedRange = .UsedRange.Rows.Count
    .Range("G6", "P" & i_usedRange).ClearContents
    '.Range("G6", "P" & i_usedRange).ClearFormats
    .Range("G6").Value = "Paste Here"
    Shell ("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
    wkb.FollowHyperlink Address:="http://www.opscontrol.com/AirportSuitability/Next24Hours"
    MsgBox ("Copy Airportlist from Chrome or Firefox" & vbCrLf & vbCrLf & "Login: " & vbCrLf & "Password: ")
End With
End Sub
