Public Sub Queries_List()
Dim int_counter As Integer
Dim bol_stop As Boolean
Dim oldStatusBar
Dim i_59 As Integer

bol_stop = False
bol_stop = Data_Warning()

If bol_stop = True Then GoTo stop1

On Error GoTo ErrorHandling


'Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

MsgBox ("This update may take upto 8 minutes")

Call Airport_Alert

'ThisWorkbook.RefreshAll


oldStatusBar = Application.DisplayStatusBar
Application.DisplayStatusBar = True
Application.StatusBar = "Retrieving NOTAM DATA "

Sheet9.Range("A1").ListObject.QueryTable.Refresh
Application.StatusBar = " 1%"
Do While Sheet9.Range("A1").ListObject.QueryTable.Refreshing = True
 DoEvents
Loop
Sheet8.Range("A1").ListObject.QueryTable.Refresh
Application.StatusBar = "2%"
Do While Sheet8.Range("A1").ListObject.QueryTable.Refreshing = True
 DoEvents
Loop
Application.StatusBar = "6%"
Sheet7.Range("A1").ListObject.QueryTable.Refresh
Do While Sheet7.Range("A1").ListObject.QueryTable.Refreshing = True
 DoEvents
Loop
Application.StatusBar = "12%"
Sheet6.Range("A1").ListObject.QueryTable.Refresh
Do While Sheet6.Range("A1").ListObject.QueryTable.Refreshing = True
 DoEvents
Loop
Application.StatusBar = "17%"
Sheet5.Range("A1").ListObject.QueryTable.Refresh
Do While Sheet5.Range("A1").ListObject.QueryTable.Refreshing = True
 DoEvents
Loop
Application.StatusBar = "23%"
Sheet3.Range("A1").ListObject.QueryTable.Refresh
Do While Sheet3.Range("A1").ListObject.QueryTable.Refreshing = True
 DoEvents
Loop
Application.StatusBar = "29%"
Sheet2.Range("A1").ListObject.QueryTable.Refresh
Do While Sheet2.Range("A1").ListObject.QueryTable.Refreshing = True
 DoEvents
Loop
Application.StatusBar = "35%"
Sheet19.Range("A1").ListObject.QueryTable.Refresh
Do While Sheet19.Range("A1").ListObject.QueryTable.Refreshing = True
 DoEvents
Loop
Application.StatusBar = "41%"
Sheet18.Range("A1").ListObject.QueryTable.Refresh
Do While Sheet18.Range("A1").ListObject.QueryTable.Refreshing = True
 DoEvents
Loop
Application.StatusBar = "47%"
Sheet17.Range("A1").ListObject.QueryTable.Refresh
Do While Sheet17.Range("A1").ListObject.QueryTable.Refreshing = True
 DoEvents
Loop
Application.StatusBar = "53%"
Sheet16.Range("A1").ListObject.QueryTable.Refresh
Do While Sheet16.Range("A1").ListObject.QueryTable.Refreshing = True
 DoEvents
Loop
Application.StatusBar = "69%"
Sheet15.Range("A1").ListObject.QueryTable.Refresh
Do While Sheet15.Range("A1").ListObject.QueryTable.Refreshing = True
 DoEvents
Loop
Application.StatusBar = "65%"
Sheet14.Range("A1").ListObject.QueryTable.Refresh
Do While Sheet14.Range("A1").ListObject.QueryTable.Refreshing = True
 DoEvents
Loop
Application.StatusBar = "71%"
Sheet13.Range("A1").ListObject.QueryTable.Refresh
Do While Sheet13.Range("A1").ListObject.QueryTable.Refreshing = True
 DoEvents
Loop
Application.StatusBar = "77%"
Sheet12.Range("A1").ListObject.QueryTable.Refresh
Do While Sheet12.Range("A1").ListObject.QueryTable.Refreshing = True
 DoEvents
Loop
Application.StatusBar = "83%"
Sheet11.Range("A1").ListObject.QueryTable.Refresh
Do While Sheet11.Range("A1").ListObject.QueryTable.Refreshing = True
 DoEvents
Loop
Application.StatusBar = "89%"
Sheet10.Range("A1").ListObject.QueryTable.Refresh
Do While Sheet10.Range("A1").ListObject.QueryTable.Refreshing = True
 DoEvents
Loop
Application.StatusBar = "Retrieving NOTAM DATA 100%"

stop1:
int_counter = int_counter + NOTAM_SortNew(Sheet2)
int_counter = int_counter + NOTAM_SortNew(Sheet3)
int_counter = int_counter + NOTAM_SortNew(Sheet5)
int_counter = int_counter + NOTAM_SortNew(Sheet6)
int_counter = int_counter + NOTAM_SortNew(Sheet7)
int_counter = int_counter + NOTAM_SortNew(Sheet8)
int_counter = int_counter + NOTAM_SortNew(Sheet9)
int_counter = int_counter + NOTAM_SortNew(Sheet10)
int_counter = int_counter + NOTAM_SortNew(Sheet11)
int_counter = int_counter + NOTAM_SortNew(Sheet12)
int_counter = int_counter + NOTAM_SortNew(Sheet13)
int_counter = int_counter + NOTAM_SortNew(Sheet14)
int_counter = int_counter + NOTAM_SortNew(Sheet15)
int_counter = int_counter + NOTAM_SortNew(Sheet16)
int_counter = int_counter + NOTAM_SortNew(Sheet17)
int_counter = int_counter + NOTAM_SortNew(Sheet18)
int_counter = int_counter + NOTAM_SortNew(Sheet19)



Application.StatusBar = False
Application.DisplayStatusBar = oldStatusBar


MsgBox (CStr(int_counter) & " NOTAMS have been added to the database")
T_MENU.Range("C28") = Now

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
T_MENU.Activate
ActiveWorkbook.Save
On Error GoTo -1
Exit Sub
ErrorHandling:
    MsgBox ("Refresh Failed, TRY AGAIN in 20 minutes")
    i_59 = T_Error.UsedRange.Rows.Count
    T_Error.Range("A" & i_59) = Err.Number
    T_Error.Range("B" & i_59) = Err.Description
    T_Error.Range("C" & i_59) = Environ$("computername")
    T_Error.Range("D" & i_59) = Now
    T_Error.Range("E" & i_59) = Application.UserName
    


End Sub
