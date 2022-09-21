# Userform (login + password) + Offset function

### Task
- Create userform which contains login and password for different users. Each user has unique sheet with data which are implemented with offset function.
### Code
- VBA code below:

````
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdClear_Click()

Me.txtUserID.Value = ""
Me.txtPassword.Value = ""
Me.txtUserID.SetFocus
End Sub

Private Sub cmdLogin_Click()

Dim user As String
Dim password As String

user = Me.txtUserID.Value
password = Me.txtPassword.Value

If (user = "admin" And password = "admin") Then
Unload Me
Application.Visible = True
Application.ScreenUpdating = False
Worksheets("Sheet1").Visible = True
Worksheets("Sheet2").Visible = True
Else
MsgBox "Invalid login credentials, Please try again", vbOKOnly + vbCritical, "Error during login phase"
End If

Private Sub Workbook_BeforeClose(Cancel As Boolean)
Application.ScreenUpdating = False
Worksheets("Sheet1").Visible = xlVeryHidden
Worksheets("Sheet2").Visible = xlVeryHidden
ThisWorkbook.Save  
End Sub
````
### Screenshoots
- Userform (login+password - VBA) and offset function [[Excel file here.]](https://github.com/Ciachula/Portfolio/tree/main/excel)
<img width="854" alt="userform+offset1" src="https://user-images.githubusercontent.com/31890259/187172384-016f4a0f-179d-4783-bdf5-b6e602626db0.PNG">
