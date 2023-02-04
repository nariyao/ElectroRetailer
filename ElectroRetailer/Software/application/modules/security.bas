Attribute VB_Name = "mod_security"
Public C As ADODB.Connection
Public R As ADODB.Recordset
Public userid As String
Public password As String
Public sql As String
Dim dump As String

'For Databse Connection
Public Function connectOracle()
Set C = New ADODB.Connection
C.Open "Provider=MSDAORA.1;User ID=electro/retailer;Persist Security Info=False"
Set R = New ADODB.Recordset
End Function

'Login Function
Public Function authoCheck(ByRef loginID As String, ByRef loginPW As String) As Integer
sql = "SELECT * FROM ER_MASTER_LOGIN WHERE USERID='" + loginID + " ' AND PASSWORD='" + loginPW + "'"
connectOracle
Set R = C.Execute(sql)
On Error GoTo er_login
loginPrivilege (Trim(R.Fields(2)))
C.Close
authoCheck = 1
Exit Function
er_login:
If (Err.Number = 3021) Then
authoCheck = 0
Else
authoCheck = -1
End If

End Function

Public Sub loginPrivilege(ByVal lp As String)
If lp = "SU" Then
main.menuReport.Visible = True
main.menuReport.Enabled = True
main.menuAdmin.Visible = True
main.menuAdmin.Enabled = True
Else
main.menuReport.Visible = False
main.menuReport.Enabled = False
main.menuAdmin.Visible = False
main.menuAdmin.Enabled = False
End If
End Sub

'Logout function
Public Sub logout()
userid = blank
password = blank
Unload main
login.Show
End Sub


