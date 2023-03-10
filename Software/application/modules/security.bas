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

'CHECK PRIVILEGE
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

'Logout fction
Public Sub logout()
userid = blank
password = blank
Unload main
login.Show
End Sub

'CREATE USER
Public Function createUser(ByVal userid As String, ByVal pass As String, ByVal acc_lev As String) As Integer
connectOracle
On Error GoTo ER_CU
C.Execute ("INSERT INTO ER_MASTER_LOGIN VALES('" + userid + "','" + pass + "','" + acc_lev + "')")
createUser = 1
Exit Sub
ER_CU:
MsgBox Err.Description
createUser = -1
End Function

'UPDATE USER
Public Function updateUser(ByVal userid As String, ByVal pass As String, ByVal acc_lev As String) As Integer
connectOracle
On Error GoTo ER_UU
C.Execute ("UPDATE ER_MASTER_LOGIN SET PASSWORD='" + pass + "', ACCESS_LEVEL='" + acc_lev + "' WHERE USERID='" + userid + "'")
createUser = 1
Exit Sub
ER_UU:
MsgBox Err.Description
createUser = -1
End Function

'DELETE USER
Public Function deleteUser(ByVal userid As String) As Integer
connectOracle
On Error GoTo ER_UU
C.Execute ("delete ER_MASTER_LOGIN WHERE USERID='" + userid + "'")
createUser = 1
Exit Function
ER_UU:
MsgBox Err.Description
createUser = -1
End Function

'DISPLAY USER
Public Function displayUser(ByVal id As String) As Integer
smfg.Rows = 1
connectOracle
If (id = "a") Then
sql = "SELECT * FROM ER_MASTER_LOGIN"
Else
sql = "SELECT * FROM ER_MASTER_LOGIN WHERE USERID LIKE '%" + id + "%' OR ACCESS_LEVEL LIKE '%" + id + "%'"
End If
Set R = C.Execute(sql)
Do While Not R.EOF
smfg.Rows = smfg.Rows + 1
For i = 0 To 2
smfg.TextMatrix(smfg.Rows - 1, i) = R.Fields(i) & ""
Next i
R.MoveNext
Loop
C.Close
End Function
