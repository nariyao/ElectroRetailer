Attribute VB_Name = "mod_employee"
'DISPLAY EMPLOYEE DETAILS
Public Function displayEmp()
search.Show
search.SetFocus
search.searchStatus = "emp"
search.search_box.ToolTipText = "Enter Employee id/name"
search.search_box.Text = ""
main.Caption = mainCaption("Employee Details")
setEmpHeader
smfgEmp ("a")
End Function

'EMPLOYEE SEARCH
Public Function smfgEmp(ByVal id As String)
search.smfg.Rows = 1
connectOracle
If (id = "a") Then
sql = "SELECT * FROM ER_MASTER_EMPLOYEE"
Else
sql = "SELECT * FROM ER_MASTER_EMPLOYEE WHERE E_ID LIKE '%" + id + "%' OR E_NAME LIKE '%" + id + "%'"
End If
Set R = C.Execute(sql)
Do While Not R.EOF
search.smfg.Rows = search.smfg.Rows + 1
For i = 0 To 14
search.smfg.TextMatrix(search.smfg.Rows - 1, i) = R.Fields(i) & ""
Next i
R.MoveNext
Loop
C.Close
End Function

'set values in table head
Private Function setEmpHeader()
search.smfg.Rows = 1
search.smfg.Cols = 15
search.smfg.TextMatrix(0, 0) = " Employee ID"
search.smfg.TextMatrix(0, 1) = " Name"
search.smfg.TextMatrix(0, 2) = " Father's Name"
search.smfg.TextMatrix(0, 3) = " DOB"
search.smfg.TextMatrix(0, 4) = " Gender"
search.smfg.TextMatrix(0, 5) = " Moblie"
search.smfg.TextMatrix(0, 6) = " Email"
search.smfg.TextMatrix(0, 7) = " Adhaar no."
search.smfg.TextMatrix(0, 8) = " DOJ"
search.smfg.TextMatrix(0, 9) = " Address"
search.smfg.TextMatrix(0, 10) = " Qualification"
search.smfg.TextMatrix(0, 11) = " Experence"
search.smfg.TextMatrix(0, 12) = " Post"
search.smfg.TextMatrix(0, 13) = " Leave"
search.smfg.TextMatrix(0, 14) = " Salary"
search.smfg.ColWidth(0) = 1500
search.smfg.ColWidth(1) = 3000
search.smfg.ColWidth(2) = 3000
search.smfg.ColWidth(3) = 1300
search.smfg.ColWidth(4) = 1000
search.smfg.ColWidth(5) = 1500
search.smfg.ColWidth(6) = 3000
search.smfg.ColWidth(7) = 2100
search.smfg.ColWidth(8) = 1300
search.smfg.ColWidth(9) = 5000
search.smfg.ColWidth(10) = 1800
search.smfg.ColWidth(11) = 1500
search.smfg.ColWidth(12) = 2000
search.smfg.ColWidth(13) = 1000
search.smfg.ColWidth(14) = 1000
End Function


'Insert employee form
Public Function insEmp()
emp.Show
emp.SetFocus
emp.add_btn.Enabled = True
emp.add_btn.Visible = True
emp.update_btn.Enabled = False
emp.update_btn.Visible = False
emp.delete_btn.Enabled = False
emp.delete_btn.Visible = False
emp.id.Text = generateEmpId
End Function

'this will generate new employee id
Public Function generateEmpId() As String
Dim a As Integer
Dim b As String
connectOracle
Set R = New ADODB.Recordset
Set R = C.Execute("SELECT E_ID FROM ER_MASTER_EMPLOYEE")
If IsNull(R.Fields(0)) Then
a = 0
Else
Do While Not R.EOF
a = Val(Mid(R.Fields(0), 8, 3))
R.MoveNext
Loop
End If
If a < 9 Then
b = "00" + Trim(Str(a + 1))
ElseIf a < 99 Then
b = "0" + Trim(Str(a + 1))
Else
b = Str(a + 1)
End If
generateEmpId = "EMP" + Trim(Str(Year(Date))) + Trim(b)
R.Close
C.Close
End Function

'Update And Delete Function
Public Function empUpDel(ByVal id As String)
connectOracle
Set R = C.Execute("SELECT * FROM ER_MASTER_EMPLOYEE WHERE E_ID='" + id + "'")
temp = setEmpValue(R.Fields(0), R.Fields(1), R.Fields(2), R.Fields(3), R.Fields(4), R.Fields(5), R.Fields(6), R.Fields(7), R.Fields(8), R.Fields(9), R.Fields(10), R.Fields(11), R.Fields(12), R.Fields(13), R.Fields(14))
emp.Show
emp.SetFocus
emp.add_btn.Enabled = False
emp.add_btn.Visible = False
emp.update_btn.Enabled = True
emp.update_btn.Visible = True
emp.delete_btn.Enabled = True
emp.delete_btn.Visible = True
End Function

'Set value in employee textbox
Public Function setEmpValue(ByVal emp_id As String, ByVal emp_name As String, ByVal emp_fname As String, ByVal dob As String, ByVal emp_gen As String, ByVal emp_mob As String, ByVal emp_email As String, ByVal emp_adhaar As String, ByVal emp_doj As String, ByVal emp_addr As String, ByVal emp_quali As String, ByVal emp_expr As String, ByVal emp_post As String, ByVal emp_leave As Integer, ByVal emp_salary As Double)
emp.id.Text = emp_id
emp.emp_name.Text = emp_name
emp.fname.Text = emp_fname
emp.doj.Value = emp_doj
emp.dob.Value = emp_dob
setGender (emp_gen)
emp.mobile.Text = emp_mob
emp.email.Text = emp_email
emp.adhaar.Text = emp_adhaar
emp.addr.Text = emp_addr
emp.quali.Text = emp_quali
emp.expr.Text = emp_expr
emp.post.Text = emp_post
emp.leaves.Text = emp_leave
emp.salary.Text = Str(emp_salary)
End Function

'get gender
Public Function getGender() As String
If emp.M.Value = True Then
getGender = "M"
ElseIf emp.F.Value = True Then
getGender = "F"
ElseIf emp.T.Value = True Then
getGender = "T"
End If
End Function

'Set Gender
Public Function setGender(ByVal g As String)
Select Case g
Case "M": emp.M.Value = True
Case "F": emp.F.Value = True
Case "T": emp.T.Value = True
End Select
End Function
