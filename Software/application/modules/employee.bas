Attribute VB_Name = "mod_employee"
Public Function displayEmp()
search.Show
search.SetFocus
main.Caption = mainCaption("Employee Details")
setEmpHeader
smfgEmp ("a")
End Function

Public Function smfgEmp(ByVal id As String)
search.smfg.Rows = 1
connectOracle
If (id = "a") Then
sql = "SELECT * FROM ER_MASTER_EMPLOYEE"
Else
sql = "SELECT * FROM ER_MASTER_EMPLOYEE WHERE S_ID= '" + id + "'"
End If
Set R = C.Execute(sql)
Do While Not R.EOF
search.smfg.Rows = search.smfg.Rows + 1
For i = 0 To 12
search.smfg.TextMatrix(search.smfg.Rows - 1, i) = R.Fields(i) & ""
Next i
R.MoveNext
Loop
C.Close
End Function

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
search.smfg.ColWidth(0) = 2500
search.smfg.ColWidth(1) = 3000
search.smfg.ColWidth(2) = 3000
search.smfg.ColWidth(3) = 1000
search.smfg.ColWidth(4) = 1000
search.smfg.ColWidth(5) = 2000
search.smfg.ColWidth(6) = 3000
search.smfg.ColWidth(7) = 3000
search.smfg.ColWidth(8) = 1000
search.smfg.ColWidth(9) = 7000
search.smfg.ColWidth(10) = 3000
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

Private Function generateEmpId() As String
Dim a As Integer
Dim b As String
connectOracle
Set R = New ADODB.Recordset
Set R = C.Execute("SELECT E_ID FROM ER_MASTER_EMPLOYEE")
R.MoveLast
a = Val(Mid(R.Fields(E_id), 7, 3))
If a < 10 Then
b = "00" + Str(a + 1)
ElseIf a < 100 Then
b = "0" + Str(a + 1)
ElseIf a < 100 Then
b = Str(a + 1)
End If
generateEmpId = "EMP" + Trim(b)
R.Close
C.Close
End Function
