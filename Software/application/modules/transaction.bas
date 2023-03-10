Attribute VB_Name = "mod_transaction"
Public Function displayTrans()
search.Show
search.SetFocus
search.searchStatus = "trans"
search.search_box.ToolTipText = "Enter transaction id/particular/CR/DR/mode/date"
search.search_box.Text = ""
main.Caption = mainCaption("Transaction Details")
setTransHeader
smfgTrans ("a")
End Function

'DISPLAY ALL TRANSCTION
Public Function smfgTrans(ByVal id As String)
search.smfg.Rows = 1
connectOracle
If (id = "a") Then
sql = "SELECT * FROM ER_MASTER_TRANSACTION ORDER BY T_ID DESC"
Else
sql = "SELECT * FROM ER_MASTER_TRANSACTION WHERE T_ID LIKE '%" + id + "%' OR T_DATE LIKE '%" + id + "%' OR T_PARTICULAR LIKE '%" + id + "%' OR T_CR_DR LIKE '%" + id + "%' OR T_MODE LIKE '%" + id + "%' ORDER BY T_ID DESC"
End If
Set R = C.Execute(sql)
Do While Not R.EOF
search.smfg.Rows = search.smfg.Rows + 1
For i = 0 To 6
search.smfg.TextMatrix(search.smfg.Rows - 1, i) = R.Fields(i) & ""
Next i
R.MoveNext
Loop
C.Close
End Function

Public Function setTransHeader()
search.smfg.Rows = 1
search.smfg.Cols = 7
search.smfg.TextMatrix(0, 0) = " Transaction ID"
search.smfg.TextMatrix(0, 1) = " Date"
search.smfg.TextMatrix(0, 2) = " Time"
search.smfg.TextMatrix(0, 3) = " Particular"
search.smfg.TextMatrix(0, 4) = " CR/DR"
search.smfg.TextMatrix(0, 5) = " Mode"
search.smfg.TextMatrix(0, 6) = " Amount"
search.smfg.ColWidth(0) = 5000
search.smfg.ColWidth(1) = 1500
search.smfg.ColWidth(2) = 1500
search.smfg.ColWidth(3) = 10000
search.smfg.ColWidth(4) = 1000
search.smfg.ColWidth(5) = 800
search.smfg.ColWidth(6) = 2000
End Function

'INSERT TRANSACTION DETAILS
Public Function insTrans(ByVal t_part As String, ByVal t_type As String, ByVal t_mode As String, ByVal t_amt As Double, Optional ByVal loan As Double = 0) As String
'On Error GoTo TRANS_ER_LB
sql = "INSERT INTO ER_MASTER_TRANSACTION VALUES('" + generateTransId(t_type) + "','" + Trim(Str(Date)) + "','" + Trim(Str(Time)) + "','" + t_part + "','" + t_type + "','" + t_mode + "'," + Str(t_amt) + ")"
connectOracle
C.Execute (sql)
C.Execute ("commit")
Select Case t_type
Case "CR": sql = "UPDATE ER_MASTER_ACCOUNT SET BALANCE=BALANCE+" + Str(t_amt) + ",LOAN=LOAN+" + Str(loan) + ",INCOME = INCOME+" + Str(t_amt) + " WHERE YEAR = '" + Trim(Str(Year(Date))) + "' AND MONTH = '" + getMonth(Month(Date)) + "'"
Case "DR": sql = "UPDATE ER_MASTER_ACCOUNT SET BALANCE=BALANCE-" + Str(t_amt) + ",LOAN=LOAN-" + Str(loan) + ",EXPENSE = EXPENSE+" + Str(t_amt) + " WHERE YEAR = '" + Trim(Str(Year(Date))) + "' AND MONTH = '" + getMonth(Month(Date)) + "'"
End Select
C.Execute (sql)
accUpdate
insTrans = "S"
Exit Function
TRANS_ER_LB:
MsgBox Err.Description
End Function

'GENERATE TRANCTION ID
Public Function generateTransId(ByVal t_type As String) As String
Dim a As Integer
Dim b As String
sql = "SELECT T_ID FROM ER_MASTER_TRANSACTION"
connectOracle
Set R = C.Execute(sql)
a = 0
Do While Not R.EOF
b = Mid(R.Fields(0), 8, 18)
a = Val(b)
R.MoveNext
Loop
b = ""
For i = 1 To 18 - digitCount(a)
b = b + "0"
Next i
b = b + Trim(Str(a + 1))
generateTransId = "T" + Trim(Str(Year(Date))) + t_type + b
End Function

'GET MONTHS IN THREE CHARACTER
Public Function getMonth(ByVal mon As Integer) As String
Select Case mon
Case 1: getMonth = "JAN"
Case 2: getMonth = "FEB"
Case 3: getMonth = "MAR"
Case 4: getMonth = "APR"
Case 5: getMonth = "MAY"
Case 6: getMonth = "JUN"
Case 7: getMonth = "JUL"
Case 8: getMonth = "AUG"
Case 9: getMonth = "SEP"
Case 10: getMonth = "OCT"
Case 11: getMonth = "NOV"
Case 12, 0: getMonth = "DEC"
End Select
End Function

'PREVIOUS YEAR
Public Function preGetYear() As String
If (Month(Date) - 1) = 0 Then
preGetYear = Year(Date) - 1
Else
preGetYear = Year(Date)
End If
End Function

'DIGIT COUNTER FUNCTION
Public Function digitCount(ByVal num As Long) As Integer
Dim count As Integer
Dim a As Integer
count = 0
a = num
Do While num > 0
num = num / 10
count = count + 1
Loop
If a Mod 9 = 0 Then
count = count + 1
End If
digitCount = count
End Function

'ACCOUNT UPDATE
Public Function accUpdate()
On Error GoTo AU_ER_LB
connectOracle
sql = "SELECT * FROM ER_MASTER_ACCOUNT WHERE YEAR = '" + Trim(Str(Year(Date))) + "'  AND MONTH ='" + getMonth(Month(Date)) + "'"
Set R = C.Execute(sql)
If R.EOF Then
Set R = C.Execute("SELECT * FROM ER_MASTER_ACCOUNT WHERE YEAR = '" + preGetYear() + "'  AND MONTH ='" + getMonth(Month(Date) - 1) + "'")
If R.EOF Then
C.Execute ("INSERT INTO ER_MASTER_ACCOUNT VALUES('" + Trim(Str(Year(Date))) + "','" + getMonth(Month(Date)) + "'," + Str(0) + "," + Str(0) + "," + Str(0) + "," + Str(0) + ")")
Else
C.Execute ("INSERT INTO ER_MASTER_ACCOUNT VALUES('" + Trim(Str(Year(Date))) + "','" + getMonth(Month(Date)) + "'," + Str(R.Fields(2)) + "," + Str(R.Fields(2)) + "," + Str(0) + "," + Str(0) + ")")
End If
Set R = C.Execute(sql)
End If
account.balance = R.Fields(2)
account.loan = R.Fields(3)
account.income = R.Fields(4)
account.expense = R.Fields(5)
C.Execute ("commit")
C.Close
Load account
Exit Function
AU_ER_LB:
MsgBox Err.Description
End Function
