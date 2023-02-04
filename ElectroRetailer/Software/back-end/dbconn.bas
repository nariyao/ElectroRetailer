Attribute VB_Name = "dbconn"
Public C As ADODB.connection
Public R As ADODB.Recordset
Public USERID As String
Public PASSWORD As String
Public sql As String
Public Function connection(USERID As String, PASSWORD As String)
Set C = New ADODB.connection
C.Open "provider=msdaora.1;user id=" + USERID + "/" + PASSWORD + ";persist security info=true"
Set R = New ADODB.Recordset
End Function
