Attribute VB_Name = "mod_supplier"
Public Function DisplaySupplier()
Dim a As String
search.Show
search.SetFocus
main.Caption = mainCaption("suppleir Details")
SetSupplierHeader
smfgSupplier ("a")
End Function
Public Function smfgSupplier(ByVal id As String)
search.smfg.Rows = 1
connectOracle
If (id = "a") Then
sql = "SELECT * FROM ER_MASTER_SUPPLIER"
Else
sql = "SELECT * FROM ER_MASTER_SUPPLIER WHERE S_ID= '" + a + "'"
End If
Set R = C.Execute(sql)
Do While Not R.EOF
For i = 0 To 8
search.smfg.TextMatrix(search.smfg.Rows, i) = R.Fields(i) & ""
Next i
R.MoveNext
search.smfg.Rows = search.smfg.Rows + 1
Loop
C.Close
End Function

Private Function SetSupplierHeader()
search.smfg.Rows = 1
search.smfg.Cols = 9
search.smfg.TextMatrix(0, 0) = " Supplier ID"
search.smfg.TextMatrix(0, 1) = " Supplier Name"
search.smfg.TextMatrix(0, 2) = " Company Name"
search.smfg.TextMatrix(0, 3) = " Email"
search.smfg.TextMatrix(0, 4) = " Mobile"
search.smfg.TextMatrix(0, 5) = " GST No."
search.smfg.TextMatrix(0, 6) = " PAN Card"
search.smfg.TextMatrix(0, 7) = " Address"
search.smfg.TextMatrix(0, 8) = " Pincode"
search.smfg.ColWidth(0) = 2500
search.smfg.ColWidth(1) = 3500
search.smfg.ColWidth(2) = 5000
search.smfg.ColWidth(3) = 3500
search.smfg.ColWidth(4) = 1500
search.smfg.ColWidth(5) = 2000
search.smfg.ColWidth(6) = 1400
search.smfg.ColWidth(7) = 6800
search.smfg.ColWidth(8) = 1500
End Function

Public Function CreateSupId() As String

End Function
