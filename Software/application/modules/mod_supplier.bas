Attribute VB_Name = "mod_supplier"
'DISPLAY SUPPLIER DETAILS
Public Function displaySupplier()
Dim a As String
search.Show
search.SetFocus
search.search_box.Text = ""
search.search_box.ToolTipText = "Enter Id/Name/Company Name "
search.searchStatus = "sup"
main.Caption = mainCaption("suppleir Details")
SetSupplierHeader
smfgSup ("a")
End Function
Public Function smfgSup(ByVal id As String)
search.smfg.Rows = 1
connectOracle
If (id = "a") Then
sql = "SELECT * FROM ER_MASTER_SUPPLIER"
Else
sql = "SELECT * FROM ER_MASTER_SUPPLIER WHERE S_ID LIKE '%" + id + "%' OR S_NAME LIKE '%" + id + "%' OR COMPANY_NAME LIKE '%" + id + "%'"
End If
Set R = C.Execute(sql)
Do While Not R.EOF
search.smfg.Rows = search.smfg.Rows + 1
For i = 0 To 8
search.smfg.TextMatrix(search.smfg.Rows - 1, i) = R.Fields(i) & ""
Next i
R.MoveNext
Loop
C.Close
End Function

'it will set header for grid
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
search.smfg.ColWidth(0) = 1500
search.smfg.ColWidth(1) = 3000
search.smfg.ColWidth(2) = 5000
search.smfg.ColWidth(3) = 3500
search.smfg.ColWidth(4) = 1500
search.smfg.ColWidth(5) = 2000
search.smfg.ColWidth(6) = 1400
search.smfg.ColWidth(7) = 5000
search.smfg.ColWidth(8) = 1500
End Function

'TO INSERT NEW SUPPLIER
Public Function insSupplier()
clearSup
supplier.Show
supplier.SetFocus
mainCaption ("Add New Supplier")
supplier.s_id.Text = generateSupId
supplier.add_btn.Enabled = True
supplier.add_btn.Visible = True
supplier.delete_btn.Enabled = False
supplier.delete_btn.Visible = False
supplier.update_btn.Enabled = False
supplier.update_btn.Visible = False
End Function

'THIS WILL GENERATE SUPPLIER ID
Public Function generateSupId() As String
Dim a As Integer
Dim b As String
connectOracle
Set R = New ADODB.Recordset
Set R = C.Execute("SELECT S_ID FROM ER_MASTER_SUPPLIER")
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
generateSupId = "SUP" + Trim(Str(Year(Date))) + Trim(b)
R.Close
C.Close
End Function

'UPDATE AND DELETE SUPPLIER
Public Function supUpDel(ByVal id As String)
connectOracle
sql = "SELECT * FROM ER_MASTER_SUPPLIER WHERE S_ID='" + id + "'"
Set R = C.Execute(sql)
temp = setSupValues(R.Fields(0), R.Fields(1), R.Fields(2), R.Fields(3), R.Fields(4), R.Fields(5), R.Fields(6), R.Fields(7), R.Fields(8))
supplier.Show
supplier.SetFocus
supplier.add_btn.Enabled = False
supplier.add_btn.Visible = False
supplier.delete_btn.Enabled = True
supplier.delete_btn.Visible = True
supplier.update_btn.Enabled = True
supplier.update_btn.Visible = True
End Function

'SET VALUES IN SUPPLIER FORM
Public Function setSupValues(ByVal sup_id As String, ByVal sup_name As String, ByVal sup_co As String, ByVal sup_email As String, ByVal sup_mob As String, ByVal sup_gst As String, ByVal sup_pan As String, ByVal sup_addr As String, ByVal sup_pin As String)
supplier.s_id.Text = sup_id
supplier.s_name.Text = sup_name
supplier.company_name.Text = sup_co
supplier.s_email.Text = sup_email
supplier.s_mobile.Text = sup_mob
supplier.s_gstno.Text = sup_gst
supplier.s_pan.Text = sup_pan
supplier.s_add_line.Text = sup_addr
supplier.s_pincode.Text = sup_pin
End Function

'CLEAR VALUES OF SUPPLIER FORM
Public Function clearSup()
supplier.s_id.Text = ""
supplier.s_name.Text = ""
supplier.company_name.Text = ""
supplier.s_email.Text = ""
supplier.s_mobile.Text = ""
supplier.s_gstno.Text = ""
supplier.s_pan.Text = ""
supplier.s_add_line.Text = ""
supplier.s_pincode.Text = ""
End Function
