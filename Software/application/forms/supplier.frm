VERSION 5.00
Begin VB.Form supplier 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   23415
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12945
   ScaleWidth      =   23415
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   9015
      Left            =   5280
      TabIndex        =   9
      Top             =   1800
      Width           =   14655
      Begin VB.TextBox s_pincode 
         Height          =   495
         Left            =   9480
         MaxLength       =   6
         TabIndex        =   8
         Top             =   5160
         Width           =   4575
      End
      Begin VB.TextBox s_add_line 
         Height          =   540
         Left            =   480
         MaxLength       =   100
         TabIndex        =   7
         Top             =   5160
         Width           =   8535
      End
      Begin VB.TextBox s_pan 
         Height          =   495
         Left            =   9480
         MaxLength       =   10
         TabIndex        =   6
         Top             =   4080
         Width           =   4575
      End
      Begin VB.TextBox s_gstno 
         Height          =   495
         Left            =   4680
         MaxLength       =   15
         TabIndex        =   5
         Top             =   4080
         Width           =   4335
      End
      Begin VB.TextBox s_mobile 
         Height          =   495
         Left            =   480
         MaxLength       =   10
         TabIndex        =   4
         Top             =   4080
         Width           =   3495
      End
      Begin VB.TextBox s_email 
         Height          =   540
         Left            =   7440
         MaxLength       =   30
         TabIndex        =   3
         Top             =   2880
         Width           =   6615
      End
      Begin VB.TextBox company_name 
         Height          =   540
         Left            =   480
         MaxLength       =   50
         TabIndex        =   2
         Top             =   2880
         Width           =   6255
      End
      Begin VB.TextBox s_name 
         Height          =   495
         Left            =   7440
         MaxLength       =   30
         TabIndex        =   1
         Top             =   1680
         Width           =   6615
      End
      Begin VB.TextBox s_id 
         Height          =   540
         Left            =   480
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1680
         Width           =   6255
      End
      Begin VB.Label sup_msg_lb 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   480
         TabIndex        =   24
         Top             =   6000
         Visible         =   0   'False
         Width           =   6960
      End
      Begin VB.Label delete_btn 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "                       Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   9840
         TabIndex        =   23
         Top             =   6120
         Width           =   1935
      End
      Begin VB.Label update_btn 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "                      Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   7680
         TabIndex        =   22
         Top             =   6120
         Width           =   1935
      End
      Begin VB.Label add_btn 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "                           Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   9840
         TabIndex        =   21
         Top             =   6840
         Width           =   1935
      End
      Begin VB.Label close_btn 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "                         Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   12000
         TabIndex        =   20
         Top             =   6120
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "Pincode:"
         Height          =   255
         Left            =   9480
         TabIndex        =   19
         Top             =   4800
         Width           =   2295
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Address line:"
         Height          =   300
         Left            =   480
         TabIndex        =   18
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Supplies PAN Card No."
         Height          =   300
         Left            =   9480
         TabIndex        =   17
         Top             =   3720
         Width           =   2775
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "GST  No."
         Height          =   300
         Left            =   4680
         TabIndex        =   16
         Top             =   3720
         Width           =   1110
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Mobile No."
         Height          =   300
         Left            =   480
         TabIndex        =   15
         Top             =   3720
         Width           =   2340
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Email:"
         Height          =   300
         Left            =   7440
         TabIndex        =   14
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Company Name"
         Height          =   300
         Left            =   480
         TabIndex        =   13
         Top             =   2520
         Width           =   1890
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Name:"
         Height          =   300
         Left            =   7440
         TabIndex        =   12
         Top             =   1320
         Width           =   2205
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Supplier ID:"
         Height          =   300
         Left            =   480
         TabIndex        =   11
         Top             =   1320
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Add New Supplier"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   10
         Top             =   240
         Width           =   2565
      End
   End
End
Attribute VB_Name = "supplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CLOSE BUTTON
Private Sub close_btn_Click()
Unload Me
displaySupplier
End Sub

'ADD NEW SUPPLIER
Private Sub add_btn_Click()
On Error GoTo SUP_ER_INS
connectOracle
sql = "INSERT INTO ER_MASTER_SUPPLIER VALUES('" + s_id.Text + "','" + s_name.Text + "','" + company_name.Text + "','" + s_email.Text + "','" + s_mobile.Text + "','" + s_gstno.Text + "','" + s_pan.Text + "','" + s_add_line.Text + "','" + s_pincode.Text + "')"
C.Execute (sql)
C.Execute ("commit")
C.Close
temp = supMsg("S", "Supplier data inserted")
clearSup
s_id.Text = generateSupId
Exit Sub
SUP_ER_INS:
temp = supMsg("E", "")
End Sub

'UPDATE SUPPLIER DETAILS
Private Sub update_btn_Click()
On Error GoTo SUP_ER_UP
Dim sql As String
sql = "UPDATE ER_MASTER_SUPPLIER SET S_NAME='" + s_name.Text + "',COMPANY_NAME='" + company_name.Text + "',S_EMAIL='" + s_email.Text + "',S_MOBILE='" + s_mobile.Text + "',S_GSTNO='" + s_gstno.Text + "',S_PAN='" + s_pan.Text + "',S_ADDRESS='" + s_add_line.Text + "',S_PINCODE='" + s_pincode.Text + "' WHERE S_ID='" + s_id.Text + "'"
connectOracle
C.Execute (sql)
C.Execute ("commit")
C.Close
temp = supMsg("S", "Updated")
Exit Sub
SUP_ER_UP:
temp = supMsg("E", "")
End Sub

'DELETE SUPPLIER DETAILS
Private Sub delete_btn_Click()
temp = MsgBox("It is going to be deleted permanently.", vbExclamation + vbOKCancel, "Deleting Supplier Details")
If temp = vbOK Then
On Error GoTo SUP_ER_DEL
sql = "DELETE FROM ER_MASTER_SUPPLIER WHERE e_id='" + id.Text + "'"
connectOracle
C.Execute (sql)
C.Execute ("commit")
C.Close
MsgBox "Employee ID deleted as well as its related data"
Unload Me
clearSup
displaySupplier
End If
Exit Sub
SUP_ER_DEL:
temp = supMsg("E", "")
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = main.ScaleWidth
Me.Height = main.ScaleHeight
add_btn.Top = delete_btn.Top
End Sub


Private Function supMsg(ByVal msg_type As String, ByVal msg As String)
Select Case msg_type
Case "S":   sup_msg_lb.Caption = "Succes! " + msg
            sup_msg_lb.Visible = True
            sup_msg_lb.ForeColor = vbGreen
Case "E":   sup_msg_lb.Caption = "Error! " + Str(Err.Number) + " : " + Err.Description
            sup_msg_lb.Visible = True
            sup_msg_lb.ForeColor = vbRed
End Select
End Function

