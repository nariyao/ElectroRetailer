VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form customer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   15645
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   28320
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   15645
   ScaleWidth      =   28320
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   10935
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   25575
      Begin VB.Frame Frame2 
         Caption         =   "Update Customer Details"
         Height          =   8535
         Left            =   19320
         TabIndex        =   4
         Top             =   1800
         Width           =   5535
         Begin VB.TextBox c_gst 
            Height          =   495
            Left            =   360
            TabIndex        =   10
            Top             =   5280
            Width           =   4815
         End
         Begin VB.TextBox c_addr 
            Height          =   495
            Left            =   360
            TabIndex        =   9
            Top             =   6360
            Width           =   4815
         End
         Begin VB.TextBox c_name 
            Height          =   495
            Left            =   360
            TabIndex        =   8
            Top             =   2040
            Width           =   4815
         End
         Begin VB.TextBox c_mob 
            Height          =   495
            Left            =   360
            TabIndex        =   7
            Top             =   3120
            Width           =   4815
         End
         Begin VB.TextBox c_email 
            Height          =   495
            Left            =   360
            TabIndex        =   6
            Top             =   4200
            Width           =   4815
         End
         Begin VB.TextBox c_id 
            Height          =   495
            Left            =   360
            TabIndex        =   5
            Top             =   960
            Width           =   4815
         End
         Begin VB.Label c_msg_lb 
            Caption         =   "Label8"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   18
            Top             =   7920
            Visible         =   0   'False
            Width           =   4815
         End
         Begin VB.Label update_btn 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "                                                                      Update"
            Enabled         =   0   'False
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
            Left            =   360
            TabIndex        =   17
            Top             =   7200
            Width           =   4815
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Id"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   16
            Top             =   600
            Width           =   1245
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   15
            Top             =   6000
            Width           =   885
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GST No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   14
            Top             =   4920
            Width           =   900
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   13
            Top             =   1680
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   12
            Top             =   2760
            Width           =   1140
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   11
            Top             =   3840
            Width           =   600
         End
      End
      Begin MSFlexGridLib.MSFlexGrid cmfg 
         Height          =   8415
         Left            =   720
         TabIndex        =   3
         Top             =   2040
         Width           =   18015
         _ExtentX        =   31776
         _ExtentY        =   14843
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   -2147483635
         ForeColorFixed  =   -2147483628
         BackColorSel    =   -2147483647
         BackColorBkg    =   16777215
         GridColor       =   0
         GridColorFixed  =   -2147483630
         WordWrap        =   -1  'True
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"customer.frx":0000
      End
      Begin VB.TextBox search_box 
         Height          =   495
         Left            =   8400
         TabIndex        =   1
         Top             =   840
         Width           =   7575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Search"
         Height          =   300
         Left            =   7320
         TabIndex        =   2
         Top             =   960
         Width           =   870
      End
   End
End
Attribute VB_Name = "customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = main.ScaleWidth
Me.Height = main.ScaleHeight
cmfgCustomer ("a")
End Sub

Private Sub cmfg_DblClick()
c_id.Text = cmfg.TextMatrix(cmfg.RowSel, 0)
c_name.Text = cmfg.TextMatrix(cmfg.RowSel, 1)
c_mob.Text = cmfg.TextMatrix(cmfg.RowSel, 2)
c_email.Text = cmfg.TextMatrix(cmfg.RowSel, 3)
c_gst.Text = cmfg.TextMatrix(cmfg.RowSel, 4)
c_addr.Text = cmfg.TextMatrix(cmfg.RowSel, 5)
update_btn.Enabled = True
End Sub


Private Sub update_btn_Click()
On Error GoTo C_ER_LB
sql = "UPDATE ER_MASTER_CUSTOMER SET C_NAME='" + c_name.Text + "', C_MOBILE='" + c_mob.Text + "', C_EMAIL='" + c_email.Text + "',C_GSTNO='" + c_gst.Text + "',C_ADDR='" + c_addr.Text + "'WHERE C_ID='" + c_id.Text + "'"
connectOracle
C.Execute (sql)
C.Execute ("Commit")
C.Close
c_msg_lb.Caption = "Succes! Updated"
c_msg_lb.ForeColor = vbGreen
c_msg_lb.Visible = True
Exit Sub
C_ER_LB:
c_msg_lb.Caption = "Error! " + Str(Err.Number) + " : " + Err.Description
c_msg_lb.ForeColor = vbRed
c_msg_lb.Visible = True
End Sub

Private Sub search_box_Change()
cmfgCustomer (Trim(search_box.Text))
End Sub

'CUSTOMER SEARCH
Public Function cmfgCustomer(ByVal id As String)
cmfg.Rows = 1
connectOracle
If (id = "a") Then
sql = "SELECT * FROM ER_MASTER_CUSTOMER"
Else
sql = "SELECT * FROM ER_MASTER_CUSTOMER WHERE C_ID LIKE '%" + id + "%' OR C_NAME LIKE '%" + id + "%' OR C_MOBILE LIKE '%" + id + "%' OR C_GSTNO LIKE '%" + id + "%'"
End If
Set R = C.Execute(sql)
Do While Not R.EOF
cmfg.Rows = cmfg.Rows + 1
For i = 0 To 5
cmfg.TextMatrix(cmfg.Rows - 1, i) = R.Fields(i) & ""
Next i
R.MoveNext
Loop
C.Close
End Function
