VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm main 
   BackColor       =   &H8000000C&
   Caption         =   "ElectroRetailer"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   19500
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   8100
      Width           =   19500
      _ExtentX        =   34396
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb_admin 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   19500
      _ExtentX        =   34396
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton add_emp_mbtn 
         Caption         =   "Add Employee"
         Height          =   615
         Left            =   960
         TabIndex        =   12
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton dis_emp_btn 
         Caption         =   "Employee"
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   735
      End
   End
   Begin MSComctlLib.Toolbar tb_report 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   1980
      Width           =   19500
      _ExtentX        =   34396
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tb_sell 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   19500
      _ExtentX        =   34396
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton selldetByCust_tbn 
         Caption         =   "Command1"
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   735
      End
   End
   Begin MSComctlLib.Toolbar tb_order 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   660
      Width           =   19500
      _ExtentX        =   34396
      _ExtentY        =   1164
      ButtonWidth     =   1323
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton s_display_btn 
         Caption         =   "Display"
         Height          =   615
         Left            =   960
         TabIndex        =   7
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton addSup_btn 
         Caption         =   "Add"
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   735
      End
   End
   Begin MSComctlLib.Toolbar tb_home 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19500
      _ExtentX        =   34396
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton home_btn 
         Caption         =   "Home"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton logout_btn 
         BackColor       =   &H8000000D&
         Caption         =   "Logout"
         Height          =   615
         Left            =   27360
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Menu menuHome 
      Caption         =   "&Home"
   End
   Begin VB.Menu menuOrder 
      Caption         =   "&Order"
   End
   Begin VB.Menu menuSell 
      Caption         =   "&Sell"
   End
   Begin VB.Menu menuReport 
      Caption         =   "&Report"
   End
   Begin VB.Menu menuAdmin 
      Caption         =   "&Administation"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function tbMenu(ByVal a As Integer)
tb_home.Visible = False
tb_home.Enabled = False
tb_order.Visible = False
tb_order.Enabled = False
tb_sell.Visible = False
tb_sell.Enabled = False
tb_report.Visible = False
tb_report.Enabled = False
tb_admin.Visible = False
tb_admin.Enabled = False
Select Case a
Case 1: tb_home.Visible = True
        tb_home.Enabled = True
Case 2: tb_order.Visible = True
        tb_order.Enabled = True
Case 3: tb_sell.Visible = True
        tb_sell.Enabled = True
Case 4: tb_report.Visible = True
        tb_report.Enabled = True
Case 5: tb_admin.Visible = True
        tb_admin.Enabled = True
End Select
End Function



Private Sub add_emp_mbtn_Click()
insEmp
End Sub

Private Sub addSup_btn_Click()
supplier.Show
supplier.SetFocus
End Sub

Private Sub dis_emp_btn_Click()
displayEmp
End Sub

Private Sub home_btn_Click()
home.Show
home.SetFocus
End Sub

Private Sub logout_btn_Click()
logout
End Sub

Private Sub MDIForm_Activate()
tbMenu (1)

End Sub

Private Sub MDIForm_Load()
home.Show
End Sub

Private Sub menuAdmin_Click()
tbMenu (5)
End Sub

Private Sub menuHome_Click()
tbMenu (1)
End Sub


Private Sub menuOrder_Click()
tbMenu (2)
End Sub

Private Sub menuReport_Click()
tbMenu (4)
End Sub

Private Sub menuSell_Click()
tbMenu (3)
End Sub


Private Sub s_display_btn_Click()
DisplaySupplier
End Sub


Private Sub selldetByCust_tbn_Click()
customer.Show
customer.SetFocus
End Sub
