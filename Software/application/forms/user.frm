VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form user 
   ClientHeight    =   13590
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20910
   LinkTopic       =   "Form1"
   ScaleHeight     =   13590
   ScaleWidth      =   20910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11175
      Begin VB.Frame Frame5 
         Caption         =   "Display User"
         Height          =   5655
         Left            =   5400
         TabIndex        =   12
         Top             =   360
         Width           =   5415
         Begin VB.TextBox serach_box 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   240
            TabIndex        =   14
            Top             =   840
            Width           =   4935
         End
         Begin MSFlexGridLib.MSFlexGrid smfg 
            Height          =   3855
            Left            =   240
            TabIndex        =   13
            Top             =   1560
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   6800
            _Version        =   393216
            Rows            =   10
            Cols            =   3
            FixedCols       =   0
            RowHeightMin    =   500
            BackColorFixed  =   -2147483635
            ForeColorFixed  =   -2147483634
            BackColorSel    =   -2147483647
            BackColorBkg    =   16777215
            FocusRect       =   0
            HighLight       =   2
            GridLinesFixed  =   0
            ScrollBars      =   2
            SelectionMode   =   1
            FormatString    =   "Login ID      | Password                    | Type  "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search"
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
            Left            =   240
            TabIndex        =   15
            Top             =   480
            Width           =   750
         End
      End
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   4695
         Begin VB.OptionButton ud_opt 
            Caption         =   "Update/Delete"
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
            Left            =   2520
            TabIndex        =   11
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton cr_opt 
            Caption         =   "Create"
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
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4575
         Left            =   360
         TabIndex        =   1
         Top             =   1440
         Width           =   4695
         Begin VB.ComboBox id 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            ItemData        =   "user.frx":0000
            Left            =   240
            List            =   "user.frx":0002
            TabIndex        =   8
            Top             =   720
            Width           =   4095
         End
         Begin VB.TextBox password 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   240
            TabIndex        =   3
            Top             =   1800
            Width           =   4095
         End
         Begin VB.ComboBox Ltype 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            ItemData        =   "user.frx":0004
            Left            =   240
            List            =   "user.frx":0006
            TabIndex        =   2
            Top             =   2880
            Width           =   4095
         End
         Begin VB.Label delete_btn 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "                       Delete"
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
            Left            =   2400
            TabIndex        =   17
            Top             =   3720
            Visible         =   0   'False
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
            Left            =   240
            TabIndex        =   16
            Top             =   3720
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Login ID"
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
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
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
            Left            =   240
            TabIndex        =   6
            Top             =   1440
            Width           =   1035
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
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
            Left            =   240
            TabIndex        =   5
            Top             =   2520
            Width           =   555
         End
         Begin VB.Label create_btn 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "                                                           Create"
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
            Left            =   240
            TabIndex        =   4
            Top             =   3720
            Width           =   4095
         End
      End
   End
End
Attribute VB_Name = "user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cr_opt_Click()
update_btn.Enabled = False
update_btn.Visible = False
delete_btn.Enabled = False
delete_btn.Visible = False
create_btn.Enabled = True
create_btn.Visible = True
getEmpID
End Sub

Private Sub create_btn_Click()
temp = createUser(id.Text, password.Text, Ltype.Text)
If (temp = 1) Then
clearVal
End If
End Sub

Public Sub clearVal()
id.Text = ""
password.Text = ""
Ltype.Text = ""
End Sub

Private Sub delete_btn_Click()
temp = MsgBox("Do you want to delete this user", vbOKCancel + vbExclamation, "Delete user")
If (temp = vbOK) Then
temp = deleteUser(id.Text)
End Sub

Private Sub Form_Load()
Ltype.AddItem "Super User (SU)"
Ltype.AddItem "Normal User (NU)"
displayUser ("a")
End Sub

Private Sub serach_box_Change()
displayUser (Trim(search_box.Text))
End Sub

Private Sub smfg_DblClick()
id.Text = smfg.TextMatrix(smfg.RowSel, 0)
password.Text = smfg.TextMatrix(smfg.RowSel, 1)
Ltype.Text = smfg.TextMatrix(smfg.RowSel, 2)
End Sub

Private Sub ud_opt_Click()
update_btn.Enabled = True
update_btn.Visible = True
delete_btn.Enabled = True
delete_btn.Visible = True
create_btn.Enabled = False
create_btn.Visible = False
id.Locked = True
End Sub

Public Function getEmpID()
connectOracle
Set R = C.Execute("SELECT E_ID FROM ER_MASTER_EMPLOYEE")
Do While Not R.EOF
id.AddItem R.Fields(0)
R.MoveNext
Loop
C.Close
End Function

Private Sub update_btn_Click()
temp = updateUser(id.Text, password.Text, Ltype.Text)
If (temp = 1) Then
clearVal
displayUser ("a")
End If
End Sub
