VERSION 5.00
Begin VB.Form db 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Database setup: Home"
   ClientHeight    =   6090
   ClientLeft      =   2085
   ClientTop       =   1905
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame pg_2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5055
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Frame cau_fr 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   3495
         Left            =   0
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   5055
         Begin VB.CommandButton addUser 
            Caption         =   "Add"
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
            Height          =   495
            Left            =   240
            TabIndex        =   21
            Top             =   2640
            Width           =   4455
         End
         Begin VB.TextBox cau_pass 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            IMEMode         =   3  'DISABLE
            Left            =   240
            PasswordChar    =   "*"
            TabIndex        =   20
            Top             =   2040
            Width           =   4455
         End
         Begin VB.TextBox cau_user 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   19
            Top             =   1080
            Width           =   4455
         End
         Begin VB.Label er_2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New Admin user created"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   1200
            TabIndex        =   23
            Top             =   3240
            Visible         =   0   'False
            Width           =   2205
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Create Administrator User"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   3030
         End
         Begin VB.Label Label4 
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
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1680
            Width           =   3255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Username"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   720
            Width           =   2775
         End
      End
      Begin VB.Frame opt_fr 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1215
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   9615
         Begin VB.OptionButton no_opt 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No, I want to use existing database "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Value           =   -1  'True
            Width           =   3375
         End
         Begin VB.OptionButton yes_opt 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Yes, I want to create new database"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "and  delete existing one"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3550
            TabIndex        =   24
            Top             =   285
            Width           =   2535
         End
      End
   End
   Begin VB.Frame pg_1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   9615
      Begin VB.CommandButton test_conn_btn 
         Caption         =   "Test Connection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4920
         Picture         =   "db.frx":0000
         TabIndex        =   2
         Top             =   2640
         Width           =   4455
      End
      Begin VB.TextBox db_password 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   4920
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   2040
         Width           =   4455
      End
      Begin VB.TextBox db_user 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4920
         TabIndex        =   0
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Image Image1 
         Height          =   2355
         Left            =   720
         Picture         =   "db.frx":5D07
         Top             =   1200
         Width           =   2820
      End
      Begin VB.Label er_1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid username or password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   6000
         TabIndex        =   11
         Top             =   3480
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
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
         Left            =   4920
         TabIndex        =   10
         Top             =   1800
         Width           =   1320
      End
      Begin VB.Label username_db 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USERNAME:"
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
         Left            =   4920
         TabIndex        =   9
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Oracle DBA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6120
         TabIndex        =   8
         Top             =   240
         Width           =   2160
      End
   End
   Begin VB.CommandButton close_btn 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton finish_btn 
      Caption         =   "Finish"
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
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton next_btn 
      Caption         =   "Next"
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
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton previous_btn 
      Caption         =   "Previous"
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
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
End
Attribute VB_Name = "db"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msg As String


Private Sub addUser_Click()
sql = connection("electro", "retailer")
sql = "INSERT INTO ER_MASTER_LOGIN VALUES('" + cau_user.Text + "','" + cau_pass.Text + "','SU')"
C.Execute (sql)
er_2.Visible = True
addUser.Enabled = False
finish_btn.Enabled = True
End Sub

Private Sub cau_pass_Change()
cau_pass.Text = Trim(cau_pass.Text)
enb_addUser
End Sub

Private Sub cau_user_Change()
cau_user.Text = Trim(cau_user.Text)
enb_addUser
End Sub

Private Sub close_btn_Click()
Unload Me
End Sub
Function enb_addUser()
If cau_user.Text = blank Or cau_pass.Text = blank Then
addUser.Enabled = False
Else
addUser.Enabled = True
End If
End Function

Private Sub finish_btn_Click()
Unload Me
End Sub

Private Sub next_btn_Click()
pg_1.Enabled = False
pg_1.Visible = False
next_btn.Enabled = False
pg_2.Enabled = True
pg_2.Visible = True
previous_btn.Enabled = True
End Sub

Private Sub previous_btn_Click()
pg_1.Enabled = True
pg_1.Visible = True
next_btn.Enabled = True
pg_2.Enabled = False
pg_2.Visible = False
previous_btn.Enabled = False
no_opt.Value = True
End Sub

Private Sub test_conn_btn_Click()
On Error GoTo er
sql = connection(db_user.Text, db_password.Text)
C.Close
er_1.Caption = "Connected"
er_1.ForeColor = RGB(0, 255, 0)
er_1.Visible = True
next_btn.Enabled = True
Exit Sub
er:
If (Err.Number = -2147217843) Then
er_1.Caption = "Invalid UserName Or PASSWORD"
er_1.ForeColor = RGB(255, 0, 0)
er_1.Visible = True
Else
MsgBox "Something went wrong. Please restart application"
End If
next_btn.Enabled = False
End Sub

Private Sub yes_opt_Click()
USERID = db_user.Text
PASSWORD = db_password.Text
yes_no_opt
If msg = vbOK Then
progress.Show
opt_fr.Enabled = False
previous_btn.Enabled = False
ElseIf msg = vbCancel Then
no_opt.Value = True
End If
End Sub

Private Sub no_opt_Click()
yes_no_opt
End Sub

Function yes_no_opt()
If (yes_opt.Value = True) Then
cau_fr.Visible = True
msg = MsgBox("It will be going create new DBA as well as database", vbExclamation + vbOKCancel, "Database")
Else
cau_fr.Visible = False
cau_fr.Enabled = False
End If
End Function
