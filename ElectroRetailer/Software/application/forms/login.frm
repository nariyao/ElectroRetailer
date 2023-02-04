VERSION 5.00
Begin VB.Form login 
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   6705
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   5040
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox pass 
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
      IMEMode         =   3  'DISABLE
      Left            =   360
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "narayan"
      ToolTipText     =   "Enter password"
      Top             =   4320
      Width           =   4215
   End
   Begin VB.TextBox userid 
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
      Left            =   360
      TabIndex        =   0
      Text            =   "amit"
      ToolTipText     =   "Enter User ID"
      Top             =   3360
      Width           =   4215
   End
   Begin VB.Label er_login 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invalid Username or Password"
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
      Left            =   960
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Label close_btn 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "                                                                                       Close"
      ForeColor       =   &H8000000E&
      Height          =   540
      Left            =   360
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   5520
      Width           =   4215
   End
   Begin VB.Label login_btn 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "                                                                                       Login"
      Enabled         =   0   'False
      ForeColor       =   &H8000000E&
      Height          =   540
      Left            =   360
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   4920
      Width           =   4215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Click Here"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3000
      TabIndex        =   4
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "did you forget password?"
      Height          =   195
      Left            =   1080
      TabIndex        =   8
      Top             =   6120
      Width           =   1770
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   7
      Top             =   3960
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label brand_name 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spice HotSpot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   5
      Top             =   2280
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   2160
      Picture         =   "login.frx":0000
      Top             =   960
      Width           =   480
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As Integer

Private Sub close_btn_Click()
Unload Me
End Sub

Private Sub Form_Load()
b = 0
Label1.ForeColor = RGB(10, 38, 71)
End Sub

Private Sub login_btn_Click()
Dim a As Variant
a = authoCheck(userid.Text, pass.Text)
If a = 1 Then
Unload Me
main.Show
ElseIf a = 0 Then
b = b + 1
If (b = 3) Then
Unload Me
Exit Sub
End If
er_login.Visible = True
userid.Text = blank
pass.Text = blank
userid.SetFocus
Else
a = MsgBox("We got an error. Please restart application again", vbOKOnly + vbExclamation, "Error")
If a = vbOK Then
Unload Me
End If
End If
End Sub

Private Sub userid_Change()
userid.Text = Trim(userid.Text)
If userid.Text = blank Or pass.Text = blank Then
login_btn.Enabled = False
Else
login_btn.Enabled = True
End If
End Sub

Private Sub pass_Change()
pass.Text = Trim(pass.Text)
If userid.Text = blank Or pass.Text = blank Then
login_btn.Enabled = False
Else
login_btn.Enabled = True
End If
End Sub
