VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form comDetals 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   14460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   21540
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   14460
   ScaleWidth      =   21540
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   720
      TabIndex        =   16
      Top             =   480
      Width           =   19215
      Begin VB.TextBox pincode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   6240
         Width           =   6735
      End
      Begin VB.TextBox state 
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
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   6240
         Width           =   5775
      End
      Begin VB.TextBox addr 
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
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   5160
         Width           =   9855
      End
      Begin MSComCtl2.DTPicker est_txt 
         Height          =   495
         Left            =   720
         TabIndex        =   2
         Top             =   2880
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   164102145
         CurrentDate     =   44987
      End
      Begin VB.TextBox email 
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
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3960
         Width           =   6855
      End
      Begin VB.TextBox desc 
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
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   5160
         Width           =   7095
      End
      Begin VB.TextBox city 
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
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   6240
         Width           =   3975
      End
      Begin VB.TextBox mob 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   3960
         Width           =   5775
      End
      Begin VB.TextBox pan 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3960
         Width           =   3975
      End
      Begin VB.TextBox gst 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2880
         Width           =   6855
      End
      Begin VB.TextBox branch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2880
         Width           =   5775
      End
      Begin VB.TextBox com_txt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1680
         Width           =   9855
      End
      Begin VB.TextBox name_txt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   1680
         Width           =   7215
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   240
         Left            =   8520
         TabIndex        =   30
         Top             =   4800
         Width           =   945
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         Height          =   240
         Left            =   5280
         TabIndex        =   29
         Top             =   5880
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pincode"
         Height          =   240
         Left            =   11520
         TabIndex        =   28
         Top             =   5880
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Company Datails"
         BeginProperty Font 
            Name            =   "Narkisim"
            Size            =   23.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   720
         TabIndex        =   27
         Top             =   360
         Width           =   3285
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         Height          =   240
         Left            =   720
         TabIndex        =   26
         Top             =   5880
         Width           =   405
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   240
         Left            =   720
         TabIndex        =   25
         Top             =   4800
         Width           =   1200
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GST No"
         Height          =   240
         Left            =   11520
         TabIndex        =   24
         Top             =   2520
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No."
         Height          =   240
         Left            =   5280
         TabIndex        =   23
         Top             =   3600
         Width           =   1140
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
         Left            =   14040
         TabIndex        =   13
         Top             =   7200
         Width           =   1935
      End
      Begin VB.Label cancel_btn 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "                       Cancel"
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
         Left            =   16320
         TabIndex        =   14
         Top             =   7920
         Width           =   1935
      End
      Begin VB.Label edit_btn 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "                           Edit"
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
         Left            =   16320
         TabIndex        =   15
         Top             =   7200
         Width           =   1935
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barnch Name"
         Height          =   240
         Left            =   5280
         TabIndex        =   22
         Top             =   2520
         Width           =   1410
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         Height          =   240
         Left            =   11520
         TabIndex        =   21
         Top             =   3600
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EST Date"
         Height          =   240
         Left            =   720
         TabIndex        =   20
         Top             =   2520
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PAN NO."
         Height          =   240
         Left            =   720
         TabIndex        =   19
         Top             =   3600
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         Height          =   240
         Left            =   8520
         TabIndex        =   18
         Top             =   1320
         Width           =   1665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Owner's Name"
         Height          =   240
         Left            =   720
         TabIndex        =   17
         Top             =   1320
         Width           =   1515
      End
   End
End
Attribute VB_Name = "comDetals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub edit_btn_Click()
name_txt.Locked = False
com_txt.Locked = False
est_txt.Enabled = True
branch.Locked = False
gst.Locked = False
pan.Locked = False
mob.Locked = False
email.Locked = False
desc.Locked = False
addr.Locked = False
city.Locked = False
state.Locked = False
pincode.Locked = False
edit_btn.Enabled = False
edit_btn.Visible = False
update_btn.Enabled = True
update_btn.Visible = True
cancel_btn.Enabled = True
cancel_btn.Visible = True
End Sub

Private Sub Form_Load()
cancel_btn.Top = edit_btn.Top
End Sub

Public Function getComDetails()
connectOracle
sql = ""
Set R = C.Execute(sql)
name_txt.Text = R.Fields()
com_txt.Text = R.Fields()
est_txt.Value = R.Fields()
branch.Text = R.Fields()
gst.Text = R.Fields()
pan.Text = R.Fields()
mob.Text = R.Fields()
email.Text = R.Fields()
desc.Text = R.Fields()
addr.Text = R.Fields()
city.Text = R.Fields()
state.Text = R.Fields()
pincode.Text = R.Fields()

End Function
