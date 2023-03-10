VERSION 5.00
Begin VB.Form account 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   14460
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   28470
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   14460
   ScaleWidth      =   28470
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   8655
      Left            =   2520
      TabIndex        =   0
      Top             =   1320
      Width           =   19455
      Begin VB.Frame Frame5 
         BackColor       =   &H8000000B&
         Height          =   2415
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   18975
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CURRENT BALANCE"
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
            TabIndex        =   22
            Top             =   360
            Width           =   2595
         End
         Begin VB.Label balance 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   1440
            TabIndex        =   21
            Top             =   1080
            Width           =   1230
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rs. "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   360
            TabIndex        =   20
            Top             =   1080
            Width           =   1140
         End
         Begin VB.Label deposit_btn 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "                               Deposit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   735
            Left            =   15960
            TabIndex        =   19
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label withdraw_btn 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "                               Withdraw"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   735
            Left            =   15960
            TabIndex        =   18
            Top             =   1320
            Width           =   2655
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Height          =   2415
         Left            =   240
         TabIndex        =   11
         Top             =   3480
         Width           =   18975
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LOAN AMOUNT"
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
            TabIndex        =   16
            Top             =   360
            Width           =   1905
         End
         Begin VB.Label loan 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   1440
            TabIndex        =   15
            Top             =   1080
            Width           =   1230
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rs. "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   360
            TabIndex        =   14
            Top             =   1080
            Width           =   1140
         End
         Begin VB.Label getloan_btn 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "                                       Get Loan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   735
            Left            =   15960
            TabIndex        =   13
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Label payloan_btn 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "                                      Pay Loan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   735
            Left            =   15960
            TabIndex        =   12
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000B&
         Height          =   2415
         Left            =   240
         TabIndex        =   6
         Top             =   6000
         Width           =   9375
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "INCOME AMOUNT"
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
            TabIndex        =   10
            Top             =   360
            Width           =   2235
         End
         Begin VB.Label income 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   1440
            TabIndex        =   9
            Top             =   1080
            Width           =   1230
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rs. "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   360
            TabIndex        =   8
            Top             =   1080
            Width           =   1140
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "INCOME PER MONTH"
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
            Left            =   7200
            TabIndex        =   7
            Top             =   2040
            Width           =   1905
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H8000000B&
         Height          =   2415
         Left            =   9840
         TabIndex        =   1
         Top             =   6000
         Width           =   9375
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EXPENSE AMOUNT"
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
            TabIndex        =   5
            Top             =   360
            Width           =   2445
         End
         Begin VB.Label expense 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   690
            Left            =   1440
            TabIndex        =   4
            Top             =   1080
            Width           =   1230
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rs. "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   360
            TabIndex        =   3
            Top             =   1080
            Width           =   1140
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EXPENSE PER MONTH"
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
            Left            =   7080
            TabIndex        =   2
            Top             =   2040
            Width           =   2040
         End
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNTS"
         BeginProperty Font 
            Name            =   "Narkisim"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   23
         Top             =   360
         Width           =   2400
      End
   End
End
Attribute VB_Name = "account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Top = 0
Left = 0
Me.Width = main.ScaleWidth
Me.Height = main.ScaleHeight
accUpdate
End Sub

Private Sub deposit_btn_Click()
pay.Show
pay.SetFocus
pay.p_type = "CR"
End Sub

Private Sub withdraw_btn_Click()
pay.Show
pay.SetFocus
pay.p_type = "DR"
End Sub

Private Sub getloan_btn_Click()
pay.Show
pay.SetFocus
pay.p_type = "CR"
End Sub

Private Sub payloan_btn_Click()
pay.Show
pay.SetFocus
pay.p_type = "DR"
End Sub


