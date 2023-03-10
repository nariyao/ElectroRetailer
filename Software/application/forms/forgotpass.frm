VERSION 5.00
Begin VB.Form forgotpass 
   Caption         =   "Forget Password"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10425
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
   ScaleHeight     =   3990
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   9855
      Begin VB.TextBox Text1 
         Height          =   540
         Left            =   1800
         TabIndex        =   2
         Top             =   840
         Width           =   5655
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
         Height          =   495
         Left            =   7440
         TabIndex        =   3
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User Id"
         Height          =   300
         Left            =   840
         TabIndex        =   1
         Top             =   960
         Width           =   900
      End
   End
End
Attribute VB_Name = "forgotpass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
