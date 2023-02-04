VERSION 5.00
Begin VB.Form s_addnew 
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
      Height          =   8535
      Left            =   1920
      TabIndex        =   0
      Top             =   1200
      Width           =   14655
      Begin VB.TextBox Text12 
         Height          =   495
         Left            =   10680
         TabIndex        =   21
         Top             =   6240
         Width           =   3375
      End
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   5400
         TabIndex        =   20
         Top             =   6240
         Width           =   4695
      End
      Begin VB.TextBox Text10 
         Height          =   540
         Left            =   480
         TabIndex        =   19
         Top             =   6240
         Width           =   4335
      End
      Begin VB.TextBox Text9 
         Height          =   540
         Left            =   9480
         TabIndex        =   18
         Top             =   5160
         Width           =   4575
      End
      Begin VB.TextBox Text8 
         Height          =   540
         Left            =   480
         TabIndex        =   12
         Top             =   5160
         Width           =   8535
      End
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   9480
         TabIndex        =   11
         Top             =   4080
         Width           =   4575
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   4680
         TabIndex        =   10
         Top             =   4080
         Width           =   4335
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   480
         TabIndex        =   9
         Top             =   4080
         Width           =   3495
      End
      Begin VB.TextBox Text4 
         Height          =   540
         Left            =   7440
         TabIndex        =   8
         Top             =   2880
         Width           =   6615
      End
      Begin VB.TextBox Text3 
         Height          =   540
         Left            =   480
         TabIndex        =   7
         Top             =   2880
         Width           =   6255
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   7440
         TabIndex        =   5
         Top             =   1680
         Width           =   6615
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   540
         Left            =   480
         TabIndex        =   3
         Top             =   1680
         Width           =   6375
      End
      Begin VB.Label Label15 
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
         TabIndex        =   27
         Top             =   7320
         Width           =   1935
      End
      Begin VB.Label Label14 
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
         TabIndex        =   26
         Top             =   7320
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "Pincode:"
         Height          =   255
         Left            =   10680
         TabIndex        =   25
         Top             =   5880
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "State:"
         Height          =   300
         Left            =   5400
         TabIndex        =   24
         Top             =   5880
         Width           =   750
      End
      Begin VB.Label Label11 
         Caption         =   "City:"
         Height          =   375
         Left            =   480
         TabIndex        =   23
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Landmark:"
         Height          =   300
         Left            =   9480
         TabIndex        =   22
         Top             =   4800
         Width           =   1275
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Address line:"
         Height          =   300
         Left            =   480
         TabIndex        =   17
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Supplies PAN Card No."
         Height          =   300
         Left            =   9480
         TabIndex        =   16
         Top             =   3720
         Width           =   2775
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "GST  No."
         Height          =   300
         Left            =   4680
         TabIndex        =   15
         Top             =   3720
         Width           =   1110
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Mobile No."
         Height          =   300
         Left            =   480
         TabIndex        =   14
         Top             =   3720
         Width           =   2340
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Email:"
         Height          =   300
         Left            =   7440
         TabIndex        =   13
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Company Name"
         Height          =   300
         Left            =   480
         TabIndex        =   6
         Top             =   2520
         Width           =   1890
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Name:"
         Height          =   300
         Left            =   7440
         TabIndex        =   4
         Top             =   1320
         Width           =   2205
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Supplier ID:"
         Height          =   300
         Left            =   480
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   240
         Width           =   2565
      End
   End
End
Attribute VB_Name = "s_addnew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
mainCaption ("Supplier/addNew")
Me.Width = main.Width
Me.Height = main.Height
End Sub

