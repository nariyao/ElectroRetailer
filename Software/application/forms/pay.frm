VERSION 5.00
Begin VB.Form pay 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Payment"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5475
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
   ScaleHeight     =   5925
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox amt_txt 
      Height          =   420
      Left            =   360
      MaxLength       =   15
      TabIndex        =   7
      Top             =   4200
      Width           =   4695
   End
   Begin VB.ComboBox p_mode 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """?"" #,##0.00;(""?"" #,##0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   2
      EndProperty
      Height          =   420
      ItemData        =   "pay.frx":0000
      Left            =   360
      List            =   "pay.frx":0002
      TabIndex        =   5
      Top             =   3240
      Width           =   4695
   End
   Begin VB.TextBox reas_txt 
      Height          =   1140
      Left            =   360
      MaxLength       =   22
      TabIndex        =   3
      Top             =   1560
      Width           =   4695
   End
   Begin VB.TextBox to_txt 
      Height          =   420
      Left            =   360
      MaxLength       =   25
      TabIndex        =   1
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label pay_btn 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "                                                                         Pay"
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
      TabIndex        =   8
      Top             =   4920
      Width           =   4695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      Height          =   300
      Left            =   360
      TabIndex        =   6
      Top             =   3840
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Mode"
      Height          =   300
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   1785
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      DataMember      =   "`"
      Height          =   300
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   315
   End
End
Attribute VB_Name = "pay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public p_type As String
Public pay_status As String
Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = (main.ScaleHeight - Me.Height) / 2
Me.Left = (main.ScaleWidth - Me.Width) / 2
p_mode.AddItem "Cash"
p_mode.AddItem "Cheque"
p_mode.AddItem "Online"
pay_btn.Enabled = True
End Sub

Private Sub pay_btn_Click()
pay_btn.Enabled = True
If pay_status = "l" Then
temp = insTrans(to_txt.Text + " - " + reas_txt.Text, p_type, p_mode.Text, Str(amt_txt.Text), Str(amt_txt.Text))
Else
temp = insTrans(to_txt.Text + " - " + reas_txt.Text, p_type, p_mode.Text, Str(amt_txt.Text))
End If
If temp = "S" Then
MsgBox "Successful"
Unload Me
End If
End Sub


