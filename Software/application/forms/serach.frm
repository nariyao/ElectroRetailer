VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form search 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   14430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   962
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1265
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid smfg 
      Height          =   11535
      Left            =   600
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1800
      Width           =   27735
      _ExtentX        =   48921
      _ExtentY        =   20346
      _Version        =   393216
      Rows            =   6
      Cols            =   9
      FixedCols       =   0
      RowHeightMin    =   400
      BackColor       =   -2147483634
      BackColorFixed  =   -2147483635
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483647
      BackColorBkg    =   16777215
      GridColor       =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      TextStyleFixed  =   1
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      MousePointer    =   1
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
   Begin VB.TextBox search_box 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12000
      TabIndex        =   0
      Top             =   480
      Width           =   6975
   End
   Begin VB.Label SID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search:"
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
      Left            =   10920
      TabIndex        =   1
      Top             =   600
      Width           =   810
   End
End
Attribute VB_Name = "search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public searchStatus As String

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = main.ScaleWidth
Me.Height = main.ScaleHeight
End Sub

Private Sub Form_LostFocus()
Me.Hide
End Sub

Private Sub search_box_Change()
Select Case searchStatus
Case "emp": smfgEmp (search_box.Text)
Case "sup": smfgSup (search_box.Text)
End Select
End Sub

Private Sub smfg_DblClick()
Select Case searchStatus
Case "emp": empUpDel (smfg.TextMatrix(smfg.RowSel, 0))
Case "sup": supUpDel (smfg.TextMatrix(smfg.RowSel, 0))
End Select
End Sub


