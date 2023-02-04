VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form s_serach 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   809
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1265
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid s_mfg 
      Height          =   11535
      Left            =   600
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1800
      Width           =   27735
      _ExtentX        =   48921
      _ExtentY        =   20346
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      RowHeightMin    =   400
      BackColor       =   -2147483634
      BackColorFixed  =   -2147483635
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483638
      BackColorBkg    =   16777215
      GridColor       =   0
      WordWrap        =   -1  'True
      HighLight       =   2
      GridLines       =   3
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   ""
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
   Begin VB.TextBox Text1 
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
      Text            =   "Text1"
      Top             =   480
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Supplier Id:"
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
      Left            =   10080
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "s_serach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
mfg_setHeader
End Sub


Private Function mfg_setHeader()
s_mfg.TextMatrix(0, 0) = " Supplier ID"
s_mfg.TextMatrix(0, 1) = " Supplier Name"
s_mfg.TextMatrix(0, 2) = " Company Name"
s_mfg.TextMatrix(0, 3) = " Email"
s_mfg.TextMatrix(0, 4) = " Mobile"
s_mfg.TextMatrix(0, 5) = " GST No."
s_mfg.TextMatrix(0, 6) = " PAN Card"
s_mfg.TextMatrix(0, 7) = " Address"
s_mfg.TextMatrix(0, 8) = " Pincode"
s_mfg.ColWidth(0) = 2500
s_mfg.ColWidth(1) = 3500
s_mfg.ColWidth(2) = 5000
s_mfg.ColWidth(3) = 3500
s_mfg.ColWidth(4) = 1500
s_mfg.ColWidth(5) = 2000
s_mfg.ColWidth(6) = 1400
s_mfg.ColWidth(7) = 6800
s_mfg.ColWidth(8) = 1500
End Function

