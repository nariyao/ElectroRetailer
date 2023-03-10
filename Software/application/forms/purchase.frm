VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form purchase 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   14445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   24195
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
   MinButton       =   0   'False
   ScaleHeight     =   14445
   ScaleWidth      =   24195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   14175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   23775
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000B&
         Caption         =   "Products"
         Height          =   10935
         Left            =   240
         TabIndex        =   18
         Top             =   3000
         Width           =   23295
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000B&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   1
            EndProperty
            Height          =   600
            Left            =   20130
            TabIndex        =   21
            Text            =   "0.00"
            Top             =   8520
            Width           =   2760
         End
         Begin MSFlexGridLib.MSFlexGrid pur_mfg 
            Height          =   5700
            Left            =   120
            TabIndex        =   19
            Top             =   1080
            Width           =   22815
            _ExtentX        =   40243
            _ExtentY        =   10054
            _Version        =   393216
            Rows            =   14
            Cols            =   9
            FixedCols       =   0
            RowHeightMin    =   400
            BackColorFixed  =   -2147483635
            ForeColorFixed  =   -2147483634
            BackColorSel    =   -2147483647
            BackColorBkg    =   16777215
            FocusRect       =   0
            HighLight       =   2
            GridLinesFixed  =   1
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "                          Print"
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
            Left            =   16680
            TabIndex        =   36
            Top             =   9960
            Width           =   1935
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            Height          =   420
            Left            =   20280
            TabIndex        =   35
            Top             =   9240
            Width           =   2535
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            Height          =   420
            Left            =   20280
            TabIndex        =   34
            Top             =   8040
            Width           =   2535
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            Height          =   420
            Left            =   20280
            TabIndex        =   33
            Top             =   7440
            Width           =   2535
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            Height          =   420
            Left            =   20280
            TabIndex        =   32
            Top             =   6840
            Width           =   2535
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label15"
            Height          =   300
            Left            =   480
            TabIndex        =   31
            Top             =   8280
            Width           =   975
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label14"
            Height          =   300
            Left            =   480
            TabIndex        =   30
            Top             =   6960
            Width           =   975
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remainting Amount"
            Height          =   300
            Left            =   16680
            TabIndex        =   29
            Top             =   9240
            Width           =   2385
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pay"
            Height          =   300
            Left            =   16680
            TabIndex        =   28
            Top             =   8640
            Width           =   450
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount"
            Height          =   300
            Left            =   16680
            TabIndex        =   27
            Top             =   8040
            Width           =   1620
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SGST"
            Height          =   300
            Left            =   16680
            TabIndex        =   26
            Top             =   7560
            Width           =   735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CGST"
            Height          =   300
            Left            =   16680
            TabIndex        =   25
            Top             =   6960
            Width           =   735
         End
         Begin VB.Label Cancel_btn 
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
            Left            =   14520
            TabIndex        =   24
            Top             =   9960
            Width           =   1935
         End
         Begin VB.Label print_bill_btn 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "                           Pay"
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
            Left            =   18840
            TabIndex        =   23
            Top             =   9960
            Width           =   1935
         End
         Begin VB.Label next_btn 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "                          Next"
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
            Left            =   21000
            TabIndex        =   22
            Top             =   9960
            Width           =   1935
         End
         Begin VB.Shape Shape6 
            Height          =   1215
            Left            =   120
            Top             =   6720
            Width           =   16020
         End
         Begin VB.Shape Shape1 
            Height          =   8655
            Left            =   120
            Top             =   1080
            Width           =   22815
         End
         Begin VB.Shape Shape2 
            Height          =   615
            Left            =   16130
            Top             =   7320
            Width           =   6810
         End
         Begin VB.Shape Shape3 
            Height          =   615
            Left            =   16130
            Top             =   8520
            Width           =   6800
         End
         Begin VB.Shape Shape5 
            Height          =   3015
            Left            =   16130
            Top             =   6720
            Width           =   3985
         End
         Begin VB.Label addPro_btn 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "                           Add Product"
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
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Invoice Details"
         Height          =   2655
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   23295
         Begin VB.TextBox Text7 
            Height          =   540
            Left            =   10320
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1920
            Width           =   12495
         End
         Begin VB.TextBox Text6 
            Height          =   540
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1920
            Width           =   4575
         End
         Begin VB.TextBox Text5 
            Height          =   540
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1920
            Width           =   4575
         End
         Begin VB.TextBox Text4 
            Height          =   540
            Left            =   20400
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   840
            Width           =   2415
         End
         Begin VB.TextBox Text3 
            Height          =   540
            Left            =   15360
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   840
            Width           =   4575
         End
         Begin VB.TextBox Text2 
            Height          =   540
            Left            =   10320
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   840
            Width           =   4575
         End
         Begin VB.ComboBox Combo1 
            Height          =   420
            Left            =   5280
            TabIndex        =   4
            Text            =   "Combo1"
            Top             =   840
            Width           =   4575
         End
         Begin VB.TextBox Text1 
            Height          =   540
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   840
            Width           =   4575
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   300
            Left            =   10320
            TabIndex        =   17
            Top             =   1560
            Width           =   1005
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            Height          =   300
            Left            =   5280
            TabIndex        =   15
            Top             =   1560
            Width           =   675
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile No."
            Height          =   300
            Left            =   240
            TabIndex        =   13
            Top             =   1560
            Width           =   1275
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PAN No."
            Height          =   300
            Left            =   20400
            TabIndex        =   11
            Top             =   480
            Width           =   1020
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GST No."
            Height          =   300
            Left            =   15360
            TabIndex        =   9
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Name"
            Height          =   300
            Left            =   10320
            TabIndex        =   7
            Top             =   480
            Width           =   1770
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier ID"
            Height          =   300
            Left            =   5280
            TabIndex        =   5
            Top             =   480
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice No."
            Height          =   300
            Left            =   240
            TabIndex        =   3
            Top             =   480
            Width           =   1350
         End
      End
   End
End
Attribute VB_Name = "purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function pur_mfgHeader()
pur_mfg.TextMatrix(0, 0) = " Product Name"
pur_mfg.TextMatrix(0, 1) = " HSN Code"
pur_mfg.TextMatrix(0, 2) = " CGST %"
pur_mfg.TextMatrix(0, 3) = " SGST %"
pur_mfg.TextMatrix(0, 4) = " QTY"
pur_mfg.TextMatrix(0, 5) = " Rate"
pur_mfg.TextMatrix(0, 6) = " CGST"
pur_mfg.TextMatrix(0, 7) = " SGST"
pur_mfg.TextMatrix(0, 8) = " Amount"

'set width
pur_mfg.ColWidth(0) = 7000
pur_mfg.ColWidth(1) = 2000
pur_mfg.ColWidth(2) = 2000
pur_mfg.ColWidth(3) = 2000
pur_mfg.ColWidth(4) = 1000
pur_mfg.ColWidth(5) = 2000
pur_mfg.ColWidth(6) = 2000
pur_mfg.ColWidth(7) = 2000
pur_mfg.ColWidth(8) = 3000

End Function

Private Sub Form_Load()
pur_mfgHeader
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
