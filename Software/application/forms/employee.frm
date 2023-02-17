VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form emp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   13920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20700
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   13920
   ScaleWidth      =   20700
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   13095
      Left            =   1560
      TabIndex        =   17
      Top             =   600
      Width           =   18615
      Begin VB.ComboBox quali 
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
         Left            =   13680
         TabIndex        =   38
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   11040
         TabIndex        =   33
         Top             =   4560
         Width           =   2175
         Begin VB.OptionButton T 
            Caption         =   "T"
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
            Left            =   1680
            TabIndex        =   36
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton F 
            Caption         =   "F"
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
            Left            =   840
            TabIndex        =   35
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton M 
            Caption         =   "M"
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
            Left            =   0
            TabIndex        =   34
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.ComboBox post 
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
         Left            =   840
         TabIndex        =   10
         Top             =   6720
         Width           =   4215
      End
      Begin VB.PictureBox Picture1 
         Height          =   2415
         Left            =   13680
         ScaleHeight     =   2355
         ScaleWidth      =   2115
         TabIndex        =   2
         Top             =   1680
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker doj 
         Height          =   375
         Left            =   11040
         TabIndex        =   1
         Top             =   2040
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
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
         Format          =   144244737
         CurrentDate     =   44955
      End
      Begin VB.TextBox id 
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
         Left            =   840
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   30
         Top             =   2040
         Width           =   4215
      End
      Begin VB.TextBox salary 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00;(#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   1
         EndProperty
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
         Left            =   11040
         MaxLength       =   5
         TabIndex        =   12
         Top             =   6720
         Width           =   4935
      End
      Begin VB.TextBox leaves 
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
         Left            =   5400
         MaxLength       =   1
         TabIndex        =   11
         Top             =   6720
         Width           =   5175
      End
      Begin VB.TextBox expr 
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
         Left            =   11040
         MaxLength       =   2
         TabIndex        =   9
         Top             =   5640
         Width           =   4935
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
         Height          =   420
         Left            =   840
         MaxLength       =   100
         TabIndex        =   8
         Top             =   5640
         Width           =   9735
      End
      Begin VB.TextBox adhaar 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   1
         EndProperty
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
         Left            =   5400
         MaxLength       =   16
         TabIndex        =   7
         Top             =   4560
         Width           =   5175
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
         Height          =   420
         Left            =   840
         MaxLength       =   30
         TabIndex        =   6
         Top             =   4560
         Width           =   4215
      End
      Begin VB.TextBox mobile 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   1
         EndProperty
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
         Left            =   5400
         MaxLength       =   10
         TabIndex        =   4
         Top             =   3240
         Width           =   5175
      End
      Begin VB.TextBox fname 
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
         Left            =   840
         MaxLength       =   30
         TabIndex        =   3
         Top             =   3240
         Width           =   4215
      End
      Begin VB.TextBox emp_name 
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
         Left            =   5400
         MaxLength       =   30
         TabIndex        =   0
         Top             =   2040
         Width           =   5175
      End
      Begin MSComCtl2.DTPicker dob 
         Height          =   375
         Left            =   11040
         TabIndex        =   5
         Top             =   3240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
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
         Format          =   143785985
         CurrentDate     =   44955
      End
      Begin VB.Label emp_msg_lb 
         AutoSize        =   -1  'True
         Caption         =   "Label16"
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
         Left            =   960
         TabIndex        =   39
         Top             =   7680
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   3
         EndProperty
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
         Left            =   11040
         TabIndex        =   37
         Top             =   4200
         Width           =   780
      End
      Begin VB.Label delete_btn 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "                       Delete"
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
         Left            =   11880
         TabIndex        =   15
         Top             =   7680
         Width           =   1935
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
         Left            =   9720
         TabIndex        =   14
         Top             =   7680
         Width           =   1935
      End
      Begin VB.Label add_btn 
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
         Left            =   11880
         TabIndex        =   13
         Top             =   8400
         Width           =   1935
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
         Height          =   615
         Left            =   14040
         TabIndex        =   16
         Top             =   7680
         Width           =   1935
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Joining"
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
         Left            =   11040
         TabIndex        =   32
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
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
         Left            =   840
         TabIndex        =   31
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of  Birth"
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
         Left            =   11040
         TabIndex        =   29
         Top             =   2880
         Width           =   1350
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary"
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
         Left            =   11040
         TabIndex        =   28
         Top             =   6360
         Width           =   690
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leaves"
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
         Left            =   5400
         TabIndex        =   27
         Top             =   6360
         Width           =   780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Post"
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
         Left            =   840
         TabIndex        =   26
         Top             =   6360
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exprience"
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
         Left            =   11040
         TabIndex        =   25
         Top             =   5280
         Width           =   1290
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qualification"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   3
         EndProperty
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
         Left            =   13680
         TabIndex        =   24
         Top             =   4200
         Width           =   1305
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   840
         TabIndex        =   23
         Top             =   5280
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adhaar Card"
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
         Left            =   5400
         TabIndex        =   22
         Top             =   4200
         Width           =   1320
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
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
         Left            =   840
         TabIndex        =   21
         Top             =   4200
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moblie Number"
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
         Left            =   5400
         TabIndex        =   20
         Top             =   2880
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Father's Name"
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
         Left            =   840
         TabIndex        =   19
         Top             =   2880
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   5400
         TabIndex        =   18
         Top             =   1680
         Width           =   630
      End
   End
End
Attribute VB_Name = "emp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Insert Button
Private Sub add_btn_Click()
On Error GoTo EMP_ER_INS
connectOracle
sql = "INSERT INTO ER_MASTER_EMPLOYEE VALUES('" + id.Text + "','" + emp_name.Text + "','" + fname.Text + "','" + Trim(Str(dob.Value)) + "','" + getGender + "','" + mobile.Text + "','" + email.Text + "','" + adhaar.Text + "','" + Trim(Str(doj.Value)) + "','" + addr.Text + "','" + quali.Text + "','" + expr.Text + "','" + post.Text + "','" + leaves.Text + "'," + salary.Text + ")"
C.Execute (sql)
C.Execute ("commit")
C.Close
emp_msg_lb.Caption = "Succes! Employee Data Uploaded"
emp_msg_lb.Visible = True
emp_msg_lb.ForeColor = green
clearEmp
enableAddBtn
Exit Sub
EMP_ER_INS:
empErr
End Sub

'update employee details
Private Sub update_btn_Click()
Dim sql As String
sql = "UPDATE ER_MASTER_EMPLOYEE SET E_NAME='" + emp_name.Text + "', E_FNAME='" + fname.Text + "',E_DOB='" + Trim(Str(dob.Value)) + "',E_GENDER='" + getGender + "',E_MOB='" + mobile.Text + "',E_MAIL='" + email.Text + "',E_ADHAAR='" + adhaar.Text + "',E_DOJ='" + Trim(Str(doj.Value)) + "',E_ADD='" + addr.Text + "',E_QUL='" + quali.Text + "',E_EXP='" + expr.Text + "',E_POST='" + post.Text + "',E_LEAVE='" + leaves.Text + "',E_SALARY=" + salary.Text + " WHERE E_id='" + id.Text + "'"
connectOracle
C.Execute (sql)
C.Execute ("commit")
C.Close
temp = empMsg("S", "Updated")
End Sub

'Delete Button
Private Sub delete_btn_Click()
temp = MsgBox("it is going to deleted permanently", vbExclamation + vbOKCancel, "Deleting Employee Details")
If temp = vbOK Then
On Error GoTo EMP_ER_DEL
sql = "DELETE FROM ER_MASTER_EMPLOYEE WHERE e_id='" + id.Text + "'"
connectOracle
C.Execute (sql)
C.Execute ("commit")
C.Close
MsgBox "Employee ID deleted as well as its related data"
Unload Me
clearEmp
displayEmp
End If
Exit Sub
EMP_ER_DEL:
temp = empMsg("E", "")
End Sub

'Close button
Private Sub close_btn_Click()
displayEmp
Unload Me
clearEmp
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = main.ScaleWidth
Me.Height = main.ScaleHeight
add_btn.Top = 7680
defaultEmpValues
End Sub

'Default value for form
Private Function defaultEmpValues()
quali.AddItem "10th Pass"
quali.AddItem "12th Pass"
quali.AddItem "Undergraduate"
quali.AddItem "Graduate"
quali.AddItem "Post Graduate"
post.AddItem "Manager"
post.AddItem "Worker"
post.AddItem "Salesman"
End Function

'clear textbox
Private Function clearEmp()
id.Text = generateEmpId
emp_name.Text = ""
fname.Text = ""
mobile.Text = ""
email.Text = ""
adhaar.Text = ""
addr.Text = ""
quali.Text = ""
expr.Text = ""
post.Text = ""
leaves.Text = ""
salary.Text = ""
emp_msg_lb.Visible = False
End Function

'Add button enable or disable
Private Function enableAddBtn()
If emp_name.Text = "" And fname.Text = "" And mobile.Text = "" And email.Text = "" And adhaar.Text = "" And addr.Text = "" And quali.Text = "" And expr.Text = "" And post.Text = "" And leaves.Text = "" And salary.Text = "" Then
add_btn.Enabled = False
Else
add_btn.Enabled = True
End If
End Function


'this is for display msg
Private Function empMsg(ByVal msg_type As String, ByVal msg As String)
Select Case msg_type
Case "S":   emp_msg_lb.Caption = "Succes! " + msg
            emp_msg_lb.Visible = True
            emp_msg_lb.ForeColor = vbGreen
Case "E":   emp_msg_lb.Caption = "Error! " + Str(Err.Number) + " : " + Err.Description
            emp_msg_lb.Visible = True
            emp_msg_lb.ForeColor = vbRed
End Select
End Function
