VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form progress 
   BorderStyle     =   0  'None
   Caption         =   "Creating DBA and related Tables"
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1085
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      OLEDropMode     =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label percent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   195
      Left            =   6480
      TabIndex        =   2
      Top             =   600
      Width           =   210
   End
   Begin VB.Label msg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   600
   End
End
Attribute VB_Name = "progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prg_msg As String
Public Function database_creation()
Dim sql As String
On Error GoTo ER1:
sql = connection(USERID, PASSWORD)
prg ("Checking for user elctro")
C.Execute ("DROP USER electro CASCADE")
msg.Caption = "user electro droped"
CRTUSER:
prg ("Creating user electro")
C.Execute ("CREATE USER electro IDENTIFIED BY retailer")

C.Execute ("GRANT DBA TO electro")
C.Close
prg ("Connecting user electro")
sql = connection("electro", "retailer")
msg.Caption = "Connected"
'BELOW TABLE CREATION
'COMPANY DETAILS
prg ("Creating company table")
C.Execute ("CREATE TABLE ER_MASTER_COMPANY_DETAILS(NAME VARCHAR(30) NOT NULL,COMPANY_NAME VARCHAR(50) NOT NULL, BRANCH_NAME VARCHAR(30),GSTNO CHAR(15) NOT NULL,PAN CHAR(10) NOT NULL, MOBILE_NO CHAR(10) NOT NULL,EMAIL VARCHAR(50), DESCRIPTION VARCHAR(100),CONSTRAINT PK_B_N PRIMARY KEY(BRANCH_NAME))")
'CREATINT ACCOUNT
prg ("Creating accounts table")
C.Execute ("CREATE TABLE ER_MASTER_ACCOUNT(ACC_NO CHAR(5), BALANCE DECIMAL(15,2) DEFAULT 0, LOAN DECIMAL(15,2) DEFAULT 0, MONTH CHAR(3) NOT NULL, INCOME DECIMAL(15,2) DEFAULT 0, EXPENSE DECIMAL(15,2) DEFAULT 0, CONSTRAINT PK_ACC_NO PRIMARY KEY(ACC_NO))")
'CREATE TRANSACTION DETAILS
prg ("Creating transaction table")
C.Execute ("CREATE TABLE ER_MASTER_TRANSACTION(T_ID VARCHAR(50),T_DATE CHAR(10) NOT NULL,T_TIME CHAR(8) NOT NULL,T_PARTICULAR VARCHAR(30) NOT NULL, T_CR_DR CHAR(2) NOT NULL, T_MODE VARCHAR(10) NOT NULL, T_AMOUNT DECIMAL(15,2) DEFAULT 0,CONSTRAINT PK_T_ID PRIMARY KEY(T_ID))")
'CREATING LOGIN TABLE
prg ("Creating login table")
C.Execute ("CREATE TABLE ER_MASTER_LOGIN(USERID CHAR(20),PASSWORD VARCHAR(20) NOT NULL, ACCESS_LEVEL CHAR(2) NOT NULL, CONSTRAINT PK_USERID PRIMARY KEY(USERID))")
'CREATE EMPLOYEE TABLE
prg ("Creating employee table")
C.Execute ("CREATE TABLE ER_MASTER_EMPLOYEE(E_ID CHAR(10),E_NAME VARCHAR(30) NOT NULL,E_FNAME VARCHAR(30)NOT NULL, E_DOB CHAR(10) NOT NULL,E_GENDER CHAR(1) NOT NULL, E_MOB CHAR(10)NOT NULL,e_MAIL CHAR(30), E_ADHAAR CHAR(16) NOT NULL,E_DOJ CHAR(10) NOT NULL,E_ADD VARCHAR(100)NOT NULL,E_QUL VARCHAR(30) NOT NULL,E_EXP CHAR(2), E_POST VARCHAR(30) NOT NULL,E_LEAVE CHAR(1), E_SALARY DECIMAL(5) DEFAULT 0,CONSTRAINT PK_E_ID PRIMARY KEY(E_ID))")
'CREATE EMPLOYEE ATTENDANCE
prg ("Creating employee table")
C.Execute ("CREATE TABLE ER_SUB_ATTENDANCE(E_ID CHAR(10),MONTH CHAR(3) NOT NULL, PRESENT NUMBER(2) NOT NULL, LEAVE NUMBER(2) NOT NULL, CONSTRAINT FK_EA_E_ID FOREIGN KEY(E_ID) REFERENCES ER_MASTER_EMPLOYEE(E_ID))")
'CREATE SUPPLIER
prg ("Creating supplier table")
C.Execute ("CREATE TABLE ER_MASTER_SUPPLIER(S_ID CHAR(10),S_NAME VARCHAR(30) NOT NULL,COMPANY_NAME VARCHAR(50) NOT NULL,S_EMAIL VARCHAR(30) NOT NULL,S_MOBILE CHAR(10) NOT NULL,S_GSTNO CHAR(15) NOT NULL, S_PAN CHAR(10),S_ADDRESS VARCHAR(100) NOT NULL,S_PINCODE CHAR(6),CONSTRAINT PK_S_ID PRIMARY KEY(S_ID))")
'CREATE PRODUCT
prg ("Creating product table")
C.Execute ("CREATE TABLE ER_MASTER_PRODUCT(P_ID VARCHAR(15),P_NAME VARCHAR(100) NOT NULL,P_COMPANY VARCHAR(50) NOT NULL, P_MODEL VARCHAR(20),HSN CHAR(8),P_GST DECIMAL(4,2), P_QTY NUMBER(2) DEFAULT 0, S_ID CHAR(10),CONSTRAINT PK_P_ID PRIMARY KEY(P_ID), CONSTRAINT FK_P_S_ID FOREIGN REFERENCES ER_MASTER_SUPPLIER(S_ID))")
'CREATE PURCHASE INVOICE
prg ("Creating purchase invoice table")
C.Execute ("CREATE TABLE ER_MASTER_PURCHASE_INVOICE(PI_ID VARCHAR(20), PI_DATE_TIME VARCHAR(20) NOT NULL, S_ID CHAR(10), CONSTRAINT FK_PI_S_ID FOREIGN KEY(S_ID) REFERENCES ER_MASTER_SUPPLIER(S_ID),CONSTRAINT PK_PI_ID PRIMARY KEY(PI_ID))")
'CREATE ORDERED PRODUCT
prg ("Creating orderer product table")
C.Execute ("CREATE TABLE ER_SUB_ORDERED_PRODUCT(PI_ID VARCHAR(20),P_ID VARCHAR(15),OP_RATE DECIMAL(7,2) NOT NULL,OP_QTY NUMBER(2) NOT NULL,OP_STATUS VARCHAR(15),CONSTRAINT FK_OP_PI_ID FOREIGN KEY(PI_ID) REFERENCES ER_MASTER_PURCHASE_INVOICE(PI_ID), CONSTRAINT FK_OP_P_ID FOREIGN KEY(P_ID) REFERENCES ER_MASTER_PRODUCT(P_ID))")
'CREATE CUSTOMER
prg ("Creating customer table")
C.Execute ("CREATE TABLE ER_MASTER_CUSTOMER(C_ID VARCHAR(10),C_NAME VARCHAR(30) NOT NULL, C_MOBILE CHAR(10) UNIQUE NOT NULL, C_EMAIL VARCHAR(50),C_ADDR VARCHAR(100) NOT NULL,CONSTRAINT PK_C_ID PRIMARY KEY(C_ID))")
'CREATE SELL INVOICE
prg ("Creating sell table")
C.Execute ("CREATE TABLE ER_MASTER_SELL_INVOICE(SI_ID VARCHAR(30), C_ID VARCHAR(10) ,T_ID VARCHAR(50),SI_DATE_TIME VARCHAR(20) NOT NULL, CONSTRAINT FK_SI_C_ID FOREIGN KEY(C_ID) REFERENCES ER_MASTER_CUSTOMER(C_ID), CONSTRAINT FK_SI_T_ID FOREIGN KEY(T_ID) REFERENCES ER_MASTER_TRANSACTION(T_ID),CONSTRAINT PK_SI_ID PRIMARY KEY(SI_ID))")
'CREATE SOLD PRODUCT
prg ("Creating sold product table")
C.Execute ("CREATE TABLE ER_SUB_SOLD_PRODUCT(SI_ID VARCHAR(30),P_ID VARCHAR(15),SP_QTY NUMBER(2) NOT NULL,SP_RATE DECIMAL(7,2) NOT NULL, CONSTRAINT FK_SP_SI_ID FOREIGN KEY(SI_ID) REFERENCES ER_MASTER_SELL_INVOICE(SI_ID), CONSTRAINT FK_SP_P_ID FOREIGN KEY(P_ID) REFERENCES ER_MASTER_PRODUCT(P_ID))")
'CREATE SALE RETURN INVOICE
prg ("Creating sale return invoice table")
C.Execute ("CREATE TABLE ER_MASTER_SALE_RETURN_INVOICE(SRI_ID VARCHAR(30),C_ID VARCHAR(10),T_ID VARCHAR(50),SRI_DATE_TIME VARCHAR(20) NOT NULL,CONSTRAINT PK_SRI_ID PRIMARY KEY(SRI_ID), CONSTRAINT FK_SRI_C_ID FOREIGN KEY(C_ID) REFERENCES ER_MASTER_CUSTOMER(C_ID), CONSTRAINT FK_SRI_T_ID FOREIGN KEY(T_ID) REFERENCES ER_MASTER_TRANSACTION(T_ID))")
'CREATE SALE RETURN
prg ("Creating sale return table")
C.Execute ("CREATE TABLE ER_SUB_SALE_RETURN(SRI_ID VARCHAR(30),P_ID VARCHAR(15),SR_QTY NUMBER(2) NOT NULL,SR_RATE DECIMAL(7,2) NOT NULL,SR_NOTE VARCHAR(50) NOT NULL, CONSTRAINT FK_SR_SRI_ID FOREIGN KEY(SRI_ID) REFERENCES ER_MASTER_SALE_RETURN_INVOICE(SRI_ID), CONSTRAINT FK_SR_P_ID FOREIGN KEY(P_ID) REFERENCES ER_MASTER_PRODUCT(P_ID))")
'CREATE PURCHASE RETURN INVOICE
prg ("Creating purchase return invoice table")
C.Execute ("CREATE TABLE ER_MASTER_PUR_RETURN_INVOICE(PRI_ID VARCHAR(30),S_ID CHAR(10),T_ID VARCHAR(50),PRI_DATE_TIME VARCHAR(20) NOT NULL, CONSTRAINT FK_PRI_S_ID FOREIGN KEY(S_ID) REFERENCES ER_MASTER_SUPPLIER(S_ID), CONSTRAINT FK_PRI_T_ID FOREIGN KEY(T_ID) REFERENCES ER_MASTER_TRANSACTION(T_ID), CONSTRAINT PK_PRI_ID PRIMARY KEY(PRI_ID))")
'CREATE PURCHASE RETURN
prg ("Creating purchase return table")
C.Execute ("CREATE TABLE ER_SUB_PURCHASE_RETURN(PRI_ID VARCHAR(30),P_ID VARCHAR(15),PR_QTY NUMBER(2) NOT NULL,PR_RATE DECIMAL(7,2) NOT NULL,PR_NOTE VARCHAR(50) NOT NULL, CONSTRAINT FR_PR_PRI_ID FOREIGN KEY(PRI_ID) REFERENCES ER_MASTER_PUR_RETURN_INVOICE(PRI_ID), CONSTRAINT FK_PR_P_ID FOREIGN KEY(P_ID) REFERENCES ER_MASTER_PRODUCT(P_ID))")
' CREATE INDIRECT EXPENSES
prg ("Creating indirect expenses table")
C.Execute ("CREATE TABLE ER_MASTER_INDIRECT_EXPENSES(YEAR CHAR(4),MONTH CHAR(3),SHOP_RENT DECIMAL(7,2) DEFAULT 0, SALARY DECIMAL(8,2) DEFAULT 0,ACC_CHR DECIMAL(7,2) DEFAULT 0,ELECTRIC_CHR DECIMAL(7,2) DEFAULT 0,TRAVEL DECIMAL(7,2) DEFAULT 0,AUDIT_FEE DECIMAL(7,2) DEFAULT 0,PRINTING_STNR DECIMAL(7,2) DEFAULT 0,LEGAL_EXP DECIMAL(7,2) DEFAULT 0,POS DECIMAL(7,2) DEFAULT 0,MIS_CHR DECIMAL(7,2) DEFAULT 0,MOBILE DECIMAL(6,2) DEFAULT 0,DEPRECIATION DECIMAL(7,2) DEFAULT 0,BANK_CHR DECIMAL(6,2) DEFAULT 0,CONSTRAINT PK_IE_YR_MON PRIMARY KEY(YEAR,MONTH))")
prg ("Done")
pb:
prg_msg = MsgBox("DBA and database are created", vbOKOnly)
If prg_msg = vbOK Then
Unload Me
db.cau_fr.Enabled = True
db.cau_user.SetFocus
End If
Exit Function
ER1:
If Err.Number = -2147467259 Then
GoTo CRTUSER:
ElseIf Err.Number = 380 Then
GoTo pb:
Else
MsgBox Err.Description
End If
End Function

Private Sub Form_Activate()
ProgressBar.Value = 0
msg.Caption = "Connecting to DBA"
database_creation
End Sub

Private Function prg(ByVal lb As String)
msg.Caption = lb
If lb = "Done" Then
ProgressBar.Value = 100
Else
ProgressBar.Value = ProgressBar.Value + 4
End If
percent = Str(ProgressBar.Value) + "%"
End Function
