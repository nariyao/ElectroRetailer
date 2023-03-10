VERSION 5.00
Begin VB.Form home 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Home"
   ClientHeight    =   11895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   21795
   LinkTopic       =   "home"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   793
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1453
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = main.ScaleWidth
Me.Height = main.ScaleHeight
End Sub


