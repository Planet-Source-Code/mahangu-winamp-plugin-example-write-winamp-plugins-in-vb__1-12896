VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Winamp Plugin Example in VB"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgWinamp 
      Height          =   1455
      Left            =   0
      Picture         =   "frmMain.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmMain.frx":2193
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()




End Sub

