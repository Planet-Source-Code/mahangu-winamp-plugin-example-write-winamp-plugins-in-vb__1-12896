VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Winamp Plugin Example in VB [Config]"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblInfo 
      Caption         =   "This is the configuartion window of your Winamp plugin. You can use this window to let the user change / customize your plugin."
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Image imgWinamp 
      Height          =   2340
      Left            =   120
      Picture         =   "frmConfig.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4440
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()





End Sub
