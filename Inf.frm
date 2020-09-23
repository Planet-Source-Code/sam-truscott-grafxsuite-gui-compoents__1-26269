VERSION 5.00
Begin VB.Form Inf 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "GrafxSuite"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "If you have any problems or comments please email them to: samst@btinternet.com"
      Height          =   855
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Grafxsuite, created and designed by KSU 2001, lead programmer Sam Truscott."
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1200
      Left            =   0
      Picture         =   "Inf.frx":0000
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "Inf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
