VERSION 5.00
Begin VB.Form preff 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Imagica - Settings"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6450
   Icon            =   "pref.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1620
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Google"
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Home Page"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Blank"
         Height          =   255
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox homeurl 
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   210
         Width           =   5415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "URL :"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "preff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error Resume Next
homeurl.Text = "about:blank"
SaveSetting "MSOFT", "Imagica", "homepage", "about:blank"
End Sub

Private Sub Command2_Click()
On Error Resume Next
homeurl.Text = "http://www.google.com"
SaveSetting "MSOFT", "Imagica", "homepage", "http://www.google.com"
End Sub


Private Sub Form_Load()
On Error Resume Next
homeurl.Text = GetSetting("MSOFT", "Imagica", "homepage", "http://www.google.com")
End Sub

Private Sub homeurl_Change()
On Error Resume Next
SaveSetting "MSOFT", "Imagica", "homepage", homeurl.Text
End Sub
