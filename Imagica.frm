VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Imaginica 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Imanigica The Web Browser"
   ClientHeight    =   5250
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   7455
   Icon            =   "Imagica.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   4100
      Left            =   15
      TabIndex        =   20
      Top             =   960
      Width           =   6000
      ExtentX         =   10583
      ExtentY         =   7232
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Imagica"
      Height          =   255
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Paint"
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "MSIE"
      Height          =   255
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Computer"
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Notepad"
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Keyboard"
      Height          =   255
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Calculator"
      Height          =   255
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   720
      Width           =   855
   End
   Begin MSComDlg.CommonDialog comdlg 
      Left            =   7200
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Imagica -- Select a file to open"
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   720
      Picture         =   "Imagica.frx":628A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Add current page as bookmark."
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdadd 
      Height          =   495
      Left            =   1200
      Picture         =   "Imagica.frx":6DBC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Add current url to history"
      Top             =   120
      Width           =   495
   End
   Begin VB.ComboBox addrr 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   4800
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdfind 
      BackColor       =   &H00FFC0FF&
      Height          =   495
      Left            =   3960
      Picture         =   "Imagica.frx":79B6
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Search Google"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Height          =   495
      Left            =   4400
      Picture         =   "Imagica.frx":86DC
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Go !! --->"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdstop 
      BackColor       =   &H008080FF&
      Height          =   495
      Left            =   3360
      Picture         =   "Imagica.frx":91F2
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Stop loading the page."
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdfd 
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   2280
      Picture         =   "Imagica.frx":9F18
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Go one step forward."
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      Picture         =   "Imagica.frx":AC3E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Go one step back."
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdreferesh 
      BackColor       =   &H00FFFF80&
      Height          =   495
      Left            =   2880
      Picture         =   "Imagica.frx":B6A8
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Reload the page."
      Top             =   120
      Width           =   495
   End
   Begin ComctlLib.ProgressBar pro 
      Height          =   150
      Left            =   4800
      TabIndex        =   0
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   265
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label my_status 
      BackColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   6720
      TabIndex        =   19
      Top             =   720
      Width           =   735
   End
   Begin VB.Label tpv 
      BackColor       =   &H00FFC0C0&
      Caption         =   "100%"
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
      Left            =   0
      TabIndex        =   18
      Top             =   735
      Width           =   855
   End
   Begin VB.Label amM 
      Caption         =   """"
      Height          =   495
      Left            =   7920
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgcheck 
      Height          =   405
      Left            =   0
      Picture         =   "Imagica.frx":C3CE
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgdone 
      Height          =   495
      Left            =   0
      Picture         =   "Imagica.frx":CE9C
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu pref 
         Caption         =   "&Preference"
         Shortcut        =   ^G
      End
      Begin VB.Menu newnote 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu openhtml 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu addhistory 
         Caption         =   "Add to history"
      End
      Begin VB.Menu quit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu action 
      Caption         =   "&Action"
      Begin VB.Menu homepage 
         Caption         =   "&Go to home page"
      End
      Begin VB.Menu reloads 
         Caption         =   "&Reload"
      End
      Begin VB.Menu stops 
         Caption         =   "&Stop"
      End
      Begin VB.Menu backs 
         Caption         =   "Go &Back"
      End
      Begin VB.Menu forwards 
         Caption         =   "Go &Forward"
      End
   End
   Begin VB.Menu helping 
      Caption         =   "&Help"
      Begin VB.Menu softhome 
         Caption         =   "&Go to Home Page"
      End
      Begin VB.Menu versions 
         Caption         =   "&Check for new Version"
      End
      Begin VB.Menu about 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu bookMenu 
      Caption         =   "&Bookmark"
      Begin VB.Menu addbook 
         Caption         =   "Add Bookmark"
      End
      Begin VB.Menu sep 
         Caption         =   "-----------------------"
      End
      Begin VB.Menu myBook 
         Caption         =   "My bookmark"
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "Imaginica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
On Error Resume Next
frmAbout.Visible = True
End Sub

Private Sub addbook_Click()
On Error Resume Next
Command2_Click
End Sub



Private Sub addhistory_Click()
On Error Resume Next
cmdadd_Click
End Sub
Private Sub addrr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If InStr(1, addrr.Text, " ") <> 0 Then
web.Navigate "http://www.google.co.in/search?q=" & addrr.Text & "&sourceid=imagica"
Else
web.Navigate addrr.Text
addrr.AddItem addrr.Text
Dim mbytes As Single, ff As Integer
ff = FreeFile
ff = FreeFile
Open "history.dat" For Binary As ff
mbytes = Seek(ff)
Put ff, , addrr.Text & vbCrLf
Close ff
End If
End If
If Shift = vbCtrlMask Then
If KeyCode = 13 Then my_nav
End If
End Sub

Private Sub backs_Click()
On Error Resume Next
web.GoBack
End Sub

Private Sub cmdadd_Click()
On Error Resume Next
addrr.AddItem web.LocationURL
Dim tot As Integer
Dim totalUrl As Integer
Dim m As Integer
If web.LocationURL = "" Then Exit Sub
totalUrl = GetSetting("MSOFT", "IMAGICA", "URL", 0)
tot = GetSetting("MSOFT", "IMAGICA", "URL", 0)
For m = 1 To totalUrl
If GetSetting("MSOFT", "IMAGICA", "URL" & m, "") = web.LocationURL Then Exit Sub
Next m
SaveSetting "MSOFT", "IMAGICA", "URL" & tot + 1, web.LocationURL
SaveSetting "MSOFT", "IMAGICA", "URL", tot + 1
End Sub

Private Sub cmdback_Click()
On Error Resume Next
web.GoBack
End Sub

Private Sub cmdfd_Click()
On Error Resume Next
web.GoForward
End Sub

Private Sub cmdfind_Click()
On Error Resume Next
If addrr.Text <> "" Then
    web.Navigate "http://www.google.co.in/search?q=" & addrr.Text & "&sourceid=imagica"
    End If
    If addrr.Text = "" Then web.Navigate "http://www.google.com"
End Sub
Private Function WebPageContains(ByVal s As String) As Boolean
On Error Resume Next
    Dim i As Long, EHTML
    For i = 1 To web.Document.All.length
        Set EHTML = _
        web.Document.All.Item(i)


        If Not (EHTML Is Nothing) Then
            If InStr(1, EHTML.innerHTML, _
            s, vbTextCompare) > 0 Then
            WebPageContains = True
            Exit Function
        End If
    End If
Next i
End Function

Private Sub cmdreferesh_Click()
On Error Resume Next
web.Refresh
End Sub

Private Sub cmdstop_Click()
On Error Resume Next
web.Stop
imgdone.Visible = flase
imgcheck.Visible = True
End Sub

Private Sub Command1_Click()
On Error Resume Next
web.Navigate addrr
Dim mbytes As Single, ff As Integer
ff = FreeFile
ff = FreeFile
Open "history.dat" For Binary As ff
mbytes = Seek(ff)
Put ff, , addrr.Text & vbCrLf
Close ff
End Sub

Private Sub Command10_Click()
On Error Resume Next
openn.Visible = True
End Sub

Private Sub Command11_Click()

End Sub

Private Sub Command2_Click()
On Error Resume Next
'--------------------------------------------
Dim tt As Integer
Load myBook(tt)
myBook(tt).Caption = web.LocationName
myBook(tt).Visible = True
If web.LocationURL = "" Then Exit Sub
Dim tot As Integer
tot = GetSetting("MSOFT", "IMAGICA", "BOOKMARK", 0)
SaveSetting "MSOFT", "IMAGICA", "BOOKMARK" & tot + 1, web.LocationURL
SaveSetting "MSOFT", "IMAGICA", "BOOKNAME" & tot + 1, web.LocationName
SaveSetting "MSOFT", "IMAGICA", "BOOKMARK", tot + 1
End Sub

Private Sub Command3_Click()
On Error GoTo rr
Shell "calc.exe"
GoTo exx
rr: MsgBox "Calculator not found. Goto help -> homepage and download it."
exx: End Sub

Private Sub Command4_Click()
On Error GoTo rr
Shell "osk.exe", vbNormalFocus
GoTo exx
rr: MsgBox "Keyboard not found. Goto help -> homepage and download it."
exx: End Sub

Private Sub Command5_Click()
On Error GoTo rr
Shell "notepad.exe", vbNormalFocus
GoTo exx
rr: MsgBox "Notepad not found. Goto help -> homepage and download it."
exx: End Sub

Private Sub Command6_Click()
On Error Resume Next
a = amM.Caption
Shell "explorer.exe " & a & "" & a, vbNormalFocus
GoTo exx
rr: MsgBox "Calculator not found. Goto help -> homepage and download it."
exx: End Sub

Private Sub Command7_Click()
On Error Resume Next
a = amM.Caption
Shell "explorer.exe " & a & "about:blank" & a, vbNormalFocus
GoTo exx
exx: End Sub

Private Sub Command8_Click()
On Error GoTo rr
Shell "mspaint.exe", vbNormalFocus
GoTo exx
rr: MsgBox "MS Paint not found. Goto help -> homepage and download it."
exx: End Sub

Private Sub Command9_Click()
On Error Resume Next
Shell App.EXEName, vbNormalFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
my_status.Width = Me.Width - 6869
addrr.Text = ""
addrr.Width = Me.Width - 4920
web.Width = Me.Width - 120
web.Height = Me.Height - 1800
pro.Width = Me.Width - 5040
'----------------------------------------------------------------------
Bookmark_Menu
generateHistory
On Error Resume Next
Dim totalUrl As Integer
Dim i As Integer
totalUrl = GetSetting("MSOFT", "IMAGICA", "URL", 0)
urls = GetSetting("MSOFT", "IMAGICA", "URL" & i, 0)
If totalUrl = 0 Then Exit Sub
For i = 1 To totalUrl
addrr.AddItem GetSetting("MSOFT", "IMAGICA", "URL" & i, "http://www.masoomyf.tk")
Next i
End Sub

Private Sub Form_Resize()

On Error Resume Next
my_status.Width = Me.Width - 6869
addrr.Width = Me.Width - 4920
web.Width = Me.Width - 120
web.Height = Me.Height - 1800
pro.Width = Me.Width - 5040
End Sub


Private Sub forwards_Click()
On Error Resume Next
web.GoForward
End Sub

Private Sub homepage_Click()
On Error Resume Next
web.Navigate GetSetting("MSOFT", "Imagica", "homepage", "www.google.com")
End Sub

Private Sub Label1_Click()

End Sub

Private Sub myBook_Click(Index As Integer)
On Error Resume Next
Dim a  As Integer
a = Index
Dim bookUrl As String
bookUrl = GetSetting("MSOFT", "IMAGICA", "BOOKMARK" & a, "www.google.com")
web.Navigate bookUrl
End Sub

Private Sub newnote_Click()
On Error Resume Next
Shell App.EXEName
End Sub

Private Sub openhtml_Click()
On Error GoTo mmm
Dim openUrl As String
comdlg.Filter = "(All supported format)"
comdlg.ShowOpen
openUrl = comdlg.FileName
If openUrl <> "" Then
web.Navigate openUrl
comdlg.FileName = ""
openUrl = ""
End If
GoTo eee
mmm: MsgBox "File format not supported"
MsgBox Error
eee: End Sub

Private Sub pref_Click()
On Error Resume Next
preff.Visible = True
End Sub

Private Sub quit_Click()
On Error Resume Next
End
End Sub


Private Sub reloads_Click()
On Error Resume Next
web.Refresh
End Sub



Private Sub stops_Click()
On Error Resume Next
web.Stop
End Sub



Private Sub versions_Click()
MsgBox "Update is not available yet. Because this is new one."
End Sub

Private Sub web_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Me.Caption = "Imagica -- " & web.LocationName
addrr.Text = web.LocationURL
End Sub

Private Sub web_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
Me.Caption = "Imagica -- " & web.LocationName
addrr.Text = web.LocationURL
End Sub

Private Sub web_NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
On Error Resume Next
Me.Caption = "Imagica -- Error"
End Sub

Private Sub web_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
Dim aaa As Integer
If Progress = -1 Then
imgdone.Visible = False
imgcheck.Visible = True
tpv.Caption = "100%"
Exit Sub
End If
    If Progress > 0 And ProgressMax > 0 Then

        pro.Visible = True
        pro.Value = Progress / ProgressMax * 100
        aaa = Progress / ProgressMax * 100
         tpv.Caption = aaa & "%"
        If aaa <> 100 Then imgdone.Visible = True
         If aaa <> 100 Then imgcheck.Visible = flase
        End If
End Sub
'This to open a new window with our browser.
Private Sub Web_NewWindow2(ppDisp As Object, Cancel As Boolean)
On Error Resume Next
Dim frm As Imaginica
Set frm = New Imaginica
Set ppDisp = frm.web.Object
frm.Show
End Sub
Sub Bookmark_Menu()
On Error Resume Next
Dim totalUrl As Integer
Dim u As Integer
totalUrl = GetSetting("MSOFT", "IMAGICA", "BOOKMARK", 0)
If totalUrl = 0 Then Exit Sub
For u = 1 To totalUrl
Load myBook(u)
myBook(u).Caption = GetSetting("MSOFT", "IMAGICA", "BOOKNAME" & u, "")
myBook(u).Visible = True
Next u
End Sub
Sub my_nav()
On Error Resume Next
web.Navigate "http://www." & addrr.Text & ".com"
addrr.Text = "http://www." & addrr.Text & ".com"
End Sub

Private Sub web_StatusTextChange(ByVal Text As String)
my_status.Caption = Text
End Sub
Private Sub generateHistory()
Dim ff As Integer, mStri As String
ff = FreeFile
Open "history.dat" For Binary As ff
mStri = Space(LOF(ff))
Get ff, , mStri
Close ff
For i = 1 To Len(mStri)
a = InStr(i, mStri, vbCrLf)
If a = 0 Then Exit Sub
addrr.AddItem Split(mStri, vbCrLf)(0)
i = a
Next i
End Sub
