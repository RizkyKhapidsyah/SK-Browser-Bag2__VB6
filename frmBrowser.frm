VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmBrowser 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Internet Looker - Browser"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   FillColor       =   &H0080FF80&
   ForeColor       =   &H0000FFFF&
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   7965
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Picture         =   "frmBrowser.frx":0442
            TextSave        =   "1:47 PM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Picture         =   "frmBrowser.frx":089E
            Text            =   "Click For Calender"
            TextSave        =   "Click For Calender"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bradley Hand ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdWindow 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      Picture         =   "frmBrowser.frx":0CFA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdTextEd 
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      Picture         =   "frmBrowser.frx":113C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      Picture         =   "frmBrowser.frx":157E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      DownPicture     =   "frmBrowser.frx":19C0
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      Picture         =   "frmBrowser.frx":1E02
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      Picture         =   "frmBrowser.frx":2244
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   "Forward"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      Picture         =   "frmBrowser.frx":2686
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      Picture         =   "frmBrowser.frx":2AC8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Go Get It"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "frmBrowser.frx":2F0A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin SHDocVwCtl.WebBrowser wbBrowser 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   11895
      ExtentX         =   20981
      ExtentY         =   12091
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
   Begin VB.Label lblNames 
      BackColor       =   &H000000FF&
      Caption         =   "Your Name/Info/Pic/Buttons"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   9
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
'Goes back to previous page
On Error GoTo errorhandler:
wbBrowser.GoBack

errorhandler:
'Error handler is blank causing it to do nothing/closing/stopping error/Nothing to go back to"
End Sub

Private Sub cmdForward_Click()
'Goes forward to previous page
On Error GoTo errorhandler:
wbBrowser.GoForward

errorhandler:
'Error handler is blank causing it to do nothing/stop closing/stopping error/Nothing to go back to"
End Sub
Private Sub cmdHome_Click()
'Returns to home web page
wbBrowser.GoHome
End Sub
Private Sub cmdRefresh_Click()
'Refreshes web page
wbBrowser.Refresh2
End Sub
Private Sub cmdSearch_Click()
Dim strSearch1 As String

On Error GoTo errorhandler
frmAddress.Text1.SetFocus
strSearch1 = frmAddress.Text1.Text
Unload frmAddress
wbBrowser.Navigate strSearch1

'If Textbox is reurn to homepage
errorhandler:
 ' wbBrowser.GoHome
  
 
End Sub

Private Sub cmdStop_Click()
wbBrowser.Stop
End Sub

Private Sub cmdTextEd_Click()
Form1.Show
End Sub

Private Sub cmdWindow_Click()
frmAddress.Show
End Sub

Private Sub Form_Load()
'Default website
wbBrowser.GoHome
StatusBar1.Panels(1).AutoSize = sbrContents
StatusBar1.Panels(2).AutoSize = sbrContents

End Sub


Private Sub StatusBar1_PanelClick(ByVal Panel As MSComCtlLib.Panel)
frmMonth.Show
End Sub
