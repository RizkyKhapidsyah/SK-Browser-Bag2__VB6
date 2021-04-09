VERSION 5.00
Begin VB.Form frmTextEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Macrohard - Sound"
   ClientHeight    =   5955
   ClientLeft      =   2865
   ClientTop       =   1725
   ClientWidth     =   6510
   Icon            =   "TextEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   6510
   Begin VB.TextBox txtLoc 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox txtEditor 
      Height          =   5055
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label lblLoc 
      Caption         =   "Location:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   5520
      Width           =   855
   End
End
Attribute VB_Name = "frmTextEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
Dim strType, strPath As String

strType = txtEditor.Text
strPath = txtLoc.Text

txtEditor.SetFocus

On Error GoTo errorhandler

Open strPath For Output As #1

Print #1, strType

Close #1

txtEditor.Text = ""
txtLoc.Text = ""

'Displays Saves message box
MsgBox "Document Saved"

errorhandler:
'stop errors and shut downing
End Sub

