VERSION 5.00
Begin VB.Form frmAddress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Address"
   ClientHeight    =   420
   ClientLeft      =   510
   ClientTop       =   570
   ClientWidth     =   3495
   Icon            =   "frmAddress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   3495
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim strSearch1 As String

If KeyAscii = 13 Then
strSearch1 = frmAddress.Text1.Text
frmBrowser.wbBrowser.Navigate strSearch1
Unload frmAddress
End If

End Sub
