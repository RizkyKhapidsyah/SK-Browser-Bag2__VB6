VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMonth 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calender"
   ClientHeight    =   2400
   ClientLeft      =   2865
   ClientTop       =   5880
   ClientWidth     =   2730
   Icon            =   "frmMonth.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   2730
   Begin MSComCtl2.MonthView mnvCal 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   16711935
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
      MonthBackColor  =   65535
      StartOfWeek     =   24510465
      TitleBackColor  =   16711935
      TitleForeColor  =   65535
      TrailingForeColor=   49152
      CurrentDate     =   37071
   End
End
Attribute VB_Name = "frmMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
