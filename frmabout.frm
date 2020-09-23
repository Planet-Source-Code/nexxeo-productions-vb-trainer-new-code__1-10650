VERSION 5.00
Begin VB.Form frmabout 
   BackColor       =   &H80000007&
   Caption         =   "About this trainer"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmabout.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmabout.frx":11CA
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label ok 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   1920
      TabIndex        =   1
      Top             =   4320
      Width           =   360
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ok_Click()
Unload frmabout
End Sub
