VERSION 5.00
Begin VB.Form frmLimit 
   Caption         =   "Limit Entry - Brute Force"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4005
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4005
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtUpper 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Text            =   "1000"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtLower 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "1"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "to"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblLimit 
      Caption         =   "Enter the range of numbers through you want me to try for below"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmLimit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
frmLimit.Hide
frmMain.Show

End Sub
