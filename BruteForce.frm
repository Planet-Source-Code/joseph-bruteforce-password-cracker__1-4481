VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Joseph's Brute force Password Cracker"
   ClientHeight    =   3780
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmWait 
      Caption         =   "Wait Mode"
      Height          =   495
      Left            =   1320
      TabIndex        =   16
      Top             =   3240
      Width           =   4455
      Begin VB.OptionButton optWaitOff 
         Caption         =   "Off State"
         Height          =   195
         Left            =   2400
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optWaitOn 
         Caption         =   "On State"
         Height          =   195
         Left            =   480
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Timer Tick 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   1200
   End
   Begin MSComDlg.CommonDialog CDial 
      Left            =   1200
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Exit BF"
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   975
   End
   Begin VB.Frame frmOpt 
      Caption         =   "Cracking Method"
      Height          =   1455
      Left            =   1320
      TabIndex        =   6
      Top             =   600
      Width           =   4455
      Begin VB.CommandButton cmdLimit 
         Caption         =   "Set Limits"
         Height          =   315
         Left            =   3360
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optTxt 
         Caption         =   "Send whatever is in the Text Box"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   2895
      End
      Begin VB.OptionButton optNum 
         Caption         =   "Sequence of Numbers"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton optFile 
         Caption         =   "Load Passwords from file"
         Height          =   435
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.TextBox txtFinal 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2160
      Width           =   4455
   End
   Begin VB.TextBox txtSetup 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
   Begin VB.TextBox txtDelay 
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Text            =   "10"
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmdBBF 
      Caption         =   "Break Away"
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdBF 
      Caption         =   "Brute Force"
      Default         =   -1  'True
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Do this after entering the Password,"
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Enter the Password using this method."
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Do This to get to the Password Box for the first time."
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Delay"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   2880
      Width           =   615
   End
   Begin VB.Menu menFile 
      Caption         =   "&File"
      Begin VB.Menu miExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menHelp 
      Caption         =   "&Help"
      Begin VB.Menu mikeys 
         Caption         =   "&Keys"
         HelpContextID   =   1
      End
      Begin VB.Menu miAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tickdelay, ticktotal As Variant

Private Sub cmdBBF_Click()
End
End Sub

Private Sub cmdBF_Click()
ticktotal = txtDelay.Text
tickdelay = ticktotal

Tick.Enabled = True
initialisemode = 1


End Sub

Private Sub cmdBrowse_Click()
CDial.ShowOpen
optFile.Caption = "Load Passwords from this file " + CDial.FileName
optFile.Value = True

End Sub

Private Sub cmdLimit_Click()
frmLimit.Show vbModal, Me


End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Activate()
ticktotal = txtDelay.Text
optWaitOff.Value = True
tickdelay = ticktotal
End Sub

Private Sub Form_Click()
If initialisemode = 0 Then End

End Sub

Private Sub Form_Load()
initilisemode = 1
ticktotal = 10
txtDelay.Text = ticktotal

count1 = 0
End Sub

Private Sub miAbout_Click()
frmAbout.Show vbModal, Me


End Sub

Private Sub miExit_Click()
End
End Sub

Private Sub mikeys_Click()
frmHelp.Show vbModal, Me

End Sub

Private Sub Tick_Timer()
tickdelay = tickdelay - 1
txtDelay.Text = tickdelay
txtDelay.Refresh
 
If tickdelay < 1 Then
initialisemode = 0
Tick.Enabled = False
SendKeys frmMain.txtSetup
If optFile.Value = True Then BFFile
If optNum.Value = True Then BFSequence
If optTxt.Value = True Then BFText
End If
End Sub

Private Sub txtDelay_Change()
ticktotal = txtDelay.Text
End Sub
