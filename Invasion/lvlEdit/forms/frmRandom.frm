VERSION 5.00
Begin VB.Form frmRandom 
   Caption         =   "Generate Random Enemy Layout"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   6120
      TabIndex        =   11
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   3360
      Width           =   1455
   End
   Begin VB.PictureBox picSplit 
      Height          =   55
      Left            =   120
      ScaleHeight     =   0
      ScaleWidth      =   7395
      TabIndex        =   9
      Top             =   3120
      Width           =   7455
   End
   Begin VB.Frame fmDifficulty 
      Caption         =   "Difficulty"
      Height          =   1335
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   3375
      Begin VB.HScrollBar scrlDifficulty 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   960
         Value           =   25
         Width           =   3135
      End
      Begin VB.Label lblValue 
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame fmInclude 
      Caption         =   "Include Enemies"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CheckBox chkE 
         Caption         =   "Enemy Six"
         Height          =   1095
         Index           =   5
         Left            =   2640
         Picture         =   "frmRandom.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox chkE 
         Caption         =   "Enemy Five"
         Height          =   1095
         Index           =   4
         Left            =   1440
         Picture         =   "frmRandom.frx":1302
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox chkE 
         Caption         =   "Enemy Four"
         Height          =   1095
         Index           =   3
         Left            =   240
         Picture         =   "frmRandom.frx":2604
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkE 
         Caption         =   "Enemy Three"
         Height          =   1095
         Index           =   2
         Left            =   2640
         Picture         =   "frmRandom.frx":3906
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkE 
         Caption         =   "Enemy Two"
         Height          =   1095
         Index           =   1
         Left            =   1440
         Picture         =   "frmRandom.frx":4388
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkE 
         Caption         =   "Enemy One"
         Height          =   1095
         Index           =   0
         Left            =   240
         Picture         =   "frmRandom.frx":568A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmRandom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub picE_Click(Index As Integer)
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdGenerate_Click()
  Dim i As Integer, x As Integer
  Dim lDif As Long, iEnemy As Integer, iAI As Integer
  
  Me.Hide
  For i = 0 To MapRows - 1
    For x = 0 To 5
      Randomize Timer
      lDif = Int(Rnd * 2000)
      If lDif < (scrlDifficulty.Value * 15) Then
        Randomize Time
        iEnemy = Int(Rnd * 6)
        Do Until chkE(iEnemy).Value = 1
          iEnemy = Int(Rnd * 6)
        Loop
        EPos(x, i).EnemyInt = iEnemy + 1
        iAI = Int(Rnd * 3) + 1
        EPos(x, i).AI = iAI
      End If
    Next x
  Next i
End Sub

Private Sub Form_Load()
  Flatten2 Me
End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub scrlDifficulty_Change()
  lblValue.Caption = scrlDifficulty.Value
End Sub

Private Sub scrlDifficulty_Scroll()
  lblValue.Caption = scrlDifficulty.Value
End Sub
