VERSION 5.00
Begin VB.Form frmInstructions 
   Caption         =   "Instructions"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   Icon            =   "frmInstructions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInstructions 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   600
      Width           =   3975
   End
End
Attribute VB_Name = "frmInstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Dim fFile As Integer
  
  fFile = FreeFile
  Open App.Path & "\instructions.txt" For Input As #fFile
    txtInstructions.Text = Input(LOF(fFile), fFile)
  Close #fFile
End Sub

Private Sub Form_Resize()
  txtInstructions.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
