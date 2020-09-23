VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About HellFire Level Editor"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10035
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   575
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   669
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   615
      Left            =   8205
      TabIndex        =   3
      Top             =   7920
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   13573
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   529
      TabCaption(0)   =   "About"
      TabPicture(0)   =   "frmAbout.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtAbout"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Credits"
      TabPicture(1)   =   "frmAbout.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "p1"
      Tab(1).Control(1)=   "picScroll"
      Tab(1).ControlCount=   2
      Begin VB.PictureBox p1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   7150
         Left            =   -79440
         ScaleHeight     =   477
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   637
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   9550
      End
      Begin VB.TextBox txtAbout 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7150
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "frmAbout.frx":0342
         Top             =   440
         Width           =   9550
      End
      Begin VB.PictureBox picScroll 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   7150
         Left            =   -74880
         Picture         =   "frmAbout.frx":074A
         ScaleHeight     =   473
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   633
         TabIndex        =   2
         Top             =   420
         Width           =   9550
         Begin VB.Timer tmrScroll 
            Interval        =   30
            Left            =   6120
            Top             =   3720
         End
      End
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   7260
      Left            =   -120
      Picture         =   "frmAbout.frx":E178C
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   -1080
      Visible         =   0   'False
      Width           =   9660
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim yPos1 As Long, yPos2 As Long
Private Type Lines
  lineText As String
  LineX As Integer
  LineY As Integer
End Type
Dim iLineCount As Integer
Dim lineTemp() As Lines

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim fFile As Integer, i As Integer
  Flatten Me
  Dim strTemp As String
  yPos1 = 0
  yPos2 = -480
  fFile = FreeFile()
  Open App.Path & "\about.txt" For Input As #fFile
    strTemp = Input(LOF(fFile), #fFile)
    For i = 1 To Len(strTemp)
      If Mid(strTemp, i, 1) = Chr(10) Then iLineCount = iLineCount + 1
    Next
    ReDim lineTemp(iLineCount)
  Close #fFile
  Open App.Path & "\about.txt" For Input As #fFile
    i = 0
    Do While Not EOF(fFile)
      Line Input #fFile, strTemp
      lineTemp(i).lineText = strTemp
      lineTemp(i).LineX = (p1.ScaleWidth - p1.TextWidth(strTemp)) \ 2
      lineTemp(i).LineY = (p1.ScaleHeight + (p1.TextHeight(strTemp) * i) + 4)
      i = i + 1
    Loop
  Close #fFile
End Sub

Private Sub tmrScroll_Timer()
  Dim b As Integer, i As Integer
  yPos1 = yPos1 + 1
  yPos2 = yPos2 + 1
  If yPos1 = p1.ScaleHeight Then
    yPos1 = yPos2 - 480
  End If
  If yPos2 = p1.ScaleHeight Then
    yPos2 = yPos1 - 480
  End If
  'p1.Cls
  BitBlt p1.hdc, 0, yPos1, picBack.ScaleWidth, picBack.ScaleHeight, picBack.hdc, 0, 0, vbSrcCopy
  BitBlt p1.hdc, 0, yPos2, picBack.ScaleWidth, picBack.ScaleHeight, picBack.hdc, 0, 0, vbSrcCopy
  For i = 0 To iLineCount - 1
    lineTemp(i).LineY = lineTemp(i).LineY - 1
    
    'Print out the shadow
    'p1.CurrentX = lineTemp(i).LineX + 2
    'p1.CurrentY = lineTemp(i).LineY + 2
    p1.ForeColor = 0
    TextOut p1.hdc, lineTemp(i).LineX + 2, lineTemp(i).LineY + 2, lineTemp(i).lineText, Len(lineTemp(i).lineText)
    'p1.Print lineTemp(i).lineText
    
    'Print out the yellow text
    'p1.CurrentX = lineTemp(i).LineX
    'p1.CurrentY = lineTemp(i).LineY
    p1.ForeColor = &HFFFF&
    'p1.Print lineTemp(i).lineText
    TextOut p1.hdc, lineTemp(i).LineX, lineTemp(i).LineY, lineTemp(i).lineText, Len(lineTemp(i).lineText)
    If (lineTemp(iLineCount - 1).LineY + p1.TextHeight(lineTemp(iLineCount - 1).lineText)) < 0 Then
      For b = 0 To iLineCount - 1
        lineTemp(b).LineY = (p1.ScaleHeight + (p1.TextHeight(strTemp) * b) + 4)
      Next
    End If
    
  Next
  BitBlt picScroll.hdc, 0, 0, picScroll.ScaleWidth, picScroll.ScaleHeight, p1.hdc, 0, 0, vbSrcCopy
    
End Sub

Sub PrintText(Text As String)
p1.CurrentX = (p1.ScaleWidth / 2) - (p1.TextWidth(Text) / 2)
p1.ForeColor = 0: X = p1.CurrentX: Y = p1.CurrentY
For i = 1 To 3
    p1.Print Text
    X = X + 1: Y = Y + 1: p1.CurrentX = X: p1.CurrentY = Y
Next i
p1.ForeColor = &HFFFF&
p1.Print Text
End Sub
