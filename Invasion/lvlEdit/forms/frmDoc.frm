VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "HellFire Level Editor"
   ClientHeight    =   9780
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10515
   Icon            =   "frmDoc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9780
   ScaleWidth      =   10515
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBHidden 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3375
      Left            =   8040
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   54
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picE 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   5
      Left            =   6240
      Picture         =   "frmDoc.frx":0CCA
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   7
      Top             =   8040
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox picE 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   4
      Left            =   5280
      Picture         =   "frmDoc.frx":1FCC
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   6
      Top             =   8040
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox picE 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   3
      Left            =   4080
      Picture         =   "frmDoc.frx":32CE
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   5
      Top             =   8040
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox picE 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   2
      Left            =   3120
      Picture         =   "frmDoc.frx":45D0
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   4
      Top             =   8040
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox picE 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   1
      Left            =   1800
      Picture         =   "frmDoc.frx":5052
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   2
      Top             =   8040
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox picE 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   0
      Left            =   840
      Picture         =   "frmDoc.frx":6354
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   1
      Top             =   8040
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox picHidden 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3375
      Left            =   6480
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   9360
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabLayout 
      Height          =   8175
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   14420
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Enemy Positioning"
      TabPicture(0)   =   "frmDoc.frx":7656
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Scrolling Background"
      TabPicture(1)   =   "frmDoc.frx":7672
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Main Boss"
      TabPicture(2)   =   "frmDoc.frx":768E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Music"
      TabPicture(3)   =   "frmDoc.frx":76AA
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Enemy Positioning"
         Height          =   7575
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   9135
         Begin VB.PictureBox picBadGuys 
            Height          =   7215
            Left            =   120
            ScaleHeight     =   7155
            ScaleWidth      =   2715
            TabIndex        =   31
            Top             =   240
            Width           =   2775
            Begin VB.ComboBox cmbAI 
               Height          =   315
               ItemData        =   "frmDoc.frx":76C6
               Left            =   120
               List            =   "frmDoc.frx":76D3
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Top             =   720
               Width           =   2535
            End
            Begin VB.PictureBox picTitle 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'None
               Height          =   375
               Left            =   0
               ScaleHeight     =   375
               ScaleWidth      =   2775
               TabIndex        =   41
               Top             =   0
               Width           =   2775
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Bad Guys"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   240
                  TabIndex        =   42
                  Top             =   75
                  Width           =   2175
               End
            End
            Begin VB.VScrollBar vEScroll 
               Height          =   5775
               LargeChange     =   300
               Left            =   2400
               Max             =   7200
               SmallChange     =   100
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   1080
               Width           =   255
            End
            Begin VB.PictureBox picEStore 
               Height          =   5775
               Left            =   120
               ScaleHeight     =   5715
               ScaleWidth      =   2115
               TabIndex        =   32
               Top             =   1080
               Width           =   2175
               Begin VB.PictureBox picScroll 
                  BorderStyle     =   0  'None
                  Height          =   7200
                  Left            =   0
                  ScaleHeight     =   7200
                  ScaleWidth      =   2175
                  TabIndex        =   33
                  Top             =   0
                  Width           =   2175
                  Begin VB.OptionButton OptE1 
                     Caption         =   "Enemy 6"
                     Height          =   1170
                     Index           =   5
                     Left            =   0
                     Picture         =   "frmDoc.frx":76FF
                     Style           =   1  'Graphical
                     TabIndex        =   39
                     Top             =   6030
                     Width           =   2100
                  End
                  Begin VB.OptionButton OptE1 
                     Caption         =   "Enemy 5"
                     Height          =   1170
                     Index           =   4
                     Left            =   0
                     Picture         =   "frmDoc.frx":AF81
                     Style           =   1  'Graphical
                     TabIndex        =   38
                     Top             =   4824
                     Width           =   2100
                  End
                  Begin VB.OptionButton OptE1 
                     Caption         =   "Enemy 4"
                     Height          =   1170
                     Index           =   3
                     Left            =   0
                     Picture         =   "frmDoc.frx":F61F
                     Style           =   1  'Graphical
                     TabIndex        =   37
                     Top             =   3618
                     Width           =   2100
                  End
                  Begin VB.OptionButton OptE1 
                     Caption         =   "Enemy 3"
                     Height          =   1170
                     Index           =   2
                     Left            =   0
                     Picture         =   "frmDoc.frx":10EC1
                     Style           =   1  'Graphical
                     TabIndex        =   36
                     Top             =   2412
                     Width           =   2100
                  End
                  Begin VB.OptionButton OptE1 
                     Caption         =   "Enemy 2"
                     Height          =   1170
                     Index           =   1
                     Left            =   0
                     Picture         =   "frmDoc.frx":117C3
                     Style           =   1  'Graphical
                     TabIndex        =   35
                     Top             =   1206
                     Width           =   2100
                  End
                  Begin VB.OptionButton OptE1 
                     Caption         =   "Enemy 1"
                     Height          =   1170
                     Index           =   0
                     Left            =   0
                     Picture         =   "frmDoc.frx":1364D
                     Style           =   1  'Graphical
                     TabIndex        =   34
                     Top             =   0
                     Value           =   -1  'True
                     Width           =   2100
                  End
               End
            End
            Begin VB.Label Label2 
               Caption         =   "AI To Use"
               Height          =   375
               Left            =   120
               TabIndex        =   44
               Top             =   480
               Width           =   1575
            End
         End
         Begin VB.PictureBox picField 
            Height          =   7215
            Left            =   3000
            ScaleHeight     =   7155
            ScaleWidth      =   5850
            TabIndex        =   28
            Top             =   240
            Width           =   5910
            Begin VB.VScrollBar vFScroll 
               Height          =   7095
               LargeChange     =   10
               Left            =   5640
               Max             =   100
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   0
               Width           =   255
            End
            Begin VB.PictureBox picMap 
               AutoRedraw      =   -1  'True
               BorderStyle     =   0  'None
               ForeColor       =   &H000000C0&
               Height          =   6495
               Left            =   0
               ScaleHeight     =   433
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   377
               TabIndex        =   29
               Top             =   0
               Width           =   5655
               Begin VB.Shape shpPos 
                  BorderColor     =   &H000000C0&
                  BorderWidth     =   2
                  Height          =   600
                  Left            =   1800
                  Top             =   1080
                  Width           =   600
               End
            End
            Begin VB.Timer tmrField 
               Interval        =   1
               Left            =   5280
               Top             =   2400
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Music"
         Height          =   7575
         Left            =   -74880
         TabIndex        =   23
         Top             =   360
         Width           =   9375
         Begin VB.ComboBox cmbMFilename 
            Height          =   315
            Left            =   120
            TabIndex        =   25
            Top             =   480
            Width           =   6015
         End
         Begin VB.CommandButton cmdMFilename 
            Caption         =   ".."
            Height          =   315
            Left            =   6240
            TabIndex        =   24
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblFilename 
            Caption         =   "Filename:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Scrolling Background"
         Height          =   7575
         Left            =   -74880
         TabIndex        =   19
         Top             =   480
         Width           =   9375
         Begin VB.PictureBox picBackgroundS 
            Height          =   6615
            Left            =   120
            ScaleHeight     =   6555
            ScaleWidth      =   9075
            TabIndex        =   59
            Top             =   840
            Width           =   9135
            Begin VB.HScrollBar hBGScroll 
               Height          =   255
               LargeChange     =   100
               Left            =   0
               SmallChange     =   50
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   6000
               Width           =   2535
            End
            Begin VB.VScrollBar vBGScroll 
               Height          =   6255
               LargeChange     =   100
               Left            =   8400
               SmallChange     =   50
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   0
               Width           =   255
            End
            Begin VB.PictureBox picBackground 
               AutoRedraw      =   -1  'True
               BorderStyle     =   0  'None
               Height          =   6615
               Left            =   0
               ScaleHeight     =   441
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   585
               TabIndex        =   60
               Top             =   0
               Width           =   8775
            End
         End
         Begin VB.ComboBox cmbFilename 
            Height          =   315
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   8175
         End
         Begin VB.CommandButton cmdFilename 
            Caption         =   ".."
            Height          =   315
            Left            =   8280
            TabIndex        =   20
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblFilename 
            Caption         =   "Filename:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Main Boss"
         Height          =   7575
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   9375
         Begin VB.PictureBox picBGS 
            AutoRedraw      =   -1  'True
            Height          =   2895
            Left            =   120
            ScaleHeight     =   2835
            ScaleWidth      =   9075
            TabIndex        =   55
            Top             =   4560
            Width           =   9135
            Begin VB.HScrollBar hMBGScroll 
               Height          =   255
               LargeChange     =   100
               Left            =   2280
               SmallChange     =   50
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   2520
               Width           =   2535
            End
            Begin VB.VScrollBar vMBGScroll 
               Height          =   6255
               LargeChange     =   100
               Left            =   0
               SmallChange     =   50
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   0
               Width           =   255
            End
            Begin VB.PictureBox picBG 
               AutoRedraw      =   -1  'True
               BorderStyle     =   0  'None
               Height          =   2775
               Left            =   0
               ScaleHeight     =   185
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   585
               TabIndex        =   56
               Top             =   0
               Width           =   8775
            End
         End
         Begin VB.ComboBox missile1X 
            Height          =   315
            Left            =   4920
            TabIndex        =   51
            Text            =   "0"
            Top             =   3480
            Width           =   4695
         End
         Begin VB.ComboBox missile2X 
            Height          =   315
            Left            =   4920
            TabIndex        =   50
            Text            =   "0"
            Top             =   4080
            Width           =   4695
         End
         Begin VB.ComboBox laser1X 
            Height          =   315
            Left            =   120
            TabIndex        =   47
            Text            =   "0"
            Top             =   3480
            Width           =   4695
         End
         Begin VB.ComboBox laser2X 
            Height          =   315
            Left            =   120
            TabIndex        =   46
            Text            =   "0"
            Top             =   4080
            Width           =   4695
         End
         Begin VB.CommandButton cmdBFilename 
            Caption         =   ".."
            Height          =   315
            Left            =   4920
            TabIndex        =   45
            Top             =   480
            Width           =   375
         End
         Begin VB.ComboBox cmbBFilename 
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   4695
         End
         Begin VB.ComboBox cmbMissile 
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Text            =   "40"
            Top             =   2880
            Width           =   4695
         End
         Begin VB.ComboBox cmbLaser 
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Text            =   "20"
            Top             =   2280
            Width           =   4695
         End
         Begin VB.ComboBox cmbHull 
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Text            =   "200"
            Top             =   1680
            Width           =   4695
         End
         Begin VB.ComboBox cmbShield 
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Text            =   "100"
            Top             =   1080
            Width           =   4695
         End
         Begin VB.Label lblTitle 
            Caption         =   "Missile1 X:"
            Height          =   615
            Index           =   8
            Left            =   4920
            TabIndex        =   53
            Top             =   3240
            Width           =   2655
         End
         Begin VB.Label lblTitle 
            Caption         =   "Missile2 X:"
            Height          =   735
            Index           =   7
            Left            =   4920
            TabIndex        =   52
            Top             =   3840
            Width           =   2655
         End
         Begin VB.Label lblTitle 
            Caption         =   "Laser1 X:"
            Height          =   615
            Index           =   6
            Left            =   120
            TabIndex        =   49
            Top             =   3240
            Width           =   2655
         End
         Begin VB.Label lblTitle 
            Caption         =   "Laser2 X:"
            Height          =   735
            Index           =   5
            Left            =   120
            TabIndex        =   48
            Top             =   3840
            Width           =   2655
         End
         Begin VB.Label lblTitle 
            Caption         =   "Filename:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label lblTitle 
            Caption         =   "Missile Damage:"
            Height          =   735
            Index           =   3
            Left            =   120
            TabIndex        =   17
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Label lblTitle 
            Caption         =   "Laser Damage:"
            Height          =   735
            Index           =   2
            Left            =   120
            TabIndex        =   16
            Top             =   2040
            Width           =   2655
         End
         Begin VB.Label lblTitle 
            Caption         =   "Hull:"
            Height          =   735
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   1440
            Width           =   2655
         End
         Begin VB.Label lblTitle 
            Caption         =   "Shield:"
            Height          =   735
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   2655
         End
      End
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuRandom 
         Caption         =   "Generate Random Enemy Layout"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuInstructions 
         Caption         =   "&Instructions"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdBFilename_Click()
  On Error GoTo exitError
  With cd
    .Filter = "Bitmap Files|*.bmp|All Files|*.*"
    .CancelError = True
    .ShowOpen
    strBFileName = .Filename
    strBFileTitle = .FileTitle
    cmbBFilename.Text = .FileTitle
    picBHidden.Picture = LoadPicture(.Filename)
  End With
  hMBGScroll.Max = picBHidden.ScaleWidth
  vMBGScroll.Max = picBHidden.ScaleHeight
  picBG_Paint
exitError:
  'Just exit. don't need to really do anything
End Sub

Private Sub cmdFilename_Click()
  On Error GoTo exitError
  With cd
    .Filter = "Bitmap Files|*.bmp|All Files|*.*"
    .CancelError = True
    .ShowOpen
    strBGFileName = .Filename
    strBGFileTitle = .FileTitle
    cmbFilename.Text = .FileTitle
    picHidden.Picture = LoadPicture(.Filename)
  End With
  hBGScroll.Max = picHidden.ScaleWidth
  vBGScroll.Max = picHidden.ScaleHeight
  picBackground_Paint
exitError:
  'Just exit. don't need to really do anything
End Sub

Private Sub cmdMFilename_Click()
  On Error GoTo exitError
  With cd
    .Filter = "MIDI Files|*.mid|All Files|*.*"
    .CancelError = True
    .ShowOpen
    strBGMFileName = .Filename
    strBGMFileTitle = .FileTitle
    cmbMFilename.Text = .FileTitle
    mmTest.Open .Filename
  End With
exitError:
End Sub

Private Sub Form_Load()
  MapRows = 100
  Flatten Me
  iPlaceEnemy = 0
  cmbAI.ListIndex = 0
  ReDim EPos(5, 100)
  PaintField
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  tabLayout.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
  Frame1.Left = 120
  Frame1.Top = Frame1.Top
  Frame1.Move 120, Frame1.Top, tabLayout.Width - Frame1.Left - 120, tabLayout.Height - Frame1.Top - 120
  Frame2.Move 120, Frame1.Top, tabLayout.Width - Frame1.Left - 120, tabLayout.Height - Frame1.Top - 120
  Frame3.Move 120, Frame1.Top, tabLayout.Width - Frame1.Left - 120, tabLayout.Height - Frame1.Top - 120
  Frame4.Move 120, Frame1.Top, tabLayout.Width - Frame1.Left - 120, tabLayout.Height - Frame1.Top - 120
  picBackgroundS.Width = tabLayout.Width - picBackgroundS.Left - 360
  picBackgroundS.Height = tabLayout.Height - picBackgroundS.Top - 560
  cmdFilename.Left = (picBackgroundS.Width + picBackgroundS.Left) - cmdFilename.Width
  cmbFilename.Width = tabLayout.Width - cmdFilename.Width - 580
  cmdFilename.Top = cmbFilename.Top
  cmdMFilename.Left = (picBackgroundS.Width + picBackgroundS.Left) - cmdFilename.Width
  cmbMFilename.Width = tabLayout.Width - cmdFilename.Width - 580
  cmdMFilename.Top = cmbFilename.Top
  cmdBFilename.Left = (picBackgroundS.Width + picBackgroundS.Left) - cmdFilename.Width
  cmbBFilename.Width = tabLayout.Width - cmdFilename.Width - 580
  cmdBFilename.Top = cmbFilename.Top
  cmbShield.Width = picBackgroundS.Width
  cmbHull.Width = picBackgroundS.Width
  cmbLaser.Width = picBackgroundS.Width
  cmbMissile.Width = picBackgroundS.Width
  picBGS.Move picBackgroundS.Left, picBGS.Top, picBackgroundS.Width, tabLayout.Height - picBGS.Top - 560
  hBGScroll.Move 0, picBackgroundS.ScaleHeight - 255, picBackgroundS.ScaleWidth - 255, 255
  vBGScroll.Move picBackgroundS.ScaleWidth - 255, 0, 255, picBackgroundS.ScaleHeight - 255
  hMBGScroll.Move 0, picBGS.ScaleHeight - 255, picBGS.ScaleWidth - 255, 255
  vMBGScroll.Move picBGS.ScaleWidth - 255, 0, 255, picBGS.ScaleHeight - 255
  
  picBadGuys.Move 120, picBadGuys.Top, picBadGuys.Width, tabLayout.Height - picBadGuys.Top - 560
  picField.Move picBadGuys.Left + picBadGuys.Width + 120, picBadGuys.Top, tabLayout.Width - 360 - picField.Left, tabLayout.Height - picField.Top - 560
  picEStore.Move 120, picEStore.Top, picEStore.Width, picBadGuys.ScaleHeight - picEStore.Top - 360
  vEScroll.Move vEScroll.Left, picEStore.Top, 255, picEStore.Height
  vFScroll.Max = (MapRows - (picMap.ScaleHeight \ 39))
  laser1X.Width = (Frame3.Width - 360) \ 2
  laser2X.Width = (Frame3.Width - 360) \ 2
  missile1X.Width = (Frame3.Width - 360) \ 2
  missile2X.Width = (Frame3.Width - 360) \ 2
  missile1X.Left = laser1X.Left + laser1X.Width + 120
  missile2X.Left = laser1X.Left + laser1X.Width + 120
  lblTitle(7).Left = missile1X.Left
  lblTitle(8).Left = missile1X.Left
  picBG.Move 0, 0, picBGS.ScaleWidth - 255, picBGS.ScaleHeight - 255
  picBackground.Move 0, 0, picBackgroundS.ScaleWidth - 255, picBackgroundS.ScaleHeight - 255
  PaintField
End Sub


Private Sub hBGScroll_Change()
  picBackground_Paint
End Sub

Private Sub hBGScroll_Scroll()
  picBackground_Paint
End Sub

Private Sub hMBGScroll_Change()
  picBG_Paint
End Sub

Private Sub hMBGScroll_Scroll()
  picBG_Paint
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show vbModal, Me
End Sub

Private Sub mnuExit_Click()
  Dim frm As Form
  For Each frm In Forms
    Unload frm
  Next
  End
End Sub

Private Sub mnuInstructions_Click()
  frmInstructions.Show vbModal, Me
End Sub

Private Sub mnuNew_Click()
  Dim iRows As Long, i As Integer, x As Integer, strTmp As String
rowMistake:
  strFileName = ""
  strTmp = InputBox("Enter the number of rows for this map.", "Enter Rows", 100)
  If Len(strTmp) = 0 Then Exit Sub
  iRows = strTmp
  If Not IsNumeric(iRows) Then
    MsgBox "You must enter a number.", vbCritical
    GoTo rowMistake
  End If
  If iRows < 50 Or iRows > 30000 Then
    MsgBox "You must enter a number between 50 and 30000", vbCritical
    GoTo rowMistake
  End If
  MapRows = iRows
  iPlaceEnemy = 0
  cmbAI.ListIndex = 0
  ReDim EPos(5, iRows)
  For x = 0 To 5
    For i = 0 To iRows
      EPos(x, i).AI = 0
      EPos(x, i).EnemyInt = 0
    Next i
  Next x
  PaintField
  Form_Resize
End Sub

Private Sub mnuOpen_Click()
 'On Error GoTo errExit
  Dim x As Integer, Y As Integer, f As Integer
  Dim strTemp As String, strStore As String
  Dim z As Long
  Dim fFile As Integer
  
  With cd
    .Filter = "Level Files (*.hla)|*.hla|All Files|*.*"
    .CancelError = True
    .ShowOpen
  End With
  fFile = FreeFile()
  strFileName = cd.Filename
  strFileTitle = cd.FileTitle
  DeArchive cd.Filename
  Open App.Path & "\temp\tmpEnemy.enm" For Input As #fFile
    Input #fFile, MapRows
    Input #fFile, strBGFileTitle
    Input #fFile, strBGMFileTitle
    Input #fFile, strBFileTitle
    Input #fFile, strBGFileName
    Input #fFile, strBGMFileName
    Input #fFile, strBFileName
    
    strBGFileName = Left(App.Path, Len(App.Path) - 7) & strBGFileName
    strBGMFileName = Left(App.Path, Len(App.Path) - 7) & strBGMFileName
    strBFileName = Left(App.Path, Len(App.Path) - 7) & strBFileName
    Input #fFile, strTemp
    cmbShield.Text = strTemp
    Input #fFile, strTemp
    cmbHull.Text = strTemp
    Input #fFile, strTemp
    cmbLaser.Text = strTemp
    Input #fFile, strTemp
    cmbMissile.Text = strTemp
    Input #fFile, strTemp
    laser1X.Text = strTemp
    Input #fFile, strTemp
    laser2X.Text = strTemp
    Input #fFile, strTemp
    missile1X.Text = strTemp
    Input #fFile, strTemp
    missile2X.Text = strTemp
    cmbFilename.Text = strBGFileTitle
    cmbMFilename.Text = strBGMFileTitle
    cmbBFilename.Text = strBFileTitle
    
    'mmTest.DeviceType = "Sequencer"
    'mmTest.FileName = strBGMFileName
    'mmTest.Command = "Open"
    strBGFileName = App.Path & "\temp\tmpBack.bmp"
    strBFileName = App.Path & "\temp\tmpBoss.bmp"
    strBGMFileName = App.Path & "\temp\tmpBGMusic.mid"
    picHidden.Picture = LoadPicture(App.Path & "\temp\tmpBack.bmp")
    picBHidden.Picture = LoadPicture(App.Path & "\temp\tmpBoss.bmp")
    'mmTest.Pause
    ReDim EPos(5, MapRows)
    Y = 0
    For z = 1 To MapRows
      Input #fFile, strTemp
      For x = 1 To Len(strTemp) Step 2
        strStore = Mid$(strTemp, x, 2)
        If Asc(Left(strStore, 1)) = 32 Then
          EPos((x - 1) \ 2, Y).EnemyInt = 0
        Else
          
          EPos((x - 1) \ 2, Y).EnemyInt = (Left(strStore, 1))
        End If
        If Asc(Right(strStore, 1)) = 32 Then
          EPos((x - 1) \ 2, Y).AI = 0
        Else
          EPos((x - 1) \ 2, Y).AI = (Right(strStore, 1))
        End If
        f = f + 1
      Next x
      
      Y = Y + 1
    Next z
  Close #fFile
  PaintField
  ResizeIt
errExit:
  'Do nothing since most likely just means cancel was pressed
End Sub

Private Sub mnuRandom_Click()
  frmRandom.Show vbModal, Me
End Sub

Private Sub mnuSave_Click()
  On Error GoTo errExit
  Dim strTemp As String
  Dim p As Level
  Dim x As Integer, Y As Integer
  If cmbFilename.Text = "" Then
    MsgBox "You must first select a background to be scrolled.", vbCritical
    Exit Sub
  End If
  If cmbMFilename.Text = "" Then
    MsgBox "You must select a background music file.", vbCritical
    Exit Sub
  End If
  If cmbBFilename.Text = "" Then
    MsgBox "You must select a boss filename.", vbCritical
    Exit Sub
  End If
  With cd
    .Filter = "Level Files (*.hla)|*.hla|All Files|*.*"
    .CancelError = True
    .ShowSave
  End With
  strFileName = cd.Filename
  strFileTitle = cd.FileTitle
  Dim fFile As Integer
  Dim strWrite As String
  fFile = FreeFile()
  Open App.Path & "\temp\tmpEnemy.enm" For Output As #fFile
  strTemp = MapRows
  Print #fFile, strTemp
  Print #fFile, strBGFileTitle
  Print #fFile, strBGMFileTitle
  Print #fFile, strBFileTitle
  'These are only used by the level editor
  'to load in the appropriate settings when
  'loading a file
  Print #fFile, strBGFileName
  Print #fFile, strBGMFileName
  Print #fFile, strBFileName
  Print #fFile, cmbShield.Text
  Print #fFile, cmbHull.Text
  Print #fFile, cmbLaser.Text
  Print #fFile, cmbMissile.Text
  Print #fFile, laser1X.Text
  Print #fFile, laser2X.Text
  Print #fFile, missile1X.Text
  Print #fFile, missile2X.Text
  For Y = 0 To MapRows
    strWrite = ""
    For x = 0 To 5
     'MsgBox EPos(x, y).EnemyInt & " | " & EPos(x, y).AI
      strWrite = strWrite & EPos(x, Y).EnemyInt & EPos(x, Y).AI
     
    Next x
    Print #fFile, strWrite
  Next Y
  Close #fFile
  MakeArchive cd.Filename, strBGFileName, strBGMFileName, strBFileName, App.Path & "\temp\tmpEnemy.enm"
errExit:
  'Do nothing since it was likely caused
  'by a cancel error
End Sub

Private Sub OptE1_Click(Index As Integer)
  iPlaceEnemy = Index
End Sub

Private Sub picBackground_Paint()
  picBackground.Cls
  BitBlt picBackground.hdc, -hBGScroll.Value, -vBGScroll.Value, picHidden.ScaleWidth, picHidden.ScaleHeight, picHidden.hdc, 0, 0, vbSrcCopy
End Sub

Private Sub picBG_Paint()
  picBG.Cls
  BitBlt picBG.hdc, -hMBGScroll.Value, -vMBGScroll.Value, picBHidden.ScaleWidth, picBHidden.ScaleHeight, picBHidden.hdc, 0, 0, vbSrcCopy

End Sub

Private Sub picField_Resize()
  vFScroll.Move picField.ScaleWidth - 255, 0, 255, picField.ScaleHeight
  'hFScroll.Move 0, picField.ScaleHeight - 255, picField.ScaleWidth - 255, 255
  picMap.Move 0, 0, picField.ScaleWidth - 255, picField.ScaleHeight
End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dim l As Long, p As Long
  l = (x \ 41)
  If l > 5 Then Exit Sub
  p = (Y \ 41)
  p = p + vFScroll.Value
  If p > MapRows Then Exit Sub
  
  EPos(l, p).AI = cmbAI.ListIndex + 1
  EPos(l, p).EnemyInt = iPlaceEnemy + 1
  PaintField
End Sub

Public Sub ResizeIt()
  vFScroll.Max = (MapRows - (picMap.ScaleHeight \ 39))
  PaintField
End Sub

Public Sub PaintField()
  Dim x As Integer, Y As Integer, v As Integer, h As Integer, g As Integer
  v = vFScroll.Value
  
  picMap.Cls
  h = (picMap.ScaleHeight \ 40)
  For x = 0 To 6
    For Y = 0 To picMap.ScaleHeight Step 41
      picMap.Line (x * 40, 0)-(x * 40, picMap.ScaleHeight)
      picMap.Line (0, Y)-(40 * 6, Y)
    Next Y
  Next x
  For x = 0 To 5
    For Y = v To h + v
      g = Y + vFScroll.Value
      
      If EPos(x, Y).EnemyInt > 0 Then
        BitBlt picMap.hdc, x * 40 + 1, Y * (40 + 1) - (v * (40 + 1)), 40, 40, picE(EPos(x, Y).EnemyInt - 1).hdc, 0, 0, vbSrcCopy
      End If
    Next Y
  Next x
End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dim l As Integer, p As Integer
  l = (x \ 41)
  If l > 6 Then Exit Sub
  p = (Y \ 41)
  p = p + vFScroll.Value
  If p > MapRows Then Exit Sub
  shpPos.Left = l * 40
  shpPos.Top = (p - vFScroll.Value) * (41)
End Sub

Private Sub tabLayout_Click(PreviousTab As Integer)
  Select Case tabLayout.Tab
    Case 0
      Frame1.Visible = True
      Frame2.Visible = False
      Frame3.Visible = False
      Frame4.Visible = False
    Case 1
      Frame2.Visible = True
      Frame1.Visible = False
      Frame3.Visible = False
      Frame4.Visible = False
    Case 2
      Frame3.Visible = True
      Frame1.Visible = False
      Frame2.Visible = False
      Frame4.Visible = False
    Case 3
      Frame4.Visible = True
      Frame1.Visible = False
      Frame3.Visible = False
      Frame2.Visible = False
      
End Select
End Sub

Private Sub vBGScroll_Change()
  picBackground_Paint
End Sub

Private Sub vBGScroll_Scroll()
  picBackground_Paint
End Sub

Private Sub VScroll1_Change()
  
End Sub

Private Sub vEScroll_Change()
  picScroll.Top = -vEScroll.Value
End Sub

Private Sub vEScroll_Scroll()
  picScroll.Top = -vEScroll.Value
End Sub

Private Sub vFScroll_Change()
  PaintField
End Sub

Private Sub vFScroll_Scroll()
  PaintField
End Sub

Private Sub vMBGScroll_Change()
  picBG_Paint
End Sub

Private Sub vMBGScroll_Scroll()
  picBG_Paint
End Sub
