Attribute VB_Name = "modGeneral"
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long


Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4


Public strBGFileName As String
Public strBGFileTitle As String
Public strBGMFileName As String
Public strBGMFileTitle As String
Public strBFileName As String
Public strBFileTitle As String
Public strFileName As String
Public strFileTitle As String
Public iPlaceEnemy As Integer

Public Type Level
  BossXL1 As Long
  BossXL2 As Long
  BossXM1 As Long
  BossXM2 As Long
  MaxRows As Long
  lPos() As Enemy
  BossShield As Long
  BossHull As Long
  BossLaserDamage As Long
End Type


Public Result As String
Public Sub FlatBorder(ByVal hwnd As Long)
  Dim TFlat As Long
  TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
  TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
  SetWindowLong hwnd, GWL_EXSTYLE, TFlat
  SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub

Public Sub Flatten(ByVal frm As Form)
  Dim CTL As Control
  For Each CTL In frm.Controls
    Select Case TypeName(CTL)
      Case "CommandButton", "TextBox", "ListBox", "FileTree", "TreeView", "ProgressBar"
        FlatBorder CTL.hwnd
    End Select
  Next
End Sub


Public Sub Flatten2(ByVal frm As Form)
  Dim CTL As Control
  For Each CTL In frm.Controls
    Select Case TypeName(CTL)
      Case "CommandButton", "TextBox", "ListBox", "FileTree", "TreeView", "ProgressBar", "PictureBox", "HScrollBar"
        FlatBorder CTL.hwnd
    End Select
  Next
End Sub


