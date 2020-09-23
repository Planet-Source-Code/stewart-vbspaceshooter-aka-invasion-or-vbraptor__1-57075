Attribute VB_Name = "modArchive"
Option Explicit



Public Sub MakeArchive(Filename As String, file1 As String, file2 As String, file3 As String, file4 As String)
  Dim Files(4) As String
  Dim k As Integer
  Dim iFile As Integer
  Dim Dat As String, binary As String
  Files(0) = file1
  Files(1) = file2
  Files(2) = file3
  Files(3) = file4
  For k = 0 To 3
    iFile = FreeFile
    Open Files(k) For Binary As #iFile
      Dat = String(LOF(iFile), Chr(0))
      Get iFile, , Dat
    Close #iFile
    If k > 0 Then
      binary = binary & "<<<|||>>>"
    End If
    binary = binary & Dat
  Next k
  iFile = FreeFile
  Open Filename For Binary Access Write As #iFile
    Put #iFile, , binary
  Close #iFile
'  DeleteFile App.Path & "\temp\tmpBack.bmp"
'  DeleteFile App.Path & "\temp\tmpBoss.bmp"
'  DeleteFile App.Path & "\temp\tmpBGMusic.mid"
'  DeleteFile App.Path & "\temp\tmpEnemy.enm"
End Sub

Public Sub DeArchive(Filename As String)
  Dim iFile As Integer
  Dim Dat As String
  Dim tmpStr As String
  Dim x As Long, Y As Long, z As Long
  iFile = FreeFile
  Open Filename For Binary As #iFile
    Dat = String(LOF(iFile), Chr(0))
    Get iFile, , Dat
  Close #iFile

  x = InStr(1, Dat, "<<<|||>>>")
  tmpStr = Mid(Dat, 1, x)
    Open App.Path & "\temp\tmpBack.bmp" For Binary Access Write As #1
      Put #1, , tmpStr
    Close #1
    
  
  
  Y = InStr(x + 9, Dat, "<<<|||>>>")
  tmpStr = Mid(Dat, x + 9, Y)
    Open App.Path & "\temp\tmpBGMusic.mid" For Binary Access Write As #1
      Put #1, , tmpStr
    Close #1
  
  
  x = InStr(Y + 9, Dat, "<<<|||>>>")
  tmpStr = Mid(Dat, Y + 9, x)
    Open App.Path & "\temp\tmpBoss.bmp" For Binary Access Write As #1
      Put #1, , tmpStr
    Close #1
  
  
  tmpStr = Mid(Dat, x + 9)
    Open App.Path & "\temp\tmpEnemy.enm" For Binary Access Write As #1
      Put #1, , tmpStr
    Close #1
  
  
End Sub
