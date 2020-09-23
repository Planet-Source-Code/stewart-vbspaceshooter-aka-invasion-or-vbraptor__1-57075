Attribute VB_Name = "modCollision"
Option Explicit

'+------------------------------------------------------------------+
'| Invasion - modCollision.bas                                      |
'+------------------------------------------------------------------+
'| Design and code by Stewart (sobert81@devedit.com)                |
'+------------------------------------------------------------------+
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Dim bytCenter() As Byte
Dim bytMoving() As Byte

Public Function CheckCollide(DDSurf1 As DirectDrawSurface7, DDSurf2 As DirectDrawSurface7, X1 As Integer, Y1 As Integer, Width1 As Integer, Height1 As Integer, x2 As Integer, y2 As Integer, Width2 As Integer, Height2 As Integer, BlitCOLORKEY As Integer) As Boolean
    On Error Resume Next
    Dim Rect1 As RECT, Rect2 As RECT
    Dim RECTOverlap As RECT 'Used to record the overlap from RECT1 and RECT2
    Dim RECT1Overlap As RECT 'Overlaped portions of RECT1
    Dim RECT2Overlap As RECT 'Overlaped portions of RECT2
    Dim OverlapWidth As Integer 'Determine the width of the overlap
    Dim OverlapHeight As Integer 'Determine the height of the overlap
    Dim ByteObj1() As Byte 'Used to analyse a pixel
    Dim ByteObj2() As Byte 'Used to analyse a pixel
    Dim DDSDBlank As DDSURFACEDESC2 'For use in (DDSurf1.Lock) (DDSurf2.Lock)
    Dim i As Integer, j As Integer 'Just for use in loops
    Dim PPCollision As Boolean 'States whether we have PixelPerfect collision
    CheckCollide = False
    With Rect1
      .Left = X1
      .Top = Y1
      .Right = X1 + Width1
      .Bottom = Y1 + Height1
    End With
    
    With Rect2
      .Left = x2
      .Top = y2
      .Right = x2 + Width2
      .Bottom = y2 + Height2
    End With
    'Check for rectangular collisions
     
        If IntersectRect(RECTOverlap, Rect1, Rect2) Then
            'For those who want faster performance on slower computers offer non pixel
            'perfect collision detection. Doesn't look as nice but is certainly a lot
            'faster.
            If PixelPerfect = False Then
              CheckCollide = True
              Exit Function
            End If
            'RECTANGULAR COLLISION
            'Get the RECT structures for the overlapped portions of both surfaces
            With RECT1Overlap 'Find the overlap difference in the first RECT
                .Top = RECTOverlap.Top - Rect1.Top
                .Bottom = RECTOverlap.Bottom - Rect1.Top
                .Right = RECTOverlap.Right - Rect1.Left
                .Left = RECTOverlap.Left - Rect1.Left
            End With
            
            With RECT2Overlap 'Find the overlap difference in the second RECT
                .Top = RECTOverlap.Top - Rect2.Top
                .Bottom = RECTOverlap.Bottom - Rect2.Top
                .Right = RECTOverlap.Right - Rect2.Left
                .Left = RECTOverlap.Left - Rect2.Left
            End With
            
            'Determine the width and height of the ovrelas (we will use this information for the loop)
            OverlapWidth = RECTOverlap.Right - RECTOverlap.Left - 1
            OverlapHeight = RECTOverlap.Bottom - RECTOverlap.Top - 1
            
            'Use Lock and GetLockedArray on each surface
            DDSurf1.Lock RECT1Overlap, DDSDBlank, DDLOCK_READONLY Or DDLOCK_WAIT, 0
            DDSurf1.GetLockedArray ByteObj1
            DDSurf2.Lock RECT2Overlap, DDSDBlank, DDLOCK_READONLY Or DDLOCK_WAIT, 0
            DDSurf2.GetLockedArray ByteObj2
            'Compare the surface data from the overlapping portions of the rectangles
            For i = 0 To OverlapWidth
                For j = 0 To OverlapHeight
                    'If BOTH surfaces are non-tranparent at this pixel...
                    If (ByteObj1(i + RECT1Overlap.Left, j + RECT1Overlap.Top) <> BlitCOLORKEY) And (ByteObj2(i + RECT2Overlap.Left, j + RECT2Overlap.Top) <> BlitCOLORKEY) Then PPCollision = True
                    'We have Pixel Perfect Collision
                    If PPCollision = True Then
                        CheckCollide = True
                        Exit For 'Exit because we don't need to check anymore, we already have pixel perfect collision
                    End If
                Next j
                If PPCollision = True Then
                    CheckCollide = True
                    Exit For 'Exit because we don't need to check anymore, we already have pixel perfect collision
                End If
            Next i
            
            'Unlock the sufaces
            DDSurf1.Unlock RECT1Overlap 'unlock DDsurf1
            DDSurf2.Unlock RECT2Overlap 'Unlock DDSurf2
        End If
End Function
