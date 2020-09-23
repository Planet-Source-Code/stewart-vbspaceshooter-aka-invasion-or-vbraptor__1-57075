Attribute VB_Name = "modTransparency"
'COLOR SHIFT VALUES
Public RedShiftLeft As Long
Public RedShiftRight As Long
Public GreenShiftLeft As Long
Public GreenShiftRight As Long
Public BlueShiftLeft As Long
Public BlueShiftRight As Long

Public Sub GetColorShiftValues(PrimarySurface As DirectDrawSurface7)
   
    Dim PixelFormat As DDPIXELFORMAT

    PrimarySurface.GetPixelFormat PixelFormat
    MaskToShiftValues PixelFormat.lRBitMask, RedShiftRight, RedShiftLeft
    MaskToShiftValues PixelFormat.lGBitMask, GreenShiftRight, GreenShiftLeft
    MaskToShiftValues PixelFormat.lBBitMask, BlueShiftRight, BlueShiftLeft
End Sub

Public Sub MaskToShiftValues(ByVal Mask As Long, ShiftRight As Long, ShiftLeft As Long)

    Dim ZeroBitCount As Long
    Dim OneBitCount As Long


    ' Count zero bits

    ZeroBitCount = 0
    Do While (Mask And 1) = 0
        ZeroBitCount = ZeroBitCount + 1
        Mask = Mask \ 2 ' Shift right
    Loop



    ' Count one bits

    OneBitCount = 0
    Do While (Mask And 1) = 1
        OneBitCount = OneBitCount + 1
        Mask = Mask \ 2 ' Shift right
    Loop

    ' Shift right 8-OneBitCount bits
    ShiftRight = 2 ^ (8 - OneBitCount)
    ' Shift left ZeroBitCount bits
    ShiftLeft = 2 ^ ZeroBitCount
End Sub
