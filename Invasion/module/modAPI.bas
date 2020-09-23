Attribute VB_Name = "modAPI"
Option Explicit

'+------------------------------------------------------------------+
'| Invasion - modAPI.bas                                            |
'+------------------------------------------------------------------+
'| Design and code by Stewart (sobert81@devedit.com)                |
'+------------------------------------------------------------------+
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Firing As Boolean
