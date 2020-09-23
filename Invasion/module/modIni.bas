Attribute VB_Name = "modIni"
Option Explicit

'+------------------------------------------------------------------+
'| Invasion - modIni.bas                                            |
'+------------------------------------------------------------------+
'| Design and code by Stewart (sobert81@devedit.com)                |
'+------------------------------------------------------------------+

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long

Function ReadINI(Section, KeyName, Filename As String) As String
    Dim sRet As String
    sRet = String(10000, Chr(0)) '10000 character max. Should hopefully be enough :)
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), Filename))
End Function

Function writeini(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function
