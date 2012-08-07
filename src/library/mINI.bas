Attribute VB_Name = "mIni"
Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" _
Alias "WritePrivateProfileStringA" _
                        (ByVal lpApplicationName As String, _
                        ByVal lpKeyName As Any, _
                        ByVal lpString As Any, _
                        ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringA" _
                        (ByVal lpApplicationName As String, _
                        ByVal lpKeyName As Any, _
                        ByVal lpDefault As String, _
                        ByVal lpReturnedString As String, _
                        ByVal nSize As Long, _
                        ByVal lpFileName As String) As Long
Public Function INIWrite(ByVal sNewString As String, sSection As String, sKeyName As String, sINIFileName As String) As Boolean
  
  Call WritePrivateProfileString(sSection, sKeyName, sNewString, sINIFileName)
  INIWrite = (Err.Number = 0)

End Function
Public Function INIRead(sSection As String, sKeyName As String, sINIFileName As String) As String
Dim sRet As String
Dim sValue As String

    sRet = String(255, Chr(0))
    INIRead = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), sINIFileName))

End Function
Public Function INIReadInt(sSection As String, sKeyName As String, sINIFileName As String) As Long
Dim sVal As String
Dim dVal As Double

    sVal = INIRead(sSection, sKeyName, sINIFileName)
    dVal = Val(sVal)
    On Error Resume Next
    INIReadInt = Int(dVal)
    On Error GoTo 0
    
End Function
