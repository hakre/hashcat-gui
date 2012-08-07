Attribute VB_Name = "mApi"
Option Explicit

'
' window position and size
' http://vb.mvps.org/hardcore/html/windowpositionsize.htm
' http://www.ex-designz.net/apidetail.asp?api_id=380
'

Public Type POINTAPI
    x As Long
    y As Long
End Type

Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long


' http://vbnet.mvps.org/code/subclass/activation.htm
Public Const WM_ACTIVATE      As Long = &H6
Public Const WM_ACTIVATEAPP   As Long = &H1C
Public Const WA_INACTIVE      As Long = 0
Public Const WA_ACTIVE        As Long = 1
Public Const WA_CLICKACTIVE   As Long = 2


'
' toolbar windows helper
' http://www.vb-helper.com/howto_detect_activate.html
' http://vbnet.mvps.org/code/faq/floating.htm
'
Public Const GWL_WNDPROC = (-4)
Public Const GWL_HWNDPARENT = (-8)
Public Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" _
   (ByVal hWnd As Long, ByVal nIndex As Long, _
    ByVal wNewLong As Long) As Long

'
' get keystates
' http://www.ex-designz.net/apidetail.asp?api_id=143
'
Declare Function GetKeyboardState Lib "user32.dll" (pbKeyState As Byte) As Long

'
' get systemfolders
' http://vbnet.mvps.org/code/browse/shpathidlist.htm
'
Public Const CSIDL_PERSONAL = &H5                'My Documents
Public Const CSIDL_APPDATA = &H1A                '{user}\Application Data
Private Const MAX_PATH As Long = 260 'WINAPI limit
Private Const S_OK = 0

'Converts an item identifier list to a file system path.
Private Declare Function SHGetPathFromIDList Lib "shell32" _
   Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, _
   ByVal pszPath As String) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32" _
   (ByVal hwndOwner As Long, _
    ByVal nFolder As Long, _
    pidl As Long) As Long

Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
   

'http://www.vbforums.com/showthread.php?t=505137
'Windows API struct - exe version
Public Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
   dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
   dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
   dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
   dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
   dwFileFlagsMask As Long        '  = &h3F for version "0.42"
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type


'Windows API function declarations
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

'
' get systemfolders
'
Public Function WINAPI_GetSpecialFolderLocation(CSIDL As Long, Optional hWnd As Long) As String

   Dim sPath As String
   Dim pidl As Long
   
  'fill the idl structure with the specified folder item
   If SHGetSpecialFolderLocation(hWnd, CSIDL, pidl) = S_OK Then
     
     'if the pidl is returned, initialize
     'and get the path from the id list
      sPath = Space$(MAX_PATH)
      
      If SHGetPathFromIDList(ByVal pidl, ByVal sPath) Then

        'return the path
         WINAPI_GetSpecialFolderLocation = Left(sPath, InStr(sPath, Chr$(0)) - 1)
         
      End If
    
     'free the pidl
      Call CoTaskMemFree(pidl)

    End If
   
End Function




'exe file version
Public Function WINAPI_GetFileVersion(ByVal FileName As String, Optional iFormat As Long = 0) As String
   Dim nDummy As Long
   Dim sBuffer()         As Byte
   Dim nBufferLen        As Long
   Dim lplpBuffer       As Long
   Dim udtVerBuffer      As VS_FIXEDFILEINFO
   Dim puLen     As Long
      
   nBufferLen = GetFileVersionInfoSize(FileName, nDummy)
   
   If nBufferLen > 0 Then
   
        ReDim sBuffer(nBufferLen) As Byte
        Call GetFileVersionInfo(FileName, 0&, nBufferLen, sBuffer(0))
        Call VerQueryValue(sBuffer(0), "\", lplpBuffer, puLen)
        Call CopyMemory(udtVerBuffer, ByVal lplpBuffer, Len(udtVerBuffer))
        
        'Format: 0 long, 1 vb-style
        If iFormat = 0 Then
            WINAPI_GetFileVersion = udtVerBuffer.dwFileVersionMSh & "." & udtVerBuffer.dwFileVersionMSl & "." & udtVerBuffer.dwFileVersionLSh & "." & udtVerBuffer.dwFileVersionLSl
        Else
            WINAPI_GetFileVersion = udtVerBuffer.dwFileVersionMSh & "." & udtVerBuffer.dwFileVersionMSl & "." & udtVerBuffer.dwFileVersionLSl
        End If
    End If
End Function

