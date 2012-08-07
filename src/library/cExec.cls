VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const STILL_ACTIVE& = &H103&
Private Const INFINITE = -1&

Private Const STARTF_USESHOWWINDOW     As Long = &H1
Private Const STARTF_USESIZE           As Long = &H2
Private Const STARTF_USESTDHANDLES     As Long = &H100
Private Const STARTF_USECOUNTCHARS     As Long = &H8

Private Const DETACHED_PROCESS         As Long = &H8
Private Const CREATE_NEW_CONSOLE       As Long = &H10

Private Type STARTUPINFO
    cb                                     As Long
    lpReserved                             As String
    lpDesktop                              As String
    lpTitle                                As String
    dwX                                    As Long
    dwY                                    As Long
    dwXSize                                As Long
    dwYSize                                As Long
    dwXCountChars                          As Long
    dwYCountChars                          As Long
    dwFillAttribute                        As Long
    dwFlags                                As Long
    wShowWindow                            As Integer
    cbReserved2                            As Integer
    lpReserved2                            As Long
    hStdInput                              As Long
    hStdOutput                             As Long
    hStdError                              As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess                               As Long
    hThread                                As Long
    dwProcessId                            As Long
    dwThreadId                             As Long
End Type

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, _
                                                                              ByVal lpCommandLine As String, _
                                                                              lpProcessAttributes As Any, _
                                                                              lpThreadAttributes As Any, _
                                                                              ByVal bInheritHandles As Long, _
                                                                              ByVal dwCreationFlags As Long, _
                                                                              lpEnvironment As Any, _
                                                                              ByVal lpCurrentDriectory As String, _
                                                                              lpStartupInfo As STARTUPINFO, _
                                                                              lpProcessInformation As PROCESS_INFORMATION) As Long

Private Const SW_HIDE            As Integer = 0
Private Const SW_SHOWNORMAL      As Integer = 1
Private Const SW_NORMAL          As Integer = 1
Private Const SW_SHOWMINIMIZED   As Integer = 2
Private Const SW_SHOWMAXIMIZED   As Integer = 3
Private Const SW_MAXIMIZE        As Integer = 3
Private Const SW_SHOWNOACTIVATE  As Integer = 4
Private Const SW_SHOW            As Integer = 5
Private Const SW_MINIMIZE        As Integer = 6
Private Const SW_SHOWMINNOACTIVE As Integer = 7
Private Const SW_SHOWNA          As Integer = 8
Private Const SW_RESTORE         As Integer = 9
Private Const SW_SHOWDEFAULT     As Integer = 10
Private Const SW_FORCEMINIMIZE   As Integer = 11
Private Const SW_MAX             As Integer = 11

Private m_sCommandLine As String
Private m_sStartInFolder As String
Private m_iCreationFlags As Long
Private m_iPriorityClass As Long
Private Sub Class_Initialize()
    m_sStartInFolder = vbNullString
    m_iCreationFlags = 0
End Sub
Public Function ExecConsole(Optional sTitle As String = vbNullString, Optional bStayOpen As Boolean = False) As Long
Dim i As Long
Dim sApplicationName As String
Dim sCommandLine As String
Dim sStartInFolder As String
Dim iCreationFlags As Long
Dim si As STARTUPINFO
Dim pi As PROCESS_INFORMATION

    sApplicationName = vbNullString
    sCommandLine = m_sCommandLine
    sStartInFolder = m_sStartInFolder
    If Len(sStartInFolder) = 0 Then
        sStartInFolder = vbNullString
    End If
    
    iCreationFlags = m_iCreationFlags Or m_iPriorityClass
    
    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESHOWWINDOW ' Or STARTF_USECOUNTCHARS Or STARTF_USESIZE
        .wShowWindow = SW_NORMAL
        .lpTitle = sTitle
    End With 'SI
    
    If bStayOpen Then
        ' TODO add an additional class flag to echo the commandline if needed (debugmode or so)
        ' sCommandLine = "CMD /K echo " & sCommandLine & " & echo. & " & sCommandLine
        sCommandLine = "CMD /K " & sCommandLine
    End If
    
    i = CreateProcess(sApplicationName, sCommandLine, ByVal 0&, ByVal 0&, 1, iCreationFlags, ByVal 0&, sStartInFolder, si, pi)

End Function
Public Property Let CommandLine(sCommandLine As String)
    m_sCommandLine = sCommandLine
End Property
Public Property Get CommandLine() As String
    CommandLine = m_sCommandLine
End Property
Public Property Let CreationFlags(iCreationFlags As Long)
    m_iCreationFlags = iCreationFlags
End Property
Public Property Let PriorityClass(iPriorityClass As PriorityClass)
    m_iPriorityClass = iPriorityClass
End Property
Public Property Get PriorityClass() As PriorityClass
    PriorityClass = m_iPriorityClass
End Property
Public Property Let StartInFolder(sStartInFolder As String)
    m_sStartInFolder = sStartInFolder
End Property
Public Property Get StartInFolder() As String
    StartInFolder = m_sStartInFolder
End Property
