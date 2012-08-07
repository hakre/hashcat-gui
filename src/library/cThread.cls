VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------
' Module CaptureConsoleOutput
'
' Purpose: Excute a console app. & capture output
'
' Author: Dipak Auddy
' E-mail: dauddy@gmail.com
'
' Date: Dec 01, 2004
'---------------------------------------------------
Option Explicit

' Public Events
Public Event StdOutWrite(ByVal sOutput As String)
Public Event StdErrWrite(ByVal sOutput As String)
Public Event Started()
Public Event Terminate()

' Private Members
Private WithEvents ProcTimer As TimerLite.clsTimerLite
Attribute ProcTimer.VB_VarHelpID = -1
Private ProcTimer_Fired As Boolean

Private m_sCommandLine As String
Private m_sStartInFolder As String
Private m_cTextBox As TextBox

Private m_hStdIn As Long
Private m_hStdOut As Long
Private m_hStdErr As Long
Private m_hStdInWrite As Long
Private m_hStdOutRead As Long
Private m_hStdErrRead As Long
Private m_hChildProcess As Long
Private m_hStdOutThread As Long
Private m_hStdErrThread As Long
Private m_hProcessThread As Long
Private m_hExitEvent As Long
Private m_bRunThread As Boolean

' Hinting Enums

Private Enum PriorityClass
    Realtime = REALTIME_PRIORITY_CLASS
    High = HIGH_PRIORITY_CLASS
    AboveNormal = ABOVE_NORMAL_PRIORITY_CLASS
    Normal = NORMAL_PRIORITY_CLASS
    BelowNormal = BELOW_NORMAL_PRIORITY_CLASS
    Idle = IDLE_PRIORITY_CLASS
End Enum

' API: Constants

Private Const REALTIME_PRIORITY_CLASS = &H100&
Private Const HIGH_PRIORITY_CLASS = &H80&
Private Const ABOVE_NORMAL_PRIORITY_CLASS = &H8000&
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const BELOW_NORMAL_PRIORITY_CLASS = &H4000&
Private Const IDLE_PRIORITY_CLASS = &H40&

Private Const CREATE_NEW_CONSOLE             As Long = &H10
Private Const DUPLICATE_SAME_ACCESS          As Long = &H2
Private Const SECURITY_DESCRIPTOR_MIN_LENGTH As Long = (20)
Private Const SECURITY_DESCRIPTOR_REVISION   As Long = 1

Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Const PIPE_ACCESS_DUPLEX = &H3
Private Const PIPE_READMODE_MESSAGE = &H2
Private Const PIPE_TYPE_MESSAGE = &H4
Private Const PIPE_WAIT = &H0

Private Const INVALID_HANDLE_VALUE = -1

Private Const STARTF_USESHOWWINDOW     As Long = &H1
Private Const STARTF_USESTDHANDLES     As Long = &H100
Private Const SW_HIDE                  As Integer = 0
Private Const SW_SHOW                  As Integer = 5

' API: Types

' ACL structure
Private Type ACL
   AclRevision As Byte
   Sbz1 As Byte
   AclSize As Integer
   AceCount As Integer
   Sbz2 As Integer
End Type

Private Type SECURITY_DESCRIPTOR
   Revision As Byte
   Sbz1 As Byte
   Control As Long
   Owner As Long
   Group As Long
   Sacl As ACL
   Dacl As ACL
End Type

Private Type SECURITY_ATTRIBUTES
    nLength                                As Long
    lpSecurityDescriptor                   As Long
    bInheritHandle                         As Long
End Type

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
    dwThreadID                             As Long
End Type

' Declares
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, _
                                                    phWritePipe As Long, _
                                                    lpPipeAttributes As Any, _
                                                    ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, _
                                                  lpBuffer As Any, _
                                                  ByVal nNumberOfBytesToRead As Long, _
                                                  lpNumberOfBytesRead As Long, _
                                                  lpOverlapped As Any) As Long
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

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" ( _
    ByVal pSecurityDescriptor As Long, _
    ByVal dwRevision As Long) As Long

Private Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" ( _
      ByVal pSecurityDescriptor As Long, _
      ByVal bDaclPresent As Long, _
      ByVal pDacl As Long, _
      ByVal bDaclDefaulted As Long) As Long

Private Declare Function GlobalAlloc Lib "kernel32" ( _
      ByVal wFlags As Long, ByVal dwBytes As Long) As Long
      
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function CreateEvent& Lib "kernel32" _
    Alias "CreateEventA" (ByVal lpEventAttributes As Long, _
    ByVal bManualReset As Long, ByVal bInitialState As Long, _
    ByVal lpname As String)
    
Public Declare Function CreateThread Lib "kernel32" ( _
    ByVal lpThreadAttributes As Any, ByVal dwStackSize As Long, _
    ByVal lpStartAddress As Long, lpParameter As Any, _
    ByVal dwCreationFlags As Long, lpThreadID As Long) As Long

' @todo add missing code
Private Function IsWinNT() As Long
    IsWinNT = 1
End Function

Private Function PrepAndLaunchRedirectedChild(lpszCmdLine As String, _
    hStdOut As Long, hStdIn As Long, hStdErr As Long, _
    bShowChildWindow As Long, _
    Optional sStartInFolder As String = vbNullString) As Long
    
Dim hProcess          As Long
Dim pi                As PROCESS_INFORMATION
Dim si                As STARTUPINFO
Dim lpSD              As Long
Dim sa                As SECURITY_ATTRIBUTES
Dim bResult           As Long
Dim lCreationFlags    As Long

    hProcess = GetCurrentProcess()
    
    ' Set up the start up info struct.
    
    si.cb = Len(si)
    si.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    si.hStdOutput = hStdOut
    si.hStdInput = hStdIn
    si.hStdError = hStdErr
    
    ' Use this if you want to show the child.
    If bShowChildWindow Then si.wShowWindow = SW_SHOW Else si.wShowWindow = SW_HIDE
    ' Note that dwFlags must include STARTF_USESHOWWINDOW if you want to
    ' use the wShowWindow flags.
    
    ' Create the NULL security token for the process
    lpSD = 0
    
    If IsWinNT Then ' assume we have winnt
        lpSD = GlobalAlloc(GPTR, SECURITY_DESCRIPTOR_MIN_LENGTH)
        If InitializeSecurityDescriptor(lpSD, SECURITY_DESCRIPTOR_REVISION) = 0 Then Stop
        If SetSecurityDescriptorDacl(lpSD, -1, 0, 0) = 0 Then Stop
        sa.nLength = LenB(sa)
        sa.lpSecurityDescriptor = lpSD
        sa.bInheritHandle = True
    Else
        Stop
        Exit Function 'needs to be null then but this ain't a pointer !
    End If
    
    lCreationFlags = CREATE_NEW_CONSOLE
    
    ' Try to spawn the process.
    bResult = CreateProcess(vbNullString, lpszCmdLine, sa, ByVal 0&, 1, _
                        lCreationFlags, ByVal 0&, sStartInFolder, si, pi)
    
    ' Cleanup memory allocation
    If lpSD <> 0 Then
        Call GlobalFree(lpSD)
    End If
    
    ' Close any unnecessary handles.
    If CloseHandle(pi.hThread) = 0 Then Stop

    ' Save global child process handle to cause threads to exit.
    PrepAndLaunchRedirectedChild = pi.hProcess
    
End Function
Public Sub ScheduleExec(ByVal sCommandLine As String, Optional ByVal sStartInFolder As String = vbNullString)

    m_sCommandLine = sCommandLine
    m_sStartInFolder = sStartInFolder
    Set m_cTextBox = Nothing
    
    ProcTimer_Fired = False
    Set ProcTimer = CreateObject("TimerLite.clsTimerLite")
    ProcTimer.Interval = 500
    ProcTimer.Enabled = True
    
End Sub
'---------------------------------------------------
' Call this sub to execute and capture a console app.
' Ex: Call ExecAndCapture("ping localhost", Text1)
Public Sub ExecAndCapture()
Dim sCommandLine As String
Dim sStartInFolder As String
Dim cTextBox As TextBox
Dim iCreationFlags As Long

sCommandLine = m_sCommandLine
sStartInFolder = m_sStartInFolder
Set cTextBox = m_cTextBox

Const BUFSIZE         As Long = 1024 * 10
Dim hPipeRead         As Long
Dim hPipeWrite        As Long
Dim sa                As SECURITY_ATTRIBUTES
Dim si                As STARTUPINFO
Dim pi                As PROCESS_INFORMATION
Dim baOutput(BUFSIZE) As Byte
Dim sOutput           As String
Dim lBytesRead        As Long
    
    With sa
        .nLength = Len(sa)
        .bInheritHandle = 1         ' get inheritable pipe handles
        .lpSecurityDescriptor = 0   ' same security descriptor then me
    End With 'SA
    
    If CreatePipe(hPipeRead, hPipeWrite, sa, 0) = 0 Then
        Exit Sub
    End If

    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = SW_SHOW
        .hStdOutput = hPipeWrite
        .hStdError = hPipeWrite
    End With 'SI
    
    iCreationFlags = PriorityClass.Idle Or CREATE_NEW_CONSOLE
    
    RaiseEvent Started
    
    If CreateProcess(vbNullString, sCommandLine, ByVal 0&, ByVal 0&, 1, iCreationFlags, ByVal 0&, sStartInFolder, si, pi) Then
        RaiseEvent Started
        Call CloseHandle(hPipeWrite)
        Call CloseHandle(pi.hThread)
        hPipeWrite = 0
        Do
            DoEvents: DoEvents: DoEvents
            If ReadFile(hPipeRead, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) = 0 Then
                Exit Do
            End If
            sOutput = Left$(StrConv(baOutput(), vbUnicode), lBytesRead)
            ' cTextBox.SelText = sOutput
            RaiseEvent StdOutWrite(sOutput)
        Loop
        Call CloseHandle(pi.hProcess)
    End If
    ' To make sure...
    Call CloseHandle(hPipeRead)
    Call CloseHandle(hPipeWrite)
    'Call CloseHandle(pi.hThread)
    'hPipeWrite = 0
    RaiseEvent Terminate
End Sub


Private Sub ProcTimer_Timer()
    If ProcTimer_Fired Then Exit Sub
    ProcTimer_Fired = True
    
    ProcTimer.Interval = 0
    ProcTimer.Enabled = False
    
    Set ProcTimer = Nothing
    
    ExecAndCapture
    
    ProcTimer_Fired = False
End Sub
'**
'* StartChildProcess
'*
'* @note port
'*
Private Function StartChildProcess(lpszCmdLine As String, bShowChildWindow As Long) As Boolean
Dim hProcess As Long
Dim sa As SECURITY_ATTRIBUTES
Dim hStdInWriteTmp As Long, hStdOutReadTmp As Long, hStdErrReadTmp As Long
Dim dwThreadID As Long

    hProcess = GetCurrentProcess()
    
    sa.nLength = Len(sa)
    sa.lpSecurityDescriptor = Null
    sa.bInheritHandle = True
    
    ' Create the child stdin pipe.
    If CreatePipe(m_hStdIn, hStdInWriteTmp, sa, 0) = 0 Then Stop
        
    ' Create the child stdout pipe.
    If CreatePipe(hStdOutReadTmp, m_hStdOut, sa, 0) = 0 Then Stop
    
    ' Create the child stderr pipe.
    If CreatePipe(hStdErrReadTmp, m_hStdErr, sa, 0) = 0 Then Stop
    
    ' Create new stdin write, stdout and stderr read handles.
    ' Set the properties to FALSE. Otherwise, the child inherits the
    ' properties and, as a result, non-closeable handles to the pipes
    ' are created.
    
    If DuplicateHandle(hProcess, hStdInWriteTmp, _
        hProcess, m_hStdInWrite, 0, 0, DUPLICATE_SAME_ACCESS) = 0 Then Stop

    If DuplicateHandle(hProcess, hStdOutReadTmp, _
        hProcess, m_hStdOutRead, 0, 0, DUPLICATE_SAME_ACCESS) = 0 Then Stop

    If DuplicateHandle(hProcess, hStdErrReadTmp, _
        hProcess, m_hStdErrRead, 0, 0, DUPLICATE_SAME_ACCESS) = 0 Then Stop
        
    ' Close inheritable copies of the handles you do not want to be
    ' inherited.
    
    If CloseHandle(hStdInWriteTmp) = 0 Then Stop
    If CloseHandle(hStdOutReadTmp) = 0 Then Stop
    If CloseHandle(hStdErrReadTmp) = 0 Then Stop
    
    ' Start child process with redirected stdout, stdin & stderr
    m_hChildProcess = PrepAndLaunchRedirectedChild(lpszCmdLine, _
        m_hStdOut, m_hStdIn, m_hStdErr, bShowChildWindow)
    
    If m_hChildProcess = 0 Then
        ' close all handles and return False
        If CloseHandle(m_hStdIn) = 0 Then Stop
        m_hStdIn = 0
        If CloseHandle(m_hStdOut) = 0 Then Stop
        m_hStdOut = 0
        If CloseHandle(m_hStdErr) = 0 Then Stop
        m_hStdErr = 0

        StartChildProcess = False
        
        Exit Function
        
    End If
    
    
    m_bRunThread = True

    ' Create Exit event
    m_hExitEvent = CreateEvent(ByVal 0&, True, False, ByVal 0&)
    
    If m_hExitEvent = 0 Then Stop
    
    ' Launch the thread that read the child stdout.
    m_hStdOutThread = CreateThread(ByVal 0&, ByVal 0&, AddressOf staticStdOutThread, ByVal 0&, 0, dwThreadID)  ' (LPVOID)this
    If m_hStdOutThread = 0 Then Stop

    ' Launch the thread that read the child stderr.
    m_hStdErrThread = CreateThread(ByVal 0&, ByVal 0&, AddressOf staticStdErrThread, ByVal 0&, 0, dwThreadID) ' (LPVOID)this
    If m_hStdErrThread = 0 Then Stop

    ' Launch the thread that monitoring the child process.
    m_hProcessThread = CreateThread(ByVal 0&, ByVal 0&, AddressOf staticProcessThread, ByVal 0&, 0, dwThreadID) '        (LPVOID)this
    If m_hProcessThread = 0 Then Stop

    ' Virtual function to notify derived class that the child is started.
            

    StartChildProcess = True
    
End Function
