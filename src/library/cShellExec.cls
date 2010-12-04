VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cShellExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' API
'
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Private Const SW_HIDE = 0
Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_RESTORE = 9
Private Const SW_SHOW = 5
Private Const SW_SHOWDEFAULT = 10
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOWNORMAL = 1

Public Enum ShellExec_SW_Constants
    ShellExec_SW_HIDE = 0
    ShellExec_SW_MAXIMIZE = 3
    ShellExec_SW_MINIMIZE = 6
    ShellExec_SW_RESTORE = 9
    ShellExec_SW_SHOW = 5
    ShellExec_SW_SHOWDEFAULT = 10
    ShellExec_SW_SHOWMAXIMIZED = 3
    ShellExec_SW_SHOWMINIMIZED = 2
    ShellExec_SW_SHOWMINNOACTIVE = 7
    ShellExec_SW_SHOWNA = 8
    ShellExec_SW_SHOWNOACTIVATE = 4
    ShellExec_SW_SHOWNORMAL = 1
End Enum

Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_ACCESSDENIED = 5            '  access denied
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_OOM = 8                     '  out of memory
Private Const SE_ERR_SHARE = 26

Private Const STYLE_NORMAL = 11

Public Function Exec(ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As ShellExec_SW_Constants) As Long
Dim r As Long

    r = ShellExecute(hWnd, lpOperation, lpFile, lpParameters, lpDirectory, nShowCmd)
    
    Exec = r
End Function
