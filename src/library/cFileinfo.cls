VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Win API
Private Declare Function GetFullPathName Lib "kernel32.dll" _
  Alias "GetFullPathNameA" ( _
  ByVal lpFileName As String, _
  ByVal nBufferLength As Long, _
  ByVal lpBuffer As String, _
  ByVal lpFilePart As String) As Long

' VB Stuff
Private m_Path As String


'
' TODO it might
'
' @return Empty string if this dir/filename's dir does not exists or the dir/filename's directory
'
Public Property Get ExistingDir() As String

    ExistingDir = ""
    
    
    If Me.Exists Then
        'if file exists return dirname
        ExistingDir = Me.Dirname
    ElseIf zFileExists(Me.Dirname) Then
        'if dirname exists return dirname
        ExistingDir = Me.Dirname
    End If
    
End Property
' get files if this is a directory
Public Property Get Files() As Collection
    
    Set Files = zGetDir(m_Path)

End Property
' get full path name of a file
Public Property Get FullPath() As String
Dim r As Long
Dim lpBuffer As String * 256
  
    r = GetFullPathName(m_Path, Len(lpBuffer), lpBuffer, "")
     
    If r <> 0 Then
        FullPath = Left(lpBuffer, r)
    End If

End Property
Public Property Let Path(sPath As String)

    m_Path = sPath

End Property

Public Property Get Path() As String

    Path = m_Path
    
End Property

Public Function RelPath(sTo As String) As String
Dim oFi As New cFileinfo
Dim sBase As String

    oFi.Path = sTo
    sBase = oFi.Dirname
    
    If Left(Me.Dirname, Len(sBase)) = sBase Then
        RelPath = Mid(Me.Path, Len(sBase) + 1)
    Else
        RelPath = Me.Path
    End If
        
    
End Function
Public Property Get Basename() As String
Dim pos As Long
    
    pos = InStrRev(m_Path, "\")
    
    If pos = 0 Then
        Basename = m_Path
        Exit Property
    End If
    
    Basename = Mid(m_Path, pos + 1)

End Property
' return value has always \ at the end
Public Property Get Dirname() As String
Dim pos As Long

    ' if directory phyiscally exists, there is no need to guess anything
    If Me.isDir Then
        Dirname = m_Path
        If Mid(m_Path, Len(m_Path), 1) <> "\" Then
            Dirname = Dirname & "\"
        End If
        Exit Property
    End If

    ' does not exists, therefore  handle as intended
    
    pos = InStrRev(m_Path, "\")
    
    Dirname = Left(m_Path, pos)

End Property

'
'
Public Property Get Exists() As Boolean
    Exists = zFileExists(m_Path)
End Property
Public Property Get isDir() As Boolean
    
    isDir = zIsDir(m_Path)

End Property
Public Property Get isFile() As Boolean

    isFile = False

    If Me.Exists And Not Me.isDir Then
        isFile = True
    End If
    
End Property
Public Property Get Lines() As Long
    If Me.isFile Then
        Dim oTextfile As New cFileText
        oTextfile.Path = m_Path
        Lines = oTextfile.LineCount
        'Lines = zLinesInFileWin(m_Path)
    End If
End Property
Public Property Get ChangeTime() As Variant
    If Me.isFile Then
        ChangeTime = FileDateTime(m_Path)
    Else
        ChangeTime = 0
    End If
End Property
Public Property Get Size() As Long
    If Me.isFile Then
        Size = FileLen(m_Path)
    End If
End Property
Private Function zFileExists(sFile As String) As Boolean
Dim iAtts As Long

    On Error Resume Next
    
        iAtts = GetAttr(sFile)
        
        If Err.Number <> 0 Then
            If Err.Number = 5 Or Err.Number = 52 Or Err.Number = 53 Or Err.Number = 76 Then
                ' 5 : Invalid Parameter (?)
                ' 52: File Name or Number wrong
                ' 53: Path not found
                ' 76: Path not found
            Else
                MsgBox "Unexpected Error: " & Err.Description, vbCritical, "Error #" & Err.Number
            End If
            zFileExists = False
        Else
            zFileExists = True
        End If
        
    On Error GoTo 0
    
End Function
Private Function zGetDir(ByVal sPath As String) As Collection
Dim newCol As New Collection
Dim oGlob As New cFileinfo
Dim sDir As String
Dim sName As String
Dim iCount As Long
Dim newFileinfo As cFileinfo


    If zIsDir(sPath) Then
        oGlob.Path = sPath
        sDir = oGlob.Dirname
        sName = Dir(sDir)
        Do While sName <> ""
            If sName <> "." And sName <> ".." Then
                Set newFileinfo = New cFileinfo
                newFileinfo.Path = oGlob.Dirname & sName
                Call newCol.Add(newFileinfo)
                Set newFileinfo = Nothing
            End If
            
            'protect against endless loops (4096 files are supported at least)
            iCount = iCount + 1
            If iCount > 4096 Then
                Exit Do
            End If
            
            sName = Dir   'get next entry
        Loop
    End If
    
Set zGetDir = newCol
End Function


Private Function zIsDir(sPath As String) As Boolean

    On Error Resume Next
        If (GetAttr(sPath) And vbDirectory) Then
            If Err.Number = 0 Then
                zIsDir = True
            End If
        End If
    On Error GoTo 0
    
End Function
Private Function zLinesInFileWin(ByVal sPath As String) As Long
Dim fnum As Integer
Dim Lines As Long
Dim one_line As String

    fnum = FreeFile
    Open sPath For Input Access Read Shared As #fnum
    Do While Not EOF(fnum)
        Line Input #fnum, one_line
        Lines = Lines + 1
    Loop
    Close fnum
    
zLinesInFileWin = Lines
End Function
