VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' Text File abstraction (Focus on hashcat Outfile (basically UNIX) but can be Fuzzy mixed crlf / lf files)
'
'
Private m_Path As String
'
'Private m_LIdx_Pos(0 To 0) As Long
'Private m_LIdx_Max As Long
'Private m_LIdx_Buffer As Long

Public Property Get Data() As String
    Data = zLinesGetFuzzy(m_Path)
End Property

'returns a string array of all the lines
Public Property Get LinesRaw() As String()
    
    LinesRaw = Split(Me.Data, vbCrLf)

End Property
Public Property Get Path() As String
    Path = m_Path
End Property
Public Property Let Path(sPath As String)
    m_Path = sPath
End Property

Public Property Get LineCount() As Long

    LineCount = zLinesInFileUnix(m_Path)

End Property
'**********************************************
'PURPOSE: Returns Number of occurrences of a character or
'or a character sequencence within a string
'
' http://www.freevbcode.com/ShowCode.Asp?ID=1025
'
'PARAMETERS:
'   OrigString: String to Search in
'   Chars: Character(s) to search for
'   iLen: Length to search up to (limit how far it should be searched)
'   CaseSensitive (Optional): Do a case sensitive search
'   Defaults to false
'
'RETURNS:
'   Number of Occurrences of Chars in OrigString
'
'EXAMPLES:
'Debug.Print CharCount("FreeVBCode.com", "E") -- returns 3
'Debug.Print CharCount("FreeVBCode.com", "E", True) -- returns 0
'Debug.Print CharCount("FreeVBCode.com", "co") -- returns 2
''**********************************************
Private Function zCharCount(OrigString As String, Chars As String, Optional iLen As Long = -1, Optional CaseSensitive As Boolean = False) As Long

Dim lLen As Long
Dim lCharLen As Long
Dim lAns As Long
Dim sInput As String
Dim sChar As String
Dim lCtr As Long
Dim lEndOfLoop As Long
Dim bytCompareType As Byte

    sInput = OrigString
    If sInput = "" Then Exit Function
    If iLen < 0 Then
        lLen = Len(sInput)
    Else
        lLen = iLen
        If lLen > Len(sInput) Then
            lLen = Len(sInput)
        End If
    End If

    lCharLen = Len(Chars)
    lEndOfLoop = (lLen - lCharLen) + 1
    bytCompareType = IIf(CaseSensitive, vbBinaryCompare, vbTextCompare)

    For lCtr = 1 To lEndOfLoop
        sChar = Mid(sInput, lCtr, lCharLen)
        If StrComp(sChar, Chars, bytCompareType) = 0 Then
            lAns = lAns + 1
        End If
    Next

    zCharCount = lAns
End Function
' returns true/false if a file exists
Private Function zFileExists(sPath As String) As Boolean
Dim oFi As New cFileinfo
Dim bReturn As Boolean

    oFi.Path = sPath
    bReturn = oFi.Exists
    zFileExists = bReturn

End Function

'
'returns: string full of lines vbcrlf seperated
Private Function zLinesGetFuzzy(ByVal sPath As String) As String
Dim fnum As Integer
Dim Lines As Long
Dim one_line As String
Dim sBuffer As String
Dim iBufferLen As Long
Dim iFrom As Long, iTo As Long, iLen As Long
Dim sData As String
Dim sNewData As String

    iBufferLen = 1048576 '819200

    sBuffer = String(iBufferLen, " ")
    Lines = 1
    fnum = FreeFile
    Open sPath For Binary Access Read Shared As #fnum
    Do While Not EOF(fnum)
        iFrom = Loc(fnum)
        Get #fnum, , sBuffer
        If EOF(fnum) Then
            iTo = LOF(fnum)
        Else
            iTo = Loc(fnum)
        End If
        iLen = iTo - iFrom
        
        Lines = Lines + zCharCount(sBuffer, Chr(10), iLen + 2)
        
        sNewData = Left(sBuffer, iLen) 'get valuable parts out of the buffer
        sNewData = Replace(sNewData, vbCrLf, vbLf) 'do the fuzzy replacement for vbcrlf
        sNewData = Replace(sNewData, vbLf, vbCrLf) 'convert to windows standard
        
        sData = sData & sNewData  'concat data
    Loop
    Close fnum
    
zLinesGetFuzzy = sData
End Function

'
' Count number of lines in a file (UNIX)
'
' Empty lines at the end are not counted
'
Private Function zLinesInFileUnix(ByVal sPath As String) As Long
Dim fnum As Integer
Dim Lines As Long
Dim one_line As String
Dim sBuffer As String
Dim iBufferLen As Long
Dim iFrom As Long, iTo As Long, iLen As Long

    If Not zFileExists(sPath) Then
        Exit Function
    End If

    iBufferLen = 8192

    sBuffer = String(iBufferLen, " ")
    Lines = 0
    fnum = FreeFile
    
    Open sPath For Binary Access Read Shared As #fnum
    
    If LOF(fnum) > 0 Then
        Lines = 1
    End If
    
    Do While Not EOF(fnum)
        iFrom = Loc(fnum)
        Get #fnum, , sBuffer
        If EOF(fnum) Then
            iTo = LOF(fnum)
        Else
            iTo = Loc(fnum)
        End If
        iLen = iTo - iFrom
        Lines = Lines + zCharCount(sBuffer, Chr(10), iLen + 2)
    Loop
    
    'do not count as additional line if file ends with chr(10) x0a
    If iLen > 0 Then
        If Mid(sBuffer, iLen, 1) = Chr(10) Then
            Lines = Lines - 1
        End If
    End If
    
    Close fnum
    
zLinesInFileUnix = Lines
End Function
'
' Count number of lines in a file (WIN)
'
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
