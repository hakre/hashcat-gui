VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "drhashEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''
'' Dr. Hash™
'' written 2009 by hakre
''
'' ....................................................................
''
'' drhashEntry
'' version 0.2a
'' drhash data model
''
'' represents a hash. based on CSV data/line  but fragile. see
'' drhashCollection.
''
Private m_Line As String
Private m_Cols() As String
Private m_Key As String
Public Property Get Col(iIndex As Long) As String
    Col = m_Cols(iIndex)
End Property
Public Property Get Cols() As String()
    Cols = m_Cols
End Property
Public Property Let Csv(sLine As String)
    m_Line = sLine
    m_Cols = zCsvParseLine(sLine)
End Property
Public Property Get Key() As String
    Key = m_Key
End Property
Public Property Let Key(sKey As String)
    m_Key = sKey
End Property
Public Property Get Title() As String
    Title = m_Cols(0)
End Property
'
' Parse a CSV line
'
Private Function zCsvParseLine(sCsvLine As String) As String()
Dim aCols() As String
Dim i As Long, c As Long

    aCols = Split(sCsvLine, ",")
        
    If UBound(aCols) < 0 Then
        ReDim aCol(0 To 2)
    ElseIf UBound(aCols) < 2 Then
        ReDim Preserve aCol(0 To 2)
    End If
    
    c = UBound(aCols)
    For i = 0 To c
        aCols(i) = zCsvUnescape(aCols(i))
    Next
    
    zCsvParseLine = aCols

End Function
'
' unescape a CSV value
'
Private Function zCsvUnescape(sData As String) As String
    If Len(sData) > 1 Then
        If Left(sData, 1) = """" And Right(sData, 1) = """" Then
            sData = Mid(sData, 2, Len(sData) - 2)
        End If
        If Left(sData, 1) = """" Then
            Debug.Print "Error: " & CStr(sData)
        End If
    End If
zCsvUnescape = sData
End Function
