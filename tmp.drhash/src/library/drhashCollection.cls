VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "drhashCollection"
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
'' drhashCollection
'' version 0.2a
'' drhash data model
''
'' represents a collection of hashes. based on CSV data but fragile.
'' seperator char is ",", never use "," within data itself even when
'' put inside "".
''
Private m_Csv As String
Private m_ColEntries As Collection
Public Property Get Csv() As String
Attribute Csv.VB_MemberFlags = "40"
    Csv = m_Csv
End Property
'set this class drhash csv data
Public Property Let Csv(sCsv As String)
    m_Csv = sCsv
    Set m_ColEntries = zCsvParseFile(sCsv)
End Property
Public Property Get Item(vIndex As Variant) As drhashEntry
    Set Item = m_ColEntries.Item(vIndex)
End Property
'load a textfile
'
' returns the number of lines read or 0 on error
Public Function LoadFile(sPath As String) As Long
Dim sContent As String
Dim sLine As String
Dim iFn As Long
Dim iCount As Long

    On Error Resume Next
        iFn = FreeFile()
        Open sPath For Input As #iFn
            Do While Not EOF(iFn)
                Line Input #iFn, sLine
                sContent = sContent & sLine & vbCrLf
                iCount = iCount + 1
            Loop
        Close #iFn
        
    If Err.Number = 0 Then
        Me.Csv = sContent
        LoadFile = iCount
    End If
    
End Function
Public Sub LoadRes(Id As Variant)
    On Error Resume Next
        Me.Csv = StrConv(LoadResData(Id, "CUSTOM"), vbUnicode)
End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    ' delegate to the private collection
    Set NewEnum = m_ColEntries.[_NewEnum]
End Function
Public Function SearchTitle(sTitle As String) As Collection
Dim newCol As New Collection
Dim oEntry As drhashEntry

    For Each oEntry In Me
        If InStr(1, oEntry.Title, sTitle, vbTextCompare) Then
            newCol.Add oEntry
        End If
    Next
    
Set SearchTitle = newCol
End Function
'
' parse csv file into a collection of drhashEtnry objects
'
Private Function zCsvParseFile(sCsv As String) As Collection
Dim newCol As New Collection
Dim aLines() As String
Dim c As Long
Dim i As Long
Dim sLine As String
Dim newEntry As drhashEntry
Dim sKey As String

    aLines = Split(sCsv, vbCrLf)
    c = UBound(aLines)
    For i = 1 To c
        If Len(aLines(i)) Then
            sKey = "h" + CStr(i)
            Set newEntry = New drhashEntry
            newEntry.Key = sKey
            newEntry.Csv = aLines(i)
            Call newCol.Add(newEntry, sKey)
        End If
    Next i
    
Set zCsvParseFile = newCol
End Function
Private Sub Class_Initialize()
    Set m_ColEntries = New Collection
End Sub
