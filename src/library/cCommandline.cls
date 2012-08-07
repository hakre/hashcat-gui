VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCommandline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_CommandRaw As String
'
' exists passedFile?
'
Public Function existsPassedFile() As Boolean
Dim oFi As New cFileinfo
Dim sFile As String

    sFile = Me.passedFile
    If Len(sFile) Then
        oFi.Path = sFile
        If oFi.Exists Then
            existsPassedFile = True
        End If
    End If
End Function
'
' single parameter pattern that is a file und gut is
'
Public Property Get passedFile() As String
Dim s As String

    s = m_CommandRaw
    If Len(s) Then
        s = Trim(s) 'remove trailing spaces that sometimes are passed by...
        'remove enclosing quotes if any
        If Len(s) > 1 Then
            If Left(s, 1) = """" And Right(s, 1) = """" Then
                If Len(s) = 2 Then
                    s = ""
                Else
                    s = Mid(s, 2, Len(s) - 2)
                End If
            End If
        End If
        passedFile = s
    End If
    
End Property

Private Sub Class_Initialize()

    m_CommandRaw = Command()

End Sub


