VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJobOptional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''
'' Optional Data Value Model (Class representing a Value that has a Use flag)
''
'' Use (Boolean)
'' Value (String)
'' External (String) - Old Ini Notation
''

Private m_Value As String
Private m_Use As Boolean

Public Property Get Use() As Boolean
    Use = m_Use
End Property
Public Property Let Use(bValue As Boolean)
    m_Use = bValue
End Property
Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property
Public Property Let Value(sValue As String)
    m_Value = sValue
End Property
Public Property Get External() As String
    External = CStr(Abs(m_Use)) & "," & m_Value
End Property
'import a string (oldskool ini notation (int)$flag,(string)$filenam
Public Property Let External(sData As String)
    If Mid(sData, 2, 1) = "," Then
        If Left(sData, 1) = "1" Then
            m_Use = True
        Else
            m_Use = False
        End If
        m_Value = Mid(sData, 3)
    Else
        m_Value = sData
    End If
End Property
