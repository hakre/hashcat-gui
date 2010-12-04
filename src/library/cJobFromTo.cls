VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJobRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sValue As String
Private m_DefMin As Long
Private m_DefMax As Long
Private m_DefFrom As Long
Private m_DefTo As Long
Public Sub Init(iFrom As Long, iTo As Long, Optional iMin As Long = 0, Optional iMax As Long = 32767)

    'some little precautions:
    
    'negative numbers not supported
    If iMin < 0 Then
        iMin = 0
    End If
    
    'maximum can not be smaller then minimum
    If iMax < iMin Then
        iMax = iMin
    End If
    
    m_DefMin = iMin
    m_DefMax = iMax
        
    m_DefFrom = iFrom
    m_DefTo = iTo
    
    m_sValue = CStr(m_DefFrom) & "-" & CStr(m_DefTo)
    
End Sub
' is parameter iValue within the range of this range or not?
'
'return: true or false
Public Function InRange(iValue As Long) As Boolean

    'If iValue < Me.ValueFrom Then
    '    Exit Function
    'End If
    
    'If iValue > Me.ValueTo Then
    '    Exit Function
    'End If

    InRange = True
    
End Function


'
Public Property Get StringFrom() As String
Dim aString() As String
    aString = Split(m_sValue, "-", 2)
    StringFrom = aString(0)
End Property
Public Property Get StringTo() As String
Dim aString() As String
    aString = Split(m_sValue, "-", 2)
    StringTo = aString(1)
End Property
Public Property Get Value() As String
    Value = m_sValue
End Property
Public Property Let Value(sValue As String)
Dim aString() As String
    If sValue = "" Then
        Exit Property
    End If
    aString = Split(sValue, "-", 2)
    If UBound(aString) > 0 Then
        m_sValue = aString(0) & "-" & aString(1)
    End If
End Property
Public Property Let ValueFrom(iValue As Long)

    If iValue < m_DefMin Then
        iValue = m_DefMin
    End If
    
    If iValue > m_DefMax Then
        iValue = m_DefMax
    End If
    
    Me.Value = CStr(iValue) & "-" & Me.StringTo
    
End Property
Public Property Let ValueTo(iValue As Long)
    
    If iValue < m_DefMin Then
        iValue = m_DefMin
    End If
    
    If iValue > m_DefMax Then
        iValue = m_DefMax
    End If
    
    Me.Value = Me.StringFrom & "-" & CStr(iValue)
    
End Property
Private Sub Class_Initialize()
    m_DefMax = 32767
    m_sValue = "-"
End Sub

