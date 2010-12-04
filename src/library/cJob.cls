VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''
'' Modell for a hashcat Job (configuration values needed for hashcat)
''

Private m_RecoveryMode As Long
Private m_BruteChars As String
Private m_BruteLen As New cJobRange
Private m_HashFile As String
Private m_HashMode As Long
Private m_HashSeperator As String
Private m_Limit As Variant
Private m_OutFile As New cJobOptional
Private m_Plains As cPlains
Private m_RuleCount As String
Private m_RuleFile As String
Private m_RuleMode As Long
Private m_SaltFile As New cJobOptional
Private m_Segment As Long
Private m_Skip As Variant
Private m_Threads As Long
Private m_ToggleLen As New cJobRange
Public Property Get BruteChars() As String
    BruteChars = m_BruteChars
End Property
Public Property Let BruteChars(sChars As String)
    m_BruteChars = sChars
End Property
Public Property Get BruteLen() As cJobRange
    Set BruteLen = m_BruteLen
End Property
Public Property Get HashFile() As String
    HashFile = m_HashFile
End Property
Public Property Let HashFile(sValue As String)
    m_HashFile = sValue
End Property
Public Property Get hashSeperator() As String
    hashSeperator = m_HashSeperator
End Property
Public Property Let hashSeperator(sValue As String)
    If Len(sValue) > 1 Then sValue = Left(sValue, 1)
    If Len(sValue) = 0 Then sValue = ":"
    m_HashSeperator = sValue
End Property
Public Property Let Limit(iValue As Variant)
    iValue = Cnv_Str2Dec(iValue)
    If TypeName(iValue) = "Decimal" Then
        If iValue < 0 Then iValue = 0
        If iValue > CDec("18446744073709551615") Then iValue = CDec("18446744073709551615")
    End If
    m_Limit = iValue
End Property
Public Property Get Limit() As Variant
    Limit = m_Limit
End Property
Public Property Get RecoveryMode() As Long
    RecoveryMode = m_RecoveryMode
End Property
Public Property Let RecoveryMode(iValue As Long)
    If iValue < 0 Then iValue = 0
    If iValue > 3 Then iValue = 3
    m_RecoveryMode = iValue
End Property
Public Property Get HashMode() As Long
    HashMode = m_HashMode
End Property
Public Property Let HashMode(iMode As Long)
    If iMode < 0 Then iMode = 0
    If iMode > 32767 Then iMode = 32767
    m_HashMode = iMode
End Property
Public Property Get OutFile() As cJobOptional
    Set OutFile = m_OutFile
End Property
Public Property Set OutFile(oValue As cJobOptional)
    Set m_OutFile = oValue
End Property
Public Property Get Plains() As cPlains
    If m_Plains Is Nothing Then
        Set m_Plains = New cPlains
    End If
    Set Plains = m_Plains
End Property
Public Property Set Plains(oPlains As cPlains)
    Set m_Plains = oPlains
End Property
Public Property Get RuleCount() As String
    RuleCount = m_RuleCount
End Property
Public Property Let RuleCount(sValue As String)
    m_RuleCount = sValue
End Property
Public Property Get RuleFile() As String
    RuleFile = m_RuleFile
End Property
Public Property Let RuleFile(sValue As String)
    m_RuleFile = sValue
End Property
Public Property Get RuleMode() As Long
    RuleMode = m_RuleMode
End Property
Public Property Let RuleMode(iValue As Long)
    If iValue < 0 Then iValue = 0
    If iValue > 2 Then iValue = 2
    m_RuleMode = iValue
End Property
Public Property Get Segment() As Long
    Segment = m_Segment
End Property
Public Property Let Segment(iValue As Long)
    If iValue < 1 Then iValue = 32
    m_Segment = iValue
End Property
Public Property Get SaltFile() As cJobOptional
    Set SaltFile = m_SaltFile
End Property
Public Property Let Skip(iValue As Variant)
    iValue = Cnv_Str2Dec(iValue)
    If TypeName(iValue) = "Decimal" Then
        If iValue < 0 Then iValue = 0
        If iValue > CDec("18446744073709551615") Then iValue = CDec("18446744073709551615")
    End If
    m_Skip = iValue
End Property
Public Property Get Skip() As Variant
    Skip = m_Skip
End Property
Public Property Get Threads() As Long
    Threads = m_Threads
End Property
Public Property Let Threads(iValue As Long)
    If iValue < 1 Then iValue = 8
    m_Threads = iValue
End Property
Public Property Get ToggleLen() As cJobRange
    Set ToggleLen = m_ToggleLen
End Property
'initialize
Private Sub Class_Initialize()
    m_Threads = 8
    m_Segment = 32
    m_HashSeperator = ":"
    Call m_BruteLen.Init(1, 8, 1, 1024)
    Call m_ToggleLen.Init(1, 16, 1, 1024)
End Sub
Private Sub Class_Terminate()
    Set m_OutFile = Nothing
    Set m_Plains = Nothing
End Sub
