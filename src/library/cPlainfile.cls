VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPlainfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Filename As String
Private m_Checked As Boolean
Public Property Get Checked() As Boolean
    Checked = m_Checked
End Property
Public Property Let Checked(bValue As Boolean)
    m_Checked = bValue
End Property
Public Property Get Filename() As String
    Filename = m_Filename
End Property
Public Property Let Filename(sFile As String)
    m_Filename = sFile
End Property
Public Property Get Fileinfo() As cFileinfo
Dim oFi As New cFileinfo
    oFi.Path = m_Filename
    Set Fileinfo = oFi
End Property
'import a string (oldskool ini notation (int)$flag,(string)$filenam
Public Property Let External(sData As String)
Dim oTemp As New cJobOptional
    oTemp.Use = m_Checked
    oTemp.External = sData
    m_Checked = oTemp.Use
    m_Filename = oTemp.Value
End Property
Public Property Get External() As String
Dim oTemp As New cJobOptional
    oTemp.Value = m_Filename
    oTemp.Use = m_Checked
    External = oTemp.External
End Property
Public Property Get Short() As String
    Short = Me.Fileinfo.Basename
End Property
