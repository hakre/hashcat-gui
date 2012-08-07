VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPlains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_PrivateCollection As Collection
Private m_NoDupes As Boolean
'Return a comma-seperated list of basenames
Public Property Get Names() As String
Dim sNames As String
Dim oFile As cPlainfile

    For Each oFile In m_PrivateCollection
        If Len(sNames) Then
            sNames = sNames & ", "
        End If
        sNames = sNames & oFile.Fileinfo.Basename
    Next

Names = sNames
End Property

Public Property Let NoDupes(forbidDupes As Boolean)
    m_NoDupes = forbidDupes
End Property

Public Property Get NoDupes() As Boolean
    NoDupes = m_NoDupes
End Property

Private Sub Class_Initialize()
    ' explicit assignment is slightly faster than auto-instancing
    Set m_PrivateCollection = New Collection
End Sub
Public Sub Add(newItem As cPlainfile, Optional Key As Variant)

    If m_NoDupes Then
        'duplicates filenames do not need to be added
        If COL_objSearchString(m_PrivateCollection, newItem.Filename, "Filename") = 0 Then
            'check compare for certain property
            m_PrivateCollection.Add newItem, Key
        End If
    Else
        m_PrivateCollection.Add newItem, Key
    End If

End Sub
Public Property Get Caption() As String
Dim sCaption As String

    sCaption = CStr(Me.Count) & " File"
    Select Case Me.Count
        Case 0:
            sCaption = sCaption & "s."
        Case 1:
            sCaption = sCaption & ": " & Me.Names & "."
        Case Else:
            sCaption = sCaption & "s: " & Me.Names & "."
    End Select
    
    Caption = sCaption
End Property
Public Sub Clear()
    Set m_PrivateCollection = New Collection
End Sub
Property Get Count() As Long
    Count = m_PrivateCollection.Count
End Property

Public Sub Import(sData As String, Optional Key As Variant)
Dim newItem As New cPlainfile
    newItem.External = sData
    m_PrivateCollection.Add newItem, Key
End Sub
Function Item(Index As Variant) As cPlainfile
    Set Item = m_PrivateCollection.Item(Index)
End Function
Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    ' delegate to the private collection
    Set NewEnum = m_PrivateCollection.[_NewEnum]
End Function
Public Sub Remove(Index As Variant)
    m_PrivateCollection.Remove Index
End Sub
