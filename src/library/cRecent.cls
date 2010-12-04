VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRecent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' TODO
'
' create cRecentMenu helper class
' and move code out of here
'

' Update is raised each time something changes
Public Event Update()

Private m_Col As Collection
Private m_EnumCol As Collection
Private m_Max As Long
Private Sub m_Add(sPath As String)
Dim oNew As New cFileinfo
    oNew.Path = sPath
    m_Col.Add oNew
End Sub
Public Sub Clear()
Dim i As Long, m As Long
    If m_Col.Count > 0 Then
        m = m_Col.Count
        For i = m To 1 Step -1
            Call m_Col.Remove(i)
        Next i
        RaiseEvent Update
    End If
End Sub
Public Property Get Count() As Long
    Count = m_Col.Count
End Property
Public Function Find(sPath As String) As Long
Dim oFor As cFileinfo
Dim iIndex As Long
    For Each oFor In m_Col
        If oFor.Path = sPath Then
            Exit For
        End If
        iIndex = iIndex + 1
    Next
    Find = m_Col.Count - iIndex
End Function
Public Property Get Item(Index As Long) As cFileinfo
Dim m As Long
    m = m_Col.Count
    If Index > m Then
        Call Err.Raise(vbObjectError + 1001, , "Invalid Index")
        Exit Property
    End If
    Set Item = m_Col.Item(m - Index + 1)
End Property
Public Property Get Max() As Long
    Max = m_Max
End Property
Public Property Let Max(Count As Long)
    If Count < 1 Then Count = 1
    m_Max = Count
End Property
' advanced properties of this function to be set to -4
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Dim i As Long, m As Long

    If m_EnumCol Is Nothing Then
        'instantiate
        Set m_EnumCol = New Collection
    Else
        'clear
        For i = m_EnumCol.Count To 1 Step -1
            m_EnumCol.Remove i
        Next i
    End If

    'create a reverse collection
    m = m_Col.Count
    For i = m To 1 Step -1
        m_EnumCol.Add m_Col.Item(i)
    Next i
    
    'pass new enum of that reverse collection
    Set NewEnum = m_EnumCol.[_NewEnum]
End Property
Private Sub Remove(Index As Long)
    m_Col.Remove (1 + m_Col.Count - Index)
End Sub
' add or put on top of the list by touching it
Public Sub Touch(sPath As String)
Dim iFound As Long
    'check wether or not it already exists in the list
    iFound = Me.Find(sPath)
    If iFound > 0 Then Remove (iFound)
    Call m_Add(sPath)
    If m_Col.Count > m_Max Then
        Call m_Col.Remove(1)
    End If
    RaiseEvent Update
End Sub
Private Sub Class_Initialize()
    Set m_Col = New Collection
    m_Max = 64
End Sub
Private Sub Class_Terminate()
    Set m_Col = Nothing
    Set m_EnumCol = Nothing
End Sub
