VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRecentCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' cRecent Collection - Handle multiple recent objects in an indexed collection
'
Public Event Update(Index As Long)

Private colItems As New Collection


Public Function Add(Optional oNewItem As cRecent = Nothing, Optional Key As String = vbNullString) As cRecent
    If oNewItem Is Nothing Then
        Set oNewItem = New cRecent
    End If
    
    If Key = vbNullString Then
        Call colItems.Add(oNewItem)
    Else
        Call colItems.Add(oNewItem, Key)
    End If
    Set Add = oNewItem
End Function
Public Function AddMultiple(ParamArray NewKeys()) As Collection
Dim i As Long, m As Long
Dim oAdded As New Collection
Dim sKey As String

    m = UBound(NewKeys)
        
    If m > -1 Then
        For i = 0 To m
            sKey = NewKeys(i)
            Call oAdded.Add(Me.Add(, sKey))
        Next
    End If
    
    Set AddMultiple = oAdded
End Function
Public Property Get Count() As Long
    Count = colItems.Count
End Property
Public Property Get Item(Index As Long) As cRecent
    If Index < 1 Or Index > colItems.Count Then
        Call Err.Raise(vbObjectError + 1001, , "Invalid Index")
        Exit Property
    End If
    Set Item = colItems.Item(Index)
End Property
