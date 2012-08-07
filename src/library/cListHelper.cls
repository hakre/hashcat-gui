VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLvHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' ListView Helper for plains list
'

Private m_list As ListView
Public Sub CheckAll(Optional ByVal bValue As Boolean = True)
Dim oLi As ListItem

    If m_list Is Nothing Then Exit Sub

    For Each oLi In m_list.ListItems
        oLi.Checked = bValue
    Next
    
End Sub
'
' count of number of checked items
'
Public Property Get CheckCount() As Long
    CheckCount = COL_arrayCountElements(zNumberedToArrayLong(zNumberedGet(1)))
End Property
Public Sub CheckInverse()
Dim oLi As ListItem

    For Each oLi In m_list.ListItems
        oLi.Checked = Not oLi.Checked
    Next
    
End Sub

Public Sub CheckNumbered(ByVal sNumbered As String, bCheck As Boolean)
Dim oLi As ListItem

    For Each oLi In zNumberedItems(sNumbered)
        oLi.Checked = bCheck
    Next
    
End Sub


'
' copy one item's data to another
'
' NOTE: this does not copy over the key
' NOTE: missing images are not caught and will result in runtime errors
'
Public Function CopyItemData(oFrom As ListItem, oTo As ListItem) As Long
    CopyItemData = zListItemCopyData(oFrom, oTo)
End Function

'
' copy a whole list to another
'
Public Sub CopyList(toList As ListView)
Dim oLi As ListItem
Dim oNew As ListItem

    toList.ListItems.Clear
    
    For Each oLi In m_list.ListItems
        Set oNew = toList.ListItems.Add()
        Call zListItemCopyData(oLi, oNew)
        oNew.Key = oLi.Key
    Next

End Sub





Public Property Get list() As ListView
    Set list = m_list
End Property
Public Property Set list(cListView As ListView)
    Set m_list = cListView
End Property
Public Function MoveNumbered(ByVal sNumbered As String, ByVal iDirection As Long) As Long
Dim aIndexies() As Long
Dim i As Long, m As Long
Dim iIndex As Long
Dim oTemp As ListItem
Dim oNew As ListItem
Dim iDirValue As Long
Dim iForFrom As Long, iForTo As Long


    aIndexies = zNumberedToArrayLong(sNumbered)
    m = COL_arrayCountElements(aIndexies)
    
    'there aren't any elements, so there is nothing to move
    If m < 1 Then Exit Function
        
    'smallest element is first, so no way to move up (up = to the begging)
    If aIndexies(m - 1) = 1 And iDirection = 0 Then Exit Function
    
    'biggest element is last, so no way to move down (down = to the end)
    If aIndexies(0) = m_list.ListItems.Count And iDirection = 1 Then Exit Function
    
    m = m - 1
    
    'move
    iDirValue = -1 + iDirection * 2
    If iDirection = 1 Then
        'iDirValue = 1
        iForFrom = 0
        iForTo = m
    Else
        'iDirValue = -1
        iForFrom = m
        iForTo = 0
    End If
    
    For i = iForFrom To iForTo Step iDirValue
        iIndex = aIndexies(i)
        Set oTemp = m_list.ListItems.Item(iIndex)
        Set oNew = m_list.ListItems.Add(iIndex + iDirValue + iDirection)
        Call zListItemCopyData(oTemp, oNew)
        Call m_list.ListItems.Remove(iIndex - iDirValue * (1 - iDirection))
        oNew.Key = oTemp.Key
    Next i
    
    m_list.ListItems.Item(aIndexies((1 - iDirection) * m) + iDirValue).EnsureVisible
        
MoveNumbered = 1
End Function

Public Function MoveNumberedMultiple(ByVal sNumbered As String, ByVal iDirection As Long, ByVal iCount As Long) As Long
Dim aIndexies() As Long
Dim r As Long
Dim iDirValue As Long
Dim iMoved As Long
Dim i As Long
Dim j As Long, m As Long


    iDirValue = -1 + iDirection * 2
    
    For i = 1 To iCount Step 1
        'move items
        r = MoveNumbered(sNumbered, iDirection)
        If r > 0 Then
            iMoved = iMoved + 1
        Else
            Exit For
        End If
        
        'shift numbered
        If i = 1 Then
            aIndexies = zNumberedToArrayLong(sNumbered)
            m = UBound(aIndexies)
        End If
        
        For j = 0 To m
            aIndexies(j) = aIndexies(j) + iDirValue
        Next j
        sNumbered = zNumberedFromArrayLong(aIndexies)
        
    Next i
    
MoveNumberedMultiple = iMoved
End Function

'retruns: collection of listitems
Public Property Get SelectedItems() As Collection
Dim colItems As New Collection
Dim oLi As ListItem

    For Each oLi In m_list.ListItems
        If oLi.Selected Then
            colItems.Add oLi
        End If
    Next
    
Set SelectedItems = colItems
End Property
'
' convert array of long values to a sNumbered string
Private Function zNumberedFromArrayLong(aEntries() As Long) As String
Dim m As Long, i As Long
Dim sNumbered As String

    m = COL_arrayCountElements(aEntries)
    
    If m > 0 Then
        m = m - 1
        For i = 0 To m Step 1
            sNumbered = sNumbered & "," & CStr(aEntries(i))
        Next
        sNumbered = Mid(sNumbered, 2)
    End If
    
zNumberedFromArrayLong = sNumbered
End Function


Public Property Get NumberedChecked() As String
    NumberedChecked = zNumberedGet(1)
End Property

Public Property Get NumberedSelected() As String
    NumberedSelected = zNumberedGet(0)
End Property




'
' remove by index (numbered in string, sorted descending)
'
Public Function RemoveNumbered(sNumbers As String) As Long
Dim Indexies() As Long
Dim i As Long, m As Long
Dim iCount As Long

    Indexies = zNumberedToArrayLong(sNumbers)
    
    m = UBound(Indexies)
    
    If m > -1 Then
        For i = 0 To m Step 1
            iCount = iCount + 1
            Call m_list.ListItems.Remove(Indexies(i))
        Next i
    End If
    
RemoveNumbered = iCount
End Function
Public Function RemoveSelected() As Long
Dim oLi As ListItem
Dim iCount As Long

retest:
        For Each oLi In m_list.ListItems
            If oLi.Selected Then
                iCount = iCount + 1
                Call m_list.ListItems.Remove(oLi.Index)
                GoTo retest
            End If
        Next
        
RemoveSelected = iCount
End Function
Public Property Get SelCount() As Long
    SelCount = COL_arrayCountElements(zNumberedToArrayLong(zNumberedGet(0)))
End Property
Public Sub SelectAll(Optional bSelect As Boolean = True)
Dim oLi As ListItem

    For Each oLi In m_list.ListItems
        oLi.Selected = bSelect
    Next

End Sub
Public Sub SelectInverse()
Dim oLi As ListItem

    For Each oLi In m_list.ListItems
        oLi.Selected = Not oLi.Selected
    Next

End Sub


'
Private Function zListItemCopyData(oFrom As ListItem, oTo As ListItem) As Long
Dim i As Long, m As Long
    
    'checked
    oTo.Checked = oFrom.Checked
    
    'icon
    oTo.Icon = oFrom.Icon
    oTo.SmallIcon = oFrom.SmallIcon

    'selected
    oTo.Selected = oFrom.Selected

    'subitems
    m = oFrom.ListSubItems.Count
    If m > 0 Then
        For i = 1 To m Step 1
            oTo.SubItems(i) = oFrom.SubItems(i)
        Next i
    End If
    
    'text
    oTo.Text = oFrom.Text

    'tooltip
    oTo.ToolTipText = oFrom.ToolTipText
    
zListItemCopyData = 1
End Function


Private Function zNumberedGet(Optional ByVal iType As Long = 0) As String
Dim sNumbered As String
Dim oLi As ListItem
Dim bFlag As Boolean

    
    For Each oLi In m_list.ListItems
        bFlag = ((iType = 0 And oLi.Selected) Or (iType = 1 And oLi.Checked))
        If bFlag Then
            sNumbered = "," & CStr(oLi.Index) & sNumbered
        End If
    Next
    
    If Len(sNumbered) > 1 Then
        sNumbered = Mid(sNumbered, 2)
    End If

zNumberedGet = sNumbered
End Function
'return: collection of listitems specified by sNumbered
Private Function zNumberedItems(ByVal sNumbered As String) As Collection
Dim colItems As New Collection
Dim iIndexies() As Long
Dim i As Long, m As Long

    iIndexies = zNumberedToArrayLong(sNumbered)
    m = COL_arrayCountElements(iIndexies)
    If m > 0 Then
        m = m - 1
        For i = 0 To m
            colItems.Add m_list.ListItems.Item(iIndexies(i))
        Next i
    End If

Set zNumberedItems = colItems
End Function


' convert a sNumbered string (comma seperated values) to an array of long values
' preventing invalid values (like 0)
Private Function zNumberedToArrayLong(sNumbered As String) As Long()
    zNumberedToArrayLong = COL_arrayNumbered(sNumbered)
End Function
