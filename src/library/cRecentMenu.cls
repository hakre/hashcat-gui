VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRecentMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Recent Menu
'
' Menu Handler for cRecent
'
' Needs a base menu and the submenu
'
' Will disable the basemenu if there
' are no elements
'
'
Option Explicit

Private m_Recent As cRecent
Private m_Menu As Object
Private m_ParentMenu As Object

Public Sub Init(oRecent As cRecent, cMenu As Object, Optional cParentMenu As Object = Nothing)
    Set m_Recent = oRecent
    Set m_Menu = cMenu
    Set m_ParentMenu = cParentMenu
    Me.refresh
End Sub
Public Property Get Menu() As Object
    Set Menu = m_Menu
End Property
Public Property Get Parent() As Object
    Set Parent = m_ParentMenu
End Property
Public Property Get Recent() As cRecent
    Set Recent = m_Recent
End Property
Public Sub refresh()
Dim i As Long
Dim d As Long
Dim oMenu As Menu
Dim oFi As cFileinfo
Dim bFlag As Boolean
Dim sCaption As String

    If m_Menu Is Nothing Then Exit Sub
    bFlag = Not CBool(m_Recent.Count = 0)
    
    If Not m_ParentMenu Is Nothing Then
        m_ParentMenu.Enabled = bFlag
    End If
    
    If m_Menu.Count > m_Recent.Count Then
        'check that all not needed items are hidden
        For i = m_Recent.Count To m_Menu.Count - 1
            If i > 0 Then 'first can not be not visible
                m_Menu.Item(i).Visible = False
            Else
                m_Menu.Item(i).Caption = "Recent List" 'some title to describe what this is
            End If
            m_Menu.Item(i).Enabled = False
        Next i
    End If
    ' need to create some more
    If m_Menu.Count < m_Recent.Count Then
        d = m_Recent.Count - m_Menu.Count
        For i = 1 To d
            Load m_Menu(m_Menu.Count)
        Next i
    End If
    
    i = 0
    For Each oFi In m_Recent
        ' get letter based on number: 1-9, 0, a-z, <none>
        ' 49 - 1; 57 - 9; 48 - 0; 97 - a; 122 - z
        Select Case i
            Case Is < 9:
                sCaption = Chr(49 + i)
            Case Is < 10:
                sCaption = Chr(48)
            Case Is < 36:
                sCaption = Chr(87 + i)
            Case Else
                sCaption = ""
        End Select
        
        If i < 36 Then
            sCaption = "&" & sCaption
        Else
            sCaption = " "
        End If
        sCaption = sCaption & "  " & oFi.Path
        
        Set oMenu = m_Menu.Item(i)
        oMenu.Tag = oFi.Path
        oMenu.Caption = sCaption
        oMenu.Visible = True
        oMenu.Enabled = True
        i = i + 1
    Next
    
End Sub
Private Sub Class_Terminate()
    Set m_Menu = Nothing
    Set m_ParentMenu = Nothing
    Set m_Recent = Nothing
End Sub
