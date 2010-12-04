VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFbHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' api defines
'

'autocomplete: http://vbnet.mvps.org/code/textapi/shautocomplete.htm
Private Declare Function SHAutoComplete Lib "shlwapi" _
  (ByVal hwndEdit As Long, _
   ByVal dwFlags As Long) As Long
   
'Includes the File System as well as the rest of the shell (Desktop\My Computer\Control Panel\)
Private Const SHACF_FILESYSTEM  As Long = &H1

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
   (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, _
   ByVal lpszWindow As String) As Long

'
' vb stuff
'
Public Event Changed()
Public Event Click(sKey As String)

Dim m_Text As Control
Dim m_Toolbar As toolbar
Private WithEvents m_Recent As cRecent
Attribute m_Recent.VB_VarHelpID = -1

Dim m_OldText As String

'
' Init Class
'
Public Sub Init(cText As Control, cToolbar As toolbar, bRecent As Boolean, Optional bAutocomplete As Boolean = True)
Dim hWnd As Long

    Set m_Text = cText
    
    If bAutocomplete Then
        Select Case TypeName(m_Text)
            Case "TextBox":
                Call SHAutoComplete(m_Text.hWnd, SHACF_FILESYSTEM)
            Case "ComboBox":
                hWnd = FindWindowEx(m_Text.hWnd, 0, "EDIT", vbNullString)
                Call SHAutoComplete(hWnd, SHACF_FILESYSTEM)
            Case Else:
        End Select
    End If
    
    Set m_Toolbar = cToolbar
    
    Me.HasRecent = bRecent

End Sub
Public Property Get HasRecent() As Boolean
    HasRecent = Not CBool(m_Recent Is Nothing)
End Property
Public Property Let HasRecent(bValue As Boolean)
    If bValue Then
        If m_Recent Is Nothing Then
            Set m_Recent = New cRecent
        End If
    Else
        Set m_Recent = Nothing
    End If
End Property
Public Sub Trigger(TheEvent As eAcTriggerEvents)
Dim Ctl As ComboBox
Dim t As String
Dim i As Long

    If TheEvent = ClickEvent Then
        If m_OldText <> m_Text.Text Then
            RaiseEvent Changed
            m_OldText = m_Text.Text
        End If
    End If

    If TheEvent = ChangeEvent Then
        If Not m_Text Is Nothing Then
            If m_OldText <> m_Text.Text Then
                RaiseEvent Changed
                m_OldText = m_Text.Text
            End If
        End If
    End If
End Sub
Public Sub TriggerClick(sKey As String)
    RaiseEvent Click(sKey)
End Sub
Public Property Get Recent() As cRecent
    Set Recent = m_Recent
End Property
Public Sub RecentTouch()
    If Not m_Recent Is Nothing Then
        If Len(m_Text.Text) Then
            m_Recent.Touch m_Text.Text
        End If
    End If
End Sub

Private Sub m_Recent_Update()
Dim oFi As cFileinfo
Dim Ctl As ComboBox
Dim i As Long, m As Long
Dim t As String

    If TypeName(m_Text) = "ComboBox" Then
        'Fill the combobox with items
        Set Ctl = m_Text
        'Ctl.Clear
        m = Ctl.ListCount
        t = Ctl.Text
        If m > 0 Then
            For i = m To 1 Step -1
                If t = Ctl.list(i - 1) Then
                    Call Ctl.RemoveItem(i - 1)
                    Ctl.Text = t
                Else
                    Call Ctl.RemoveItem(i - 1)
                End If
            Next i
        End If
        For Each oFi In m_Recent
            Ctl.AddItem oFi.Path
        Next
        Ctl.Text = t
        Ctl.ListIndex = 0
    End If
    
End Sub
