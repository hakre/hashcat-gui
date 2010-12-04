VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLvPlains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_list As ListView
Private m_helper As cLvHelper
Private m_form As fMain
Public Function cmd(ByVal sCmd As String) As Long
Dim oLi As ListItem
Dim sString As String
Dim sFiles() As String
Dim i As Long, m As Long
Dim cCol As Collection

    cmd = 1
    
    Select Case sCmd
        Case "cb_copy", "cb_cut":
            If m_helper.SelCount > 0 Then
                For Each oLi In m_helper.SelectedItems
                    sString = sString & oLi.SubItems(1) & vbCrLf
                Next
                Clipboard.Clear
                Clipboard.SetText sString
            End If
            If sCmd = "cb_cut" Then
                'FIXME "item_del" is not supported in this cLvPlains right now
            End If
            
        Case "cb_paste"
            If Clipboard.GetFormat(vbCFText) Then
                sString = Clipboard.GetText
                sFiles = Split(sString, vbCrLf)
                m = COL_arrayCountElements(sFiles)
                If m > 0 Then
                    Set cCol = New Collection
                    m = m - 1
                    For i = 0 To m
                        Call cCol.Add(sFiles(i))
                    Next i
                    'Call m_form.plainsList_AddFile
                    Call m_form.plainsList_AddFiles(cCol)
                End If
            End If
            
        Case Else:
            Stop
            cmd = 0
            
    End Select

End Function
Public Property Get Form() As fMain
    Set Form = m_form
End Property
Public Property Get Helper() As cLvHelper
    Set Helper = m_helper
End Property
Public Sub Init(list As ListView, Helper As cLvHelper, Form As fMain)
    Set m_list = list
    Set m_helper = Helper
    Set m_form = Form
End Sub
Public Property Get list() As ListView
    Set list = m_list
End Property

