VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTwHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' Toolwindow Helper
'


Private m_Main As fMain
Private m_Toolwindow As Form
Private m_ShowedOnce As Boolean
Private m_Pos As cTwPos

Public Sub Init(MainForm As fMain, Toolwindow As Form)
    ' private members
    Set m_Main = MainForm
    Set m_Toolwindow = Toolwindow
    
    ' pos private member handling
    If m_Pos Is Nothing Then
        Set m_Pos = New cTwPos
    End If
    Call m_Pos.Init(Toolwindow)
    
    
End Sub
Public Property Get isLoaded() As Boolean
    isLoaded = HCGUI_Form_IsLoaded(m_Toolwindow)
End Property
Public Sub Load()
    'simulate loading by setting something
    m_Toolwindow.Caption = m_Toolwindow.Caption
End Sub
Public Property Get Main() As fMain
    Set Main = m_Main
End Property
Public Sub Toggle()
Dim oPos As cTwPos

    If Me.Visible = False Then
        If m_ShowedOnce Then
            Call m_Pos.RestoreRelativeTo(m_Main)
            
            'prevent the move out of the screen
            Call POS_window_forceonscreen(m_Toolwindow)
            
            'store that value
            Set oPos = New cTwPos
            oPos.Init(m_Toolwindow).Store
            m_Toolwindow.Visible = True
            oPos.Restore 'recall if window has got a save windowpos
        Else
            m_ShowedOnce = True
            Set m_Toolwindow.Parent = m_Main
            m_Toolwindow.Show False, m_Main
        End If
    Else
        Call m_Pos.RelativeTo(m_Main)
        m_Toolwindow.Visible = False
    End If
    
End Sub
Public Property Get Toolwindow() As Form
    Set Toolwindow = m_Toolwindow
End Property
Public Property Get Visible() As Boolean
    If Me.isLoaded Then
        Visible = Me.Toolwindow.Visible
    End If
End Property

