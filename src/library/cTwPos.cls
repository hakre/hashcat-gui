VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTwPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' Toolwindow Position / Size Helper
'

Private m_PosLeft As Long
Private m_PosTop As Long

Private m_form As Form
Private m_StoreRelative As cTwPos
Public Function Init(Toolwindow As Form) As cTwPos
    Set m_form = Toolwindow
    Set Init = Me
End Function
Public Property Get isLoaded() As Boolean
    isLoaded = HCGUI_Form_IsLoaded(m_form)
End Property
Public Property Get Left() As Long
    Left = m_PosLeft
End Property
Public Property Let Left(iValue As Long)
    m_PosLeft = iValue
End Property
Public Sub Restore()
    If Me.isLoaded Then
        m_form.Move m_PosLeft, m_PosTop
    End If
End Sub
Public Sub RestoreRelativeTo(theForm As Form, Optional RelativePos As cTwPos = Nothing)
    If Me.isLoaded Then
        If RelativePos Is Nothing Then
            Set RelativePos = m_StoreRelative
        End If
        If Not RelativePos Is Nothing Then
            m_form.Move RelativePos.Left + theForm.Left, RelativePos.Top + theForm.Top
        End If
    End If
End Sub
Public Function RelativeTo(theForm As Form) As cTwPos
Dim newPos As cTwPos
    If Me.isLoaded Then
        Set newPos = New cTwPos
        newPos.Left = m_form.Left - theForm.Left
        newPos.Top = m_form.Top - theForm.Top
        Set m_StoreRelative = newPos
        Set RelativeTo = newPos
    End If
End Function
Public Sub Store()
    m_PosLeft = m_form.Left
    m_PosTop = m_form.Top
End Sub
Public Property Get Top() As Long
    Top = m_PosTop
End Property
Public Property Let Top(iValue As Long)
    m_PosTop = iValue
End Property
Public Property Get Window() As Form
    Set Window = m_form
End Property

