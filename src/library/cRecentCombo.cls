VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRecentCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Recent As cRecent
Private m_Control As ComboBox
Public Sub Init(oRecent As cRecent, cCombo As ComboBox)
    Set m_Recent = oRecent
    Set m_Control = cCombo
    Me.refresh
End Sub
Public Sub refresh()
End Sub
